"""Unit tests using generated, non-sensitive .xlsx fixtures."""
import io
import json
import tempfile
import unittest
from contextlib import redirect_stdout
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from sheetguard_oss import compare, inspect
from sheetguard_oss.cli import main
from sheetguard_oss.core import _safe_sheet_path, _xml

S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
P = "http://schemas.openxmlformats.org/package/2006/relationships"


def generate(path, protect=True, formulas=1, validations=1, broken=False,
             missing_sheet=False, hidden=False, bad_path=False, dtd=False):
    wb = f'<workbook xmlns="{S}" xmlns:r="{R}"><sheets><sheet name="Contacts" sheetId="1" r:id="rId1" state="{"hidden" if hidden else "visible"}"/></sheets></workbook>'
    rels = f'<Relationships xmlns="{P}"><Relationship Id="rId1" Type="{R}/worksheet" Target="{"../../escape.xml" if bad_path else "worksheets/sheet1.xml"}"/></Relationships>'
    fs = "".join(f'<c r="A{i+1}"><f>{"#REF!+1" if broken else "1+2"}</f><v>3</v></c>' for i in range(formulas))
    rules = "".join('<dataValidation type="list" sqref="A1:A200"><formula1>"a,b"</formula1></dataValidation>' for _ in range(validations))
    protection = '<sheetProtection sheet="1"/>' if protect else ''
    sheet = f'<worksheet xmlns="{S}">{protection}<sheetData><row r="1">{fs}</row></sheetData><dataValidations count="{validations}">{rules}</dataValidations></worksheet>'
    if dtd:
        sheet = '<!DOCTYPE worksheet [<!ENTITY x "unsafe">]>' + sheet
    with ZipFile(path, "w", ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>')
        z.writestr("_rels/.rels", f'<Relationships xmlns="{P}"><Relationship Id="document" Type="{R}/officeDocument" Target="/xl/workbook.xml"/></Relationships>')
        z.writestr("xl/workbook.xml", wb)
        z.writestr("xl/_rels/workbook.xml.rels", rels)
        if not missing_sheet:
            z.writestr("xl/worksheets/sheet1.xml", sheet)
    return path


class InspectionTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)

    def book(self, name="test.xlsx", **kwargs):
        return generate(self.root / name, **kwargs)

    def codes(self, report):
        return {f.code for f in report.findings}

    def test_valid(self):
        r = inspect(self.book())
        self.assertTrue(r.valid)
        self.assertEqual(r.total_formulas, 1)
        self.assertEqual(r.total_validations, 1)

    def test_read_only(self):
        p = self.book()
        data = p.read_bytes()
        inspect(p)
        self.assertEqual(data, p.read_bytes())

    def test_default_allows_no_protection(self):
        self.assertTrue(inspect(self.book(protect=False)).valid)

    def test_requires_protection(self):
        r = inspect(self.book(protect=False), require_protection=True)
        self.assertFalse(r.valid)
        self.assertIn("PROTECTION_REQUIRED", self.codes(r))

    def test_broken_formula(self):
        self.assertIn("BROKEN_FORMULA_REF", self.codes(inspect(self.book(broken=True))))

    def test_hidden(self):
        self.assertIn("HIDDEN_SHEET", self.codes(inspect(self.book(hidden=True))))

    def test_missing_sheet(self):
        self.assertIn("MISSING_SHEET", self.codes(inspect(self.book(missing_sheet=True))))

    def test_escaped_target(self):
        self.assertIn("BAD_REL_TARGET", self.codes(inspect(self.book(bad_path=True))))

    def test_dtd_rejected(self):
        self.assertIn("MALFORMED_SHEET", self.codes(inspect(self.book(dtd=True))))

    def test_corrupt_archive(self):
        p = self.root / "corrupt.xlsx"
        p.write_bytes(b"not a zip")
        self.assertIn("INVALID_WORKBOOK", self.codes(inspect(p)))

    def test_missing_file(self):
        self.assertIn("IO", self.codes(inspect(self.root / "missing.xlsx")))

    def test_non_xlsx(self):
        self.assertIn("FORMAT", self.codes(inspect(self.root / "wrong.xls")))

    def test_baseline_protection(self):
        a = inspect(self.book("a.xlsx"))
        b = inspect(self.book("b.xlsx", protect=False))
        self.assertIn("PROTECTION_LOST", {f.code for f in compare(a, b)})

    def test_baseline_counts(self):
        a = inspect(self.book("a.xlsx", formulas=3, validations=2))
        b = inspect(self.book("b.xlsx", formulas=1, validations=0))
        codes = {f.code for f in compare(a, b)}
        self.assertIn("FORMULAS_REDUCED", codes)
        self.assertIn("VALIDATIONS_REDUCED", codes)

    def test_invalid_comparison(self):
        a = inspect(self.book("a.xlsx"))
        b = inspect(self.book("b.xlsx", broken=True))
        self.assertIn("COMPARE_INVALID", {f.code for f in compare(a, b)})

    def test_cli_success(self):
        p = self.book()
        with redirect_stdout(io.StringIO()) as output:
            result = main([str(p), "--format", "json"])
        self.assertEqual(result, 0)
        self.assertEqual(json.loads(output.getvalue())["summary"]["sheets"], 1)

    def test_cli_failed_policy(self):
        p = self.book(protect=False)
        with redirect_stdout(io.StringIO()):
            self.assertEqual(main([str(p), "--require-protection"]), 2)

    def test_cli_baseline_error(self):
        a = self.book("old.xlsx")
        b = self.book("new.xlsx", protect=False)
        with redirect_stdout(io.StringIO()):
            self.assertEqual(main([str(b), "--baseline", str(a)]), 1)

    def test_cli_report_output(self):
        p = self.book()
        report = self.root / "output" / "report.json"
        self.assertEqual(main([str(p), "--format", "json", "--output", str(report)]), 0)
        self.assertEqual(json.loads(report.read_text())["summary"]["formulas"], 1)

    def test_json_serialization(self):
        self.assertEqual(json.loads(json.dumps(inspect(self.book()).as_dict()))["schema_version"], 1)

    def test_path_safety(self):
        self.assertEqual(_safe_sheet_path("worksheets/sheet1.xml"), "xl/worksheets/sheet1.xml")
        with self.assertRaises(ValueError):
            _safe_sheet_path("../../../outside.xml")

    def test_xml_dtd(self):
        with self.assertRaises(ValueError):
            _xml(b'<!DOCTYPE a><a/>', 'unsafe.xml')


if __name__ == "__main__":
    unittest.main()
