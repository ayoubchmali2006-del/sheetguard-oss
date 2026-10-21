"""Regression checks for OOXML edge cases using synthetic workbooks."""
import tempfile
import unittest
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from sheetguard_oss import compare, inspect
from test_sheetguard import generate


def rewrite(source, destination, part, old, new):
    with ZipFile(source) as zin, ZipFile(destination, "w", ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == part:
                data = data.replace(old.encode(), new.encode())
            zout.writestr(item.filename, data)
    return destination


class RegressionEdgeCases(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)

    def book(self, name="book.xlsx", **kwargs):
        return generate(self.root / name, **kwargs)

    def change(self, source, part, old, new, name="changed.xlsx"):
        return rewrite(source, self.root / name, part, old, new)

    def codes(self, report):
        return {f.code for f in report.findings}

    def test_disabled_protection_is_not_protected(self):
        baseline = self.book("base.xlsx")
        changed = self.change(baseline, "xl/worksheets/sheet1.xml",
                              'sheet="1"', 'sheet="0"')
        report = inspect(changed, require_protection=True)
        self.assertFalse(report.sheets[0].protected)
        self.assertIn("PROTECTION_REQUIRED", self.codes(report))

    def test_same_count_smaller_validation_scope_detected(self):
        original = self.book("base.xlsx")
        baseline = inspect(original)
        changed = self.change(original, "xl/worksheets/sheet1.xml",
                              'sqref="A1:A200"', 'sqref="A1:A10"')
        self.assertIn("VALIDATION_SCOPE_CHANGED",
                      {f.code for f in compare(baseline, inspect(changed))})

    def test_same_count_moved_formula_detected(self):
        original = self.book("base.xlsx")
        baseline = inspect(original)
        changed = self.change(original, "xl/worksheets/sheet1.xml",
                              '<c r="A1">', '<c r="D7">')
        self.assertIn("FORMULA_CELLS_CHANGED",
                      {f.code for f in compare(baseline, inspect(changed))})

    def test_invalid_workbook_root_detected(self):
        original = self.book("base.xlsx")
        changed = self.change(original, "xl/workbook.xml",
                              '<workbook ', '<notworkbook ')
        changed2 = self.change(changed, "xl/workbook.xml",
                               '</workbook>', '</notworkbook>', "changed2.xlsx")
        self.assertIn("BAD_WORKBOOK_ROOT", self.codes(inspect(changed2)))

    def test_invalid_worksheet_root_detected(self):
        original = self.book("base.xlsx")
        changed = self.change(original, "xl/worksheets/sheet1.xml",
                              '<worksheet ', '<notworksheet ')
        changed2 = self.change(changed, "xl/worksheets/sheet1.xml",
                               '</worksheet>', '</notworksheet>', "changed2.xlsx")
        self.assertIn("BAD_SHEET_ROOT", self.codes(inspect(changed2)))

    def test_duplicate_sheet_names_case_insensitive(self):
        original = self.book("base.xlsx")
        changed = self.change(original, "xl/workbook.xml", '</sheets>',
                              '<sheet name="contacts" sheetId="2" r:id="rId1"/></sheets>')
        self.assertIn("DUPLICATE_SHEET_NAME", self.codes(inspect(changed)))

    def test_empty_validation_scope_warning(self):
        original = self.book("base.xlsx")
        changed = self.change(original, "xl/worksheets/sheet1.xml",
                              'sqref="A1:A200"', 'sqref=""')
        self.assertIn("EMPTY_VALIDATION_SCOPE", self.codes(inspect(changed)))


if __name__ == "__main__":
    unittest.main()
