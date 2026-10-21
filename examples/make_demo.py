"""Create synthetic OOXML workbook, never real customer data."""
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED

out = Path(__file__).with_name("demo.xlsx")
with ZipFile(out, "w", ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>')
    z.writestr("_rels/.rels", '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdRoot" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="/xl/workbook.xml"/></Relationships>')
    z.writestr("xl/workbook.xml", '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Demo" sheetId="1" r:id="rId1"/></sheets></workbook>')
    z.writestr("xl/_rels/workbook.xml.rels", '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>')
    z.writestr("xl/worksheets/sheet1.xml", '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetProtection sheet="1"/><sheetData><row r="1"><c r="A1"><f>1+2</f><v>3</v></c></row></sheetData><dataValidations count="1"><dataValidation type="list" sqref="B1:B10"><formula1>"Yes,No"</formula1></dataValidation></dataValidations></worksheet>')
print(f"Synthetic fixture created: {out}")
