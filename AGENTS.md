# Contributor instructions

- Preserve read-only behavior: never modify or overwrite an input workbook.
- Keep the runtime standard-library only unless strictly justified.
- Support Python 3.10+.
- Add regression tests for defect fixes and changes to OOXML relationship handling.
- Do not claim that XML structure checks validate formulas, business logic, encryption or workbook security.
- Do not use real customer/proprietary spreadsheets in public tests or documentation.
- Run `python -m unittest discover -s tests -v` before merging.
