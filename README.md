# SheetGuard OSS

**Read-only, local-first release QA for Excel `.xlsx` workbooks.**

SheetGuard finds common structural regressions in OOXML spreadsheets before release. It inspects formula element counts, data validation rules and worksheet protection, compares workbooks against a previous baseline, and outputs machine-readable reports suitable for CI.

**Status:** v0.1.1 alpha · Python 3.10+ · MIT · zero runtime dependencies.

## Why this exists

Some spreadsheet editing and conversion tools can remove worksheet protection or validation while leaving the workbook apparently usable. SheetGuard offers an automated way to detect these specific regressions without opening the file in Microsoft Excel.

## Installation

```bash
git clone https://github.com/ayoubchmali2006-del/sheetguard-oss.git
cd sheetguard-oss
python -m pip install .
```

Not published to PyPI. `pip install sheetguard-oss` is not yet supported.

## Quick start

Create a synthetic workbook example:

```bash
python examples/make_demo.py
sheetguard examples/demo.xlsx --require-protection
sheetguard examples/demo.xlsx --format json
```

Check real workbooks locally:

```bash
sheetguard workbook.xlsx
sheetguard after.xlsx --baseline before.xlsx --fail-on warning
sheetguard after.xlsx --baseline before.xlsx --format json --output report.json
```

`--require-protection` enforces the presence of a worksheet protection XML element. Protection is not encryption.

### Python API

```python
from sheetguard_oss import inspect, compare
before = inspect("before.xlsx")
after = inspect("after.xlsx")
changes = compare(before, after)
for finding in changes:
    print(finding)
```

## What gets checked

- Broken or missing worksheet relationships and worksheet parts
- Missing worksheet protection when required by policy
- Loss of protection, validation rules and formula elements relative to a baseline
- Formula XML containing `#REF!` and cached Excel error cells
- Missing sheets, visibility changes and external link parts
- Some archive-integrity issues, suspicious ZIP member names and XML DTD/entity declarations

The tool only reads input. No workbook editing, uploads or Microsoft Excel installation required.

## CLI exit codes

| Exit code | Meaning |
| --- | --- |
| 0 | No findings exceeding selected threshold |
| 1 | A policy warning/error or regression exceeded the threshold |
| 2 | Invalid/unreadable workbook or report output failure |

Use `--fail-on error` (default), `--fail-on warning`, or `--fail-on none`. Structural inspection errors return 2.

## Tests

```bash
python -m unittest discover -s tests -v
```

GitHub Actions runs tests across Python 3.10–3.13 for pushes and pull requests. Always verify the actual run status in the Actions tab; workflow presence alone is not proof of passing tests.

## Limitations

- Does **not** recalculate formulas or guarantee workbook/business correctness.
- Does **not** certify workbook security; worksheet protection is not encryption.
- Does **not** fully validate every OOXML structure.
- Does not support `.xls`, `.xlsm`, Google Sheets or encrypted workbooks.
- XML and ZIP checks reduce common risks but are not a malware scan or security certification.
- A clean report only covers checks implemented in this alpha release.

## Roadmap

- Per-worksheet protection and validation policy configuration
- Versioned baseline manifests and trend reports
- Better compatibility fixtures and stronger OOXML relationship checks
- Community feedback and independent usage evidence

## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md) for tests and pull requests, [SECURITY.md](SECURITY.md) for reporting vulnerabilities and [LICENSE](LICENSE) for terms. Do not share real customer spreadsheets in public issues.

Maintained as an independent open-source utility.