# SheetGuard OSS

**Read-only, local-first release QA for Excel `.xlsx` templates.**

SheetGuard detects structural regressions in Excel workbooks before release. It can compare new and baseline versions and identify missing worksheet protection, removed validations, missing sheets, reductions in formula counts, and some package-integrity issues.

**Status:** v0.1.0 alpha. Source archive is being prepared for publication. Do not treat this repository as a working installation until `src/`, `tests/` and `pyproject.toml` are committed.

## Main goals

- Inspect OOXML workbook and worksheet relationships.
- Report formulas, validation rules, and protection flags.
- Detect regressions against a baseline workbook.
- Provide a deterministic, local and read-only command-line quality gate.
- Run without Microsoft Excel or proprietary services.

## Safety and limitations

SheetGuard does not recalculate formulas, validate business logic, prove workbook correctness, or provide a security certification. Excel worksheet protection is not encryption.

## Installation (after the source files are committed)

```bash
python -m pip install .
sheetguard workbook.xlsx --require-protection
sheetguard new.xlsx --baseline previous.xlsx --fail-on warning
```

## License

MIT. Please see LICENSE in the forthcoming source upload.

## Contributing

Issues and reproducible bug reports are welcome.