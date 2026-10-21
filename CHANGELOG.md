# Changelog

## 0.1.1 (2026-10-21)
- Correct sheet protection detection when sheetProtection exists with sheet=0.
- Compare validation target ranges and formula cell locations even when element counts are unchanged.
- Reject duplicate sheet names and unexpected workbook/worksheet XML roots.
- Add synthetic edge-case regression tests and local QA.

## 0.1.0 (2026-10-21)
- Initial alpha implementation of read-only OOXML worksheet inspections.
- Optional protection policy and baseline regression comparison.
- CLI with human-readable and JSON reporting and deterministic exit codes.
- Unit tests and GitHub Actions matrix.
- MIT licensing and contributor/security documentation.
