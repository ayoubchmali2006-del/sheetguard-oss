"""Command-line interface for SheetGuard OSS."""
from __future__ import annotations
import argparse
import json
from pathlib import Path
from .core import SEVERITIES, VERSION, compare, inspect


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="sheetguard", description="Read-only .xlsx release QA: inspect or compare workbook structure and policy.")
    parser.add_argument("workbook", type=Path, help="Current .xlsx file")
    parser.add_argument("--baseline", type=Path, help="Original .xlsx to compare against")
    parser.add_argument("--require-protection", action="store_true", help="Require sheet protection on every worksheet")
    parser.add_argument("--format", choices=("text", "json"), default="text")
    parser.add_argument("--output", type=Path, help="Write the report to a local file")
    parser.add_argument("--fail-on", choices=("none", "warning", "error"), default="error")
    parser.add_argument("--version", action="version", version=VERSION)
    args = parser.parse_args(argv)

    if args.output:
        protected_inputs = {args.workbook.resolve()}
        if args.baseline:
            protected_inputs.add(args.baseline.resolve())
        if args.output.resolve() in protected_inputs:
            parser.exit(2, "Refusing to overwrite an input workbook with the report\n")

    current = inspect(args.workbook, require_protection=args.require_protection)
    baseline = inspect(args.baseline) if args.baseline else None
    if baseline:
        current.findings.extend(compare(baseline, current))
    data = current.as_dict()
    if baseline:
        data["baseline"] = str(args.baseline)
        data["baseline_valid"] = baseline.valid
        data["summary"]["errors"] = sum(f.severity == "error" for f in current.findings)
        data["summary"]["warnings"] = sum(f.severity == "warning" for f in current.findings)
    if args.format == "json":
        rendered = json.dumps(data, indent=2, sort_keys=True, ensure_ascii=False) + "\n"
    else:
        lines = [f"SheetGuard OSS {VERSION} | {args.workbook}",
                 f"Inspected sheets: {len(current.sheets)} | formulas: {current.total_formulas} | data validations: {current.total_validations}"]
        lines += [f"[{f.severity.upper()}] {f.code} ({f.sheet or 'workbook'}): {f.message}" for f in current.findings]
        if not current.findings:
            lines.append("No findings in supported checks (not a formula correctness guarantee).")
        rendered = "\n".join(lines) + "\n"
    if args.output:
        try:
            args.output.parent.mkdir(parents=True, exist_ok=True)
            args.output.write_text(rendered, encoding="utf-8")
        except OSError as exc:
            parser.exit(2, f"Cannot write report: {exc}\n")
    else:
        print(rendered, end="")
    if not current.valid or (baseline is not None and not baseline.valid):
        return 2
    if args.fail_on != "none" and any(SEVERITIES[f.severity] >= SEVERITIES[args.fail_on] for f in current.findings):
        return 1
    return 0
