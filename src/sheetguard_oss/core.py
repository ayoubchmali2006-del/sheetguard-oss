"""Read-only OOXML workbook release-policy and regression inspection.

No Excel calculation engine is used. All input data stays local.
"""
from __future__ import annotations

from dataclasses import asdict, dataclass, field
from pathlib import Path
import posixpath
import re
from zipfile import BadZipFile, ZipFile
import xml.etree.ElementTree as ET

VERSION = "0.1.0"
S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
P = "http://schemas.openxmlformats.org/package/2006/relationships"
STRICT_S = "http://purl.oclc.org/ooxml/spreadsheetml/main"
STRICT_R = "http://purl.oclc.org/ooxml/officeDocument/relationships"
DEFAULT_TOTAL_LIMIT = 256 * 1024 * 1024
DEFAULT_PART_LIMIT = 64 * 1024 * 1024
FORMULA_ERROR = re.compile(r"#REF!", re.IGNORECASE)
SEVERITIES = {"info": 0, "warning": 1, "error": 2}


@dataclass(frozen=True)
class Finding:
    severity: str
    code: str
    message: str
    sheet: str | None = None


@dataclass(frozen=True)
class Sheet:
    name: str
    path: str
    state: str
    protected: bool
    validations: int
    formulas: int
    cached_errors: int


@dataclass
class Report:
    file: str
    sheets: list[Sheet] = field(default_factory=list)
    findings: list[Finding] = field(default_factory=list)
    total_formulas: int = 0
    total_validations: int = 0
    valid: bool = False

    def add(self, severity: str, code: str, message: str, sheet: str | None = None) -> None:
        self.findings.append(Finding(severity, code, message, sheet))

    def as_dict(self) -> dict:
        data = asdict(self)
        data["schema_version"] = 1
        data["tool_version"] = VERSION
        data["summary"] = {
            "sheets": len(self.sheets),
            "formulas": self.total_formulas,
            "validations": self.total_validations,
            "errors": sum(f.severity == "error" for f in self.findings),
            "warnings": sum(f.severity == "warning" for f in self.findings),
        }
        return data


def _name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _xml(raw: bytes, part: str) -> ET.Element:
    """Reject DTD/entity declarations before using stdlib XML parsing."""
    if re.search(br"<!\s*(?:DOCTYPE|ENTITY)\b", raw, flags=re.I):
        raise ValueError(f"XML DTD/entity declaration forbidden in {part}")
    try:
        return ET.fromstring(raw)
    except ET.ParseError as exc:
        raise ValueError(f"Malformed XML in {part}: {exc}") from exc


def _part(zf: ZipFile, path: str, max_part: int) -> ET.Element:
    member = zf.getinfo(path)
    if member.file_size > max_part:
        raise ValueError(f"Part exceeds size limit: {path}")
    return _xml(zf.read(path), path)


def _relationships(zf: ZipFile, max_part: int) -> dict[str, tuple[str, str, str]]:
    if "xl/_rels/workbook.xml.rels" not in zf.namelist():
        raise ValueError("Missing workbook relationships: xl/_rels/workbook.xml.rels")
    root = _part(zf, "xl/_rels/workbook.xml.rels", max_part)
    out: dict[str, tuple[str, str, str]] = {}
    for element in root:
        if _name(element.tag) != "Relationship":
            continue
        rid = element.get("Id", "")
        out[rid] = (
            element.get("Target", ""),
            element.get("TargetMode", "Internal"),
            element.get("Type", ""),
        )
    return out


def _safe_sheet_path(target: str) -> str:
    if not target or "\\" in target or "?" in target or "#" in target:
        raise ValueError("Invalid worksheet relationship target")
    path = (
        posixpath.normpath(target.lstrip("/"))
        if target.startswith("/")
        else posixpath.normpath(posixpath.join("xl", target))
    )
    if path.startswith("../") or path == ".." or not path.startswith("xl/"):
        raise ValueError("Worksheet relationship escapes xl/")
    return path


def inspect(path: str | Path, *, require_protection: bool = False,
            max_total: int = DEFAULT_TOTAL_LIMIT,
            max_part: int = DEFAULT_PART_LIMIT) -> Report:
    filename = Path(path)
    report = Report(file=str(filename))
    if filename.suffix.lower() != ".xlsx":
        report.add("error", "FORMAT", "Only .xlsx workbooks are supported")
        return report
    try:
        size = filename.stat().st_size
        if size > max_total:
            report.add("error", "ARCHIVE_SIZE", "Compressed workbook exceeds configured size limit")
            return report
        with ZipFile(filename) as zf:
            infos = zf.infolist()
            paths = [i.filename for i in infos]
            if len(set(paths)) != len(paths):
                report.add("error", "DUPLICATE_PART", "Archive contains duplicate ZIP member names")
                return report
            for info in infos:
                normalized = info.filename.replace("\\", "/")
                if (info.filename.startswith("/") or "\\" in info.filename
                    or any(seg == ".." for seg in normalized.split("/"))):
                    report.add("error", "UNSAFE_PATH", f"Unsafe member path: {info.filename}")
                    return report
                if (info.external_attr >> 16) & 0o170000 == 0o120000:
                    report.add("error", "SYMLINK", f"Symbolic link ZIP member: {info.filename}")
                    return report
            if sum(i.file_size for i in infos) > max_total:
                report.add("error", "DECOMPRESS_LIMIT", "Uncompressed archive exceeds configured size limit")
                return report
            if any(i.file_size > max_part for i in infos):
                report.add("error", "PART_SIZE", "One or more archive parts exceed size limit")
                return report
            if zf.testzip() is not None:
                report.add("error", "CRC", "Archive has a member with invalid CRC")
                return report
            members = set(paths)
            for required in ("[Content_Types].xml", "_rels/.rels", "xl/workbook.xml"):
                if required not in members:
                    report.add("error", "MISSING_PART", f"Required part not found: {required}")
                    return report
            ct = _part(zf, "[Content_Types].xml", max_part)
            if _name(ct.tag) != "Types":
                report.add("error", "BAD_CONTENT_TYPES", "Content types part does not contain a Types root")
                return report
            package_rels = _part(zf, "_rels/.rels", max_part)
            doc_targets = []
            for rel in package_rels:
                if _name(rel.tag) != "Relationship" or not rel.get("Type", "").endswith("/officeDocument"):
                    continue
                if rel.get("TargetMode", "Internal").lower() == "external":
                    report.add("error", "EXTERNAL_WORKBOOK", "Root office document target cannot be external")
                    return report
                doc_targets.append(posixpath.normpath(rel.get("Target", "").lstrip("/")))
            if "xl/workbook.xml" not in doc_targets:
                report.add("error", "MISSING_ROOT_REL", "Root relationships do not point to xl/workbook.xml")
                return report
            workbook = _part(zf, "xl/workbook.xml", max_part)
            relationships = _relationships(zf, max_part)
            if any("externalLinks/" in p or p.endswith("externalLink.xml") for p in members):
                report.add("warning", "EXTERNAL_LINK_PARTS", "Workbook includes external-link parts")
            if any(p.endswith("vbaProject.bin") for p in members):
                report.add("warning", "MACRO_CONTENT", "Workbook contains VBA binary content")
            sheet_nodes = [el for el in workbook.iter() if _name(el.tag) == "sheet"]
            if not sheet_nodes:
                report.add("error", "NO_SHEETS", "No workbook sheets declared")
                return report
            for entry in sheet_nodes:
                name = entry.get("name", "(unnamed)")
                state = entry.get("state", "visible")
                rid = entry.get(f"{{{R}}}id") or entry.get(f"{{{STRICT_R}}}id")
                if not rid or rid not in relationships:
                    report.add("error", "MISSING_REL", f"No valid relationship for sheet {name}", name)
                    continue
                target, mode, rtype = relationships[rid]
                if mode.lower() == "external":
                    report.add("error", "EXTERNAL_SHEET", "Worksheet relationship is external", name)
                    continue
                if not rtype.endswith("/worksheet"):
                    report.add("info", "NON_WORKSHEET", "Skipping a non-worksheet sheet type", name)
                    continue
                try:
                    part_path = _safe_sheet_path(target)
                except ValueError as exc:
                    report.add("error", "BAD_REL_TARGET", str(exc), name)
                    continue
                if part_path not in members:
                    report.add("error", "MISSING_SHEET", f"Worksheet XML missing: {part_path}", name)
                    continue
                try:
                    sheet_xml = _part(zf, part_path, max_part)
                except ValueError as exc:
                    report.add("error", "MALFORMED_SHEET", str(exc), name)
                    continue
                formula_count = 0
                validation_count = 0
                cached_errors = 0
                protected = False
                for el in sheet_xml.iter():
                    tag = _name(el.tag)
                    if tag == "sheetProtection":
                        protected = True
                    elif tag == "dataValidation":
                        validation_count += 1
                    elif tag == "f":
                        formula_count += 1
                        if FORMULA_ERROR.search(el.text or ""):
                            report.add("error", "BROKEN_FORMULA_REF", "Formula contains #REF!", name)
                    elif tag == "c" and el.get("t") == "e":
                        cached_errors += 1
                report.sheets.append(Sheet(name, part_path, state, protected,
                                           validation_count, formula_count, cached_errors))
                report.total_formulas += formula_count
                report.total_validations += validation_count
                if cached_errors:
                    report.add("warning", "CACHED_FORMULA_ERRORS",
                               f"{cached_errors} cached Excel error cell(s); recalculate in Excel to verify", name)
                if require_protection and not protected:
                    report.add("error", "PROTECTION_REQUIRED", "Sheet has no sheetProtection element", name)
                if state != "visible":
                    report.add("info", "HIDDEN_SHEET", f"Sheet state is {state}", name)
            report.valid = not any(f.severity == "error" for f in report.findings)
    except (FileNotFoundError, PermissionError, IsADirectoryError, OSError) as exc:
        report.add("error", "IO", f"Cannot open workbook: {type(exc).__name__}")
    except (BadZipFile, RuntimeError, EOFError, ValueError, KeyError, ET.ParseError) as exc:
        report.add("error", "INVALID_WORKBOOK", str(exc))
    return report


def compare(baseline: Report, current: Report) -> list[Finding]:
    """Policy regression signals only; intentional changes may need review."""
    if not baseline.valid or not current.valid:
        return [Finding("error", "COMPARE_INVALID", "Both workbooks must pass structural inspection before comparison")]
    old = {s.name: s for s in baseline.sheets}
    new = {s.name: s for s in current.sheets}
    findings: list[Finding] = []
    for name in sorted(old.keys() - new.keys()):
        findings.append(Finding("error", "SHEET_REMOVED", "Sheet removed since baseline", name))
    for name in sorted(new.keys() - old.keys()):
        findings.append(Finding("info", "SHEET_ADDED", "New sheet compared with baseline", name))
    for name in sorted(old.keys() & new.keys()):
        a, b = old[name], new[name]
        if a.protected and not b.protected:
            findings.append(Finding("error", "PROTECTION_LOST", "Sheet protection was lost", name))
        if a.validations > b.validations:
            findings.append(Finding("warning", "VALIDATIONS_REDUCED",
                                    f"Validation rules dropped: {a.validations} -> {b.validations}", name))
        if a.formulas > b.formulas:
            findings.append(Finding("warning", "FORMULAS_REDUCED",
                                    f"Formula elements dropped: {a.formulas} -> {b.formulas}", name))
        if a.state != b.state:
            findings.append(Finding("info", "VISIBILITY_CHANGED",
                                    f"Sheet visibility changed: {a.state} -> {b.state}", name))
    return findings
