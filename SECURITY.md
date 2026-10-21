# Security policy

SheetGuard parses user-provided files; treat any workbook from an unknown source as untrusted. The tool is deliberately read-only and limits archive sizes, rejects suspicious ZIP paths and blocks XML DTD/entity declarations, but it is **not** a malware scanner or sandbox.

Do not use it to open suspicious spreadsheets with Excel. Never commit real customer data to the repository.

Please report suspected security vulnerabilities privately using GitHub's private vulnerability reporting feature (if enabled) rather than a public issue. If private reporting is unavailable, contact the maintainer through a private channel without publicly disclosing exploit details.

No security guarantees or remediation SLA are offered for this alpha release.
