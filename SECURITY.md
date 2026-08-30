# Security Policy

## Supported versions

Security fixes are provided for the latest published release line. Release candidates are for evaluation and should not replace a stable production version without application-level testing.

## Reporting a vulnerability

Please use GitHub's **Report a vulnerability** private security-advisory flow for this repository. Do not open a public issue containing exploit details, sensitive workbook contents, credentials, or customer data.

Include the affected version, minimal reproduction, impact, and any suggested mitigation. Maintainers will acknowledge a complete report as soon as practical and coordinate disclosure after a fix is available. There is currently no bug-bounty program.

## Application responsibilities

Excel Cursor is a workbook library, not a sandbox or file-upload security product.

- Treat output paths as trusted configuration. Confine them to an application-owned root and handle symlinks, permissions, temporary files, atomic replacement, retention, and cleanup in the host application.
- Use `setSafeText()` for untrusted strings. Only trusted code may call formula APIs or construct ExcelJS formula/hyperlink values.
- Choose workload-specific resource limits. Defaults reduce accidental large synchronous operations but do not provide tenant isolation.
- Validate uploaded files before use and run untrusted-workbook processing in an isolated, resource-limited worker. This project does not scan macros, malware, external links, or archive bombs.
- Protect generated workbooks at rest and in transit. This project does not add password protection or encryption.
- A cancelled or failed streaming export may leave a partial file; remove it according to application policy.

Never attach confidential workbooks to public security reports.
