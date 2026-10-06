# Security Policy

## Report a vulnerability

Report suspected vulnerabilities privately through [GitHub private vulnerability reporting](https://github.com/EvotecIT/OfficeIMO/security/advisories/new). Include the affected OfficeIMO package and version, a minimal reproduction, the expected and observed behavior, and the likely impact. Do not include secrets or private documents in the report.

## Security boundaries

OfficeIMO reads caller-supplied documents, archives, and markup. Treat source bytes, package parts, metadata, links, and MCP tool arguments as untrusted. Parsing and conversion should enforce configured input, entry, expansion, recursion, and output limits. Interactive links and active package content must follow explicit policies. OfficeIMO.Tool MCP filesystem operations must stay within configured allowed roots.

OfficeIMO.Reader.Web uses a caller-owned `HttpClient`. Its URI checks do not validate DNS answers or redirects before connection; applications that accept untrusted URLs must enforce those network rules in their HTTP handler.
