# OfficeIMO.Zip (Preview)

`OfficeIMO.Zip` provides reusable, dependency-light ZIP traversal primitives for ingestion scenarios.

Current scope:
- deterministic entry enumeration
- structured path safety guards (relative traversal, absolute/drive paths)
- depth and entry-count limits
- uncompressed-size budget limits
- per-entry size and compression-ratio limits
- traversal warnings for rejected/limited entries

Status:
- packaged as `OfficeIMO.Zip`
- used directly by `OfficeIMO.Reader.Zip`
- still preview-scoped while traversal policy and adapter integration continue to evolve
