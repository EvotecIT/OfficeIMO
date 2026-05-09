# OfficeIMO.Epub (Preview)

`OfficeIMO.Epub` is an early reusable EPUB extraction package intended for `OfficeIMO.Reader` adapters.

Current scope:
- opens EPUB as ZIP container
- parses `META-INF/container.xml` and OPF package metadata
- follows OPF manifest + spine ordering
- reads nav/NCX labels for chapter titles when available
- extracts chapter text from XHTML/XML AST (no regex-driven text parsing)
- emits extraction warnings for malformed/unreadable content

Status:
- packaged as `OfficeIMO.Epub`
- used directly by `OfficeIMO.Reader.Epub`
- still preview-scoped while full OPF/spine/nav semantics continue to evolve
- full OPF/spine/nav semantics are planned next
