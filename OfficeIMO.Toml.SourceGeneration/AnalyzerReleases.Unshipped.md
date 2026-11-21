; Unshipped analyzer release
; https://github.com/dotnet/roslyn-analyzers/blob/main/src/Microsoft.CodeAnalysis.Analyzers/ReleaseTrackingAnalyzers.Help.md

### New Rules

Rule ID | Category | Severity | Notes
--------|----------|----------|--------------------
TOML001 | OfficeIMO.Toml | Error | A TOML serializer context must be a partial, non-generic class
TOML002 | OfficeIMO.Toml | Error | A type reached by a TOML serializer context has a shape the serializer does not support
TOML003 | OfficeIMO.Toml | Error | A type reached by a TOML serializer context has no constructor reading can use
TOML004 | OfficeIMO.Toml | Error | Two properties of a type have the same TOML name
TOML005 | OfficeIMO.Toml | Error | A derived-type registration is invalid
