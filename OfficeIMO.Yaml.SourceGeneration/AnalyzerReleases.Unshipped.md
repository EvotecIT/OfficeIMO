; Unshipped analyzer release
; https://github.com/dotnet/roslyn-analyzers/blob/main/src/Microsoft.CodeAnalysis.Analyzers/ReleaseTrackingAnalyzers.Help.md

### New Rules

Rule ID | Category | Severity | Notes
--------|----------|----------|--------------------
YAML001 | OfficeIMO.Yaml | Error | A YAML serializer context must be a partial, non-generic class
YAML002 | OfficeIMO.Yaml | Error | A type reached by a YAML serializer context has a shape the serializer does not support
YAML003 | OfficeIMO.Yaml | Error | A type reached by a YAML serializer context has no constructor reading can use
YAML004 | OfficeIMO.Yaml | Error | Two properties of a type have the same YAML name
YAML005 | OfficeIMO.Yaml | Error | A derived-type registration is invalid
YAML006 | OfficeIMO.Yaml | Error | A type reached by a YAML serializer context is not accessible from the context
