# OfficeIMO.Html.Core

Owned HTML document, node, mutation and provider contracts. This package has no third-party runtime dependencies and does not load resources or execute scripts. Parsing, selectors and serialization are supplied through explicit providers.

```sh
dotnet add package OfficeIMO.Html.Core
```

Keep directly referenced OfficeIMO packages on the same version. Applications that
use the AngleSharp provider can install `OfficeIMO.Html.AngleSharp` instead; it
includes this package as a dependency.

Parsed documents are immutable snapshots. Use `document.Edit(editor => ...)` to produce a new snapshot, or `Clone()` for a mutable tree with one owner. Node IDs survive cloning; snapshot identities are independent. Attribute values and text are decoded data. Serializing HTML does not sanitize it.

Use `OfficeIMO.Html.AngleSharp` for the current parser and syntax implementation, and `OfficeIMO.Html` for resource policy, conversion, styles and rendering.
