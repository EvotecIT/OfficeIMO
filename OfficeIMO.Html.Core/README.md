# OfficeIMO.Html.Core

Owned HTML document, node, mutation and provider contracts. This package has no third-party runtime dependencies and does not load resources or execute scripts. Parsing, selectors and serialization are supplied through explicit providers.

Parsed documents are immutable snapshots. Use `document.Edit(editor => ...)` to produce a new snapshot, or `Clone()` for a mutable tree with one owner. Node IDs survive cloning; snapshot identities are independent. Attribute values and text are decoded data. Serializing HTML does not sanitize it.

Use `OfficeIMO.Html.AngleSharp` for the current parser and syntax implementation, and `OfficeIMO.Html` for resource policy, conversion, styles and rendering.
