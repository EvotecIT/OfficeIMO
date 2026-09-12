# OfficeIMO.Html.Core

Owned HTML document, node, mutation and provider contracts. This package has no third-party runtime dependencies and does not load resources or execute scripts. Parsing, selectors and serialization are supplied through explicit providers.

```sh
dotnet add package OfficeIMO.Html.Core
```

Keep directly referenced OfficeIMO packages on the same version. Applications that
use the AngleSharp provider can install `OfficeIMO.Html.AngleSharp` instead; it
includes this package as a dependency.

Parsed documents are immutable snapshots. Use `document.Edit(editor => ...)` to produce a new snapshot, or `Clone()` for a mutable tree with one owner. Node IDs survive cloning; snapshot identities are independent. Attribute values and text are decoded data. Serializing HTML does not sanitize it.

`Clone()` retains detached nodes so existing node IDs remain addressable. For a
long-lived editor that no longer needs removed content, start the next session
with `document.CloneAttached(cancellationToken)`. It copies the attached tree and
template contents, preserving their node IDs and source offsets, and omits
detached editing history. The original document and its handles remain unchanged;
release them when they are no longer needed. The returned copy is mutable and can
be frozen after editing. A cancelled attached clone returns no partial document.

Use `OfficeIMO.Html.AngleSharp` for the current parser and syntax implementation, and `OfficeIMO.Html` for resource policy, conversion, styles and rendering.

To copy content between documents, import it into the destination and then insert
the returned detached node:

```csharp
HtmlDocument edited = destination.Edit(editor => {
    HtmlNode imported = editor.ImportNode(source.QuerySelector("section")!);
    editor.Body!.AppendChild(imported);
});
```

Import preserves namespaces, decoded text, attributes and template contents. It
assigns new destination-local IDs and clears source offsets because the copy has
no position in the destination's original input. `deep: false` copies only the
node and its attributes. Importing a fragment copies its children; appending that
fragment splices them into the destination. Neither operation sanitizes content.
For cancellation with atomic snapshot publication, perform the import in `Edit`;
a cancelled import into a mutable document can leave partial detached copies.
