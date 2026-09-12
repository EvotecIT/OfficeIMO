# OfficeIMO.Html.Core

Owned HTML document, node, mutation and provider contracts. This package has no third-party runtime dependencies and does not load resources or execute scripts. Parsing, selectors and serialization are supplied through explicit providers.

```sh
dotnet add package OfficeIMO.Html.Core
```

Keep directly referenced OfficeIMO packages on the same version. Applications that
use the AngleSharp provider can install `OfficeIMO.Html.AngleSharp` instead; it
includes this package as a dependency.

Parsed documents are immutable snapshots. Use `document.Edit(editor => ...)` to produce a new snapshot, or `Clone()` for a mutable tree with one owner. Node IDs survive cloning; snapshot identities are independent. Attribute values and text are decoded data. Serializing HTML does not sanitize it.

Runtime captures can attach immutable `HtmlFormControlState` to an element's
`FormState` property. This retains current input and textarea values, input
checkedness and indeterminate state, and option selections independently of HTML
attributes and text. Clone and import preserve the state. On a mutable tree,
replace it to edit the captured value or set it to null to return to authored
defaults. Ordinary HTML serialization retains attributes and text; it does not
encode this separate state. Use the owned document with OfficeIMO conversion to
inspect or render the captured values.
Option state also applies when an option is imported into a select with no live
state. A select with explicit live state can retain an empty selection. When an
edit changes a multi-select into a dropdown, the last selected option in document
order becomes its effective selection; the stored option flags remain available.

`Clone()` retains detached nodes so existing node IDs remain addressable. For a
long-lived editor that no longer needs removed content, start the next session
with `document.CloneAttached(cancellationToken)`. It copies the attached tree and
template contents, preserving their node IDs and source offsets, and omits
detached editing history. The original document and its handles remain unchanged;
release them when they are no longer needed. The returned copy is mutable and can
be frozen after editing. A cancelled attached clone returns no partial document.

`GetAttribute(name)`, `HasAttribute(name)` and `RemoveAttribute(name)` use the
qualified attribute name, including namespaced attributes. Use the namespace and
local-name overloads to select a specific namespace. `Id`, `ClassName` and
`ClassList` reflect only empty-namespace attributes. `SetAttribute` adds or
replaces an attribute in the supplied namespace; its default is the empty namespace.

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

CSS inspection is available without an HTML parser or rendering dependency:

```csharp
using OfficeIMO.Html.Css;

string css = "color: red; --layout: {columns: 2; gap: 12px};";
foreach (HtmlCssToken token in HtmlCssTokenizer.Tokenize(css)) {
    Console.WriteLine($"{token.Kind} at {token.Offset}: {token.GetText(css)}");
}
```

The tokenizer preserves comments and exact UTF-16 source spans, including original
line endings. `Value` contains decoded identifiers, strings, URLs and dimension
units; numeric spelling is available through `GetText`. A zero-length
`EndOfFile` token terminates the result. Comments are source trivia, not whitespace.
Tokenization does not validate property values, resolve selectors or compute styles.

`HtmlCssTokenizationOptions` defaults to 8,388,608 UTF-16 characters and one million
tokens, excluding `EndOfFile`. Set either limit to `null` for caller-bounded input.
An exceeded budget throws `HtmlCssTokenizationLimitException`; cancellation throws
`OperationCanceledException`. Neither operation returns a partial list.
