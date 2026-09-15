# OfficeIMO.Html.AngleSharp

AngleSharp-backed parsing, selectors and serialization for the owned OfficeIMO HTML document model. Parsing is inert: it does not execute scripts or fetch resources.

```sh
dotnet add package OfficeIMO.Html.AngleSharp
```

This includes `OfficeIMO.Html.Core`. Keep directly referenced OfficeIMO packages on
the same version. Install `OfficeIMO.Html` for resource policy, conversion, styles
and rendering; it includes this provider.

```csharp
using OfficeIMO.Html.Dom;
using OfficeIMO.Html.Providers;

HtmlDocument document = AngleSharpHtmlParser.Instance.ParseDocument(
    "<h1>Hello</h1>", new HtmlParseOptions());
HtmlDocument edited = document.Edit(editor => {
    editor.QuerySelector("h1")!.TextContent = "Updated";
});
string html = edited.OuterHtml;

HtmlDocument table = AngleSharpHtmlParser.Instance.ParseDocument(
    "<table><tbody><tr id='items'></tr></tbody></table>", new HtmlParseOptions());
HtmlDocumentFragment cells = AngleSharpHtmlParser.Instance.ParseFragment(
    "<td>A</td><td>B</td>", table.QuerySelector("#items")!, new HtmlParseOptions());
```

The provider retains a native tree alongside owned nodes for selector and conversion reuse. The existing CSS/layout engine still uses a structural native-DOM adapter internally. This package does not claim dependency independence or browser execution. Source and tree budgets are enforced before downstream conversion work; parsing remains subject to the underlying parser's cooperative cancellation behavior.

`AngleSharpEncodingProvider` implements the owned `IHtmlEncodingProvider` charset
contract. It retains AngleSharp web aliases and `System.Text.Encoding.CodePages`;
initialization preserves the existing process-wide code-page registration. The
HTML conversion package owns byte-order-mark precedence, HTML prescanning, CSS
decoding rules and stream lifetime. This provider supplies label resolution.

Native trees are constructed before node/depth budgets can be checked. Source
length is checked first, and parser cancellation is cooperative. Query results use
owned nodes, including detached nodes. Invalid selectors throw `ArgumentException`.
Conversion exports the attached tree, including template content. A detached tree
is materialized only when that tree is queried or serialized; unrelated detached
trees do not affect conversion or document queries.
`ParseDocument` handles complete HTML documents. `ParseFragment` accepts an owned
context element and returns a fragment in an independent owned document. The adapter
reconstructs foreign and ancestor-form context where the retained native fragment API
does not preserve it. Table insertion modes use the native fragment path. Both paths
apply the same source, node, depth and cooperative-cancellation options.
