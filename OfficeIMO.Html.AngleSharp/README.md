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

HtmlDocument document = AngleSharpHtmlParser.Instance.Parse(
    "<h1>Hello</h1>", new HtmlParseOptions());
HtmlDocument edited = document.Edit(editor => {
    editor.QuerySelector("h1")!.TextContent = "Updated";
});
string html = edited.OuterHtml;
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
The parser handles a full document; `HtmlParseOptions` does not select a fragment
context.
