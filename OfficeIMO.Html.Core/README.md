# OfficeIMO.Html.Core

Owned HTML document, node, mutation, CSS syntax, typed value, and selected selector contracts. This package has no third-party runtime dependencies and does not load resources or execute scripts. HTML parsing and serialization, plus selector features outside the owned subset, are supplied through explicit providers.

```sh
dotnet add package OfficeIMO.Html.Core
```

Keep directly referenced OfficeIMO packages on the same version. Applications that
use the AngleSharp provider can install `OfficeIMO.Html.AngleSharp` instead; it
includes this package as a dependency.

Parsed documents and contextual fragments are immutable snapshots. Implementations of
`IHtmlParserProvider` supply both operations without exposing provider nodes. Use
`document.Edit(editor => ...)` to produce a new snapshot, or `Clone()` for a mutable tree
with one owner. Node IDs survive cloning; snapshot identities are independent. Attribute
values and text are decoded data. Serializing HTML does not sanitize it.

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

Use `OfficeIMO.Html.AngleSharp` for the current HTML parser, selectors, serializer and charset labels, and `OfficeIMO.Html` for resource policy, conversion, styles and rendering.

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

CSS inspection is available without an HTML parser or rendering dependency. Parse a
stylesheet when you need rules and nested contents, or a style block for inline CSS:

```csharp
using OfficeIMO.Html.Css;

string css = "@future demo; .card { color: red; --layout: {columns: 2}; }";
HtmlCssStyleSheet sheet = HtmlCssSyntaxParser.ParseStyleSheet(css);
foreach (HtmlCssRule rule in sheet.Rules) {
    Console.WriteLine($"{rule.Kind} at {rule.Span}: {rule.GetText()}");
}

HtmlCssStyleBlock inline = HtmlCssSyntaxParser.ParseStyleBlock(
    "color: red; future-property: paint(foo(1, [two]));");
foreach (HtmlCssDeclaration declaration in inline.Declarations)
    Console.WriteLine($"{declaration.Name}: {declaration.ValueSpan.GetText(inline.Source)}");
```

`HtmlCssStyleSheet` and `HtmlCssStyleBlock` preserve the complete original input.
Rules, declarations, unknown at-rules, duplicate properties, invalid recovered source,
comments, whitespace, functions and simple blocks retain exact UTF-16 spans and authored
order. `GetPosition(offset)` maps a span to a one-based line and column. `ToCss()` returns
the original source character-for-character, including its original line endings. Diagnostics report syntax
recovery; an unknown name or value is preserved without being treated as invalid.

Apply the owned property grammar when tooling needs more than lossless syntax:

```csharp
HtmlCssPropertyParseResult value = HtmlCssPropertyParser.Parse("opacity", "62.5%");
if (value.IsAccepted)
    Console.WriteLine($"{value.Value!.Kind}: {value.Value.CanonicalText}");

HtmlCssDeclaration color = HtmlCssSyntaxParser
    .ParseStyleBlock("color: rebeccapurple !important")
    .Declarations[0];
HtmlCssPropertyParseResult declarationValue = HtmlCssPropertyParser.Parse(color);

HtmlCssPropertyParseResult calculated = HtmlCssPropertyParser.Parse(
    "opacity", "clamp(10%, 75%, 60%)");
HtmlCssNumericValue numeric = calculated.Value!.NumericValue!;
```

`HtmlCssPropertyCatalog` currently owns CSS-wide keywords and selected values for
`display`, `visibility`, `opacity`, and `color`. The result distinguishes a parsed
value, a `var()` value deferred until substitution, an unknown property, a value
outside the implemented slice, and malformed component syntax. Typed grammar support
is an inspection contract; it does not imply that a renderer paints every parsed value.
Named, hexadecimal, and CSS Color 4 system colors plus `currentColor` are typed now;
the `color` definition exposes `CanvasText` as its initial value. The typed functional
slice covers legacy and modern `rgb()`/`rgba()` and `hsl()`/`hsla()`, plus modern
`hwb()`. Numeric expressions cover constant number and percentage arithmetic through
`calc()`, `min()`, `max()`, and `clamp()` with type checking. Wider Color 4 spaces,
relative colors, color interpolation, dimensions, multi-keyword display values, and
`visibility: force-hidden` remain explicit grammar gaps even where the full conversion
package may already render some of them through its retained implementation.

Parse and match the selected selector subset directly against an owned element:

```csharp
HtmlCssSelectorParseResult parsedSelector = HtmlCssSelectorParser.Parse(
    "main > article.card[data-state='READY' i]");
if (parsedSelector.IsSupported) {
    bool matches = parsedSelector.Selector!.Matches(article);
    Console.WriteLine(parsedSelector.Selector.Specificity);
}
```

The owned selector slice covers type, universal, id, class, and attribute selectors,
including attribute comparison modifiers, with descendant, child, adjacent-sibling,
and general-sibling combinators. A parse result distinguishes malformed syntax from a
valid selector that needs a provider. Selector lists are split by stylesheet consumers;
the standalone parser accepts one complex selector. Namespace selectors, pseudo-classes,
pseudo-elements, nesting, and functional selectors currently return `Unsupported`.
`HtmlCssSelectorOptions` bounds source length, token count, compounds, and simple selectors;
cancellation and limit failures publish no partial selector.

The lossless syntax model deliberately preserves selector preludes without forcing the
owned subset on every rule. Apply `HtmlCssSelectorParser` when a consumer wants validation
or matching. `OfficeIMO.Html.Core` does not apply the cascade, resolve computed values, load
resources, or claim rendering support, so consumers can retain and inspect newer CSS while
the full `OfficeIMO.Html` package continues to use its qualified style and rendering pipeline.

`HtmlCssTokenizer.Tokenize` remains available for lexical tooling. It preserves comments
and exact UTF-16 source spans, including original line endings. `Value` contains decoded
identifiers, strings, URLs and dimension units; numeric spelling is available through
`GetText`. A zero-length `EndOfFile` token terminates the result. Comments are source
trivia, not whitespace.

`HtmlCssTokenizationOptions` defaults to 8,388,608 UTF-16 characters and one million
tokens, excluding `EndOfFile`. Set either limit to `null` for caller-bounded input.
An exceeded budget throws `HtmlCssTokenizationLimitException`; cancellation throws
`OperationCanceledException`. Neither operation returns a partial list.

`HtmlCssSyntaxOptions` adds nesting and syntax-node limits to the same input and token
bounds. An exceeded syntax limit throws `HtmlCssSyntaxLimitException`. A canceled or
failed syntax parse publishes no partial document.
