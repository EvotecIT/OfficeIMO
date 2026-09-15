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
Every `HtmlCssRule` exposes its direct `Declarations` and child `Rules` as convenience
views over `Contents`; the combined `Contents` sequence remains the source for authored
interleaving. These views do not assign semantics to unknown at-rule block grammars, whose
raw `Block` remains authoritative.

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

HtmlCssMathParseResult width = HtmlCssMathParser.ParseLengthPercentage(
    "calc(24px + 25%)");
HtmlCssLengthResolutionResult usedWidth = HtmlCssMathResolver.ResolveLength(
    width.Expression!,
    new HtmlCssLengthResolutionContext {
        PercentageReference = 640,
        FontSize = 16,
        RootFontSize = 16,
        ViewportWidth = 1280,
        ViewportHeight = 720
    });
```

`HtmlCssPropertyCatalog` currently owns CSS-wide keywords and selected values for
`display`, `visibility`, `opacity`, `color`, width and height constraints, and the four
physical margin and padding longhands. The result distinguishes a parsed value, a `var()`
value deferred until substitution, an unknown property, a value outside the implemented
slice, and malformed component syntax. `HtmlCssPropertyValue.MathExpression` exposes the
typed tree for the sizing and spacing properties. Typed grammar support is an inspection
contract; it does not imply that a renderer paints every parsed value.
Named, hexadecimal, and CSS Color 4 system colors plus `currentColor` are typed now;
the `color` definition exposes `CanvasText` as its initial value. The typed functional
slice covers legacy and modern `rgb()`/`rgba()` and `hsl()`/`hsla()`, plus modern
`hwb()`. Numeric expressions cover constant number and percentage arithmetic through
`calc()`, `min()`, `max()`, and `clamp()` with type checking. Length-percentage expressions
retain percentages through parsing and computed-style inspection, then resolve only when
the caller supplies the consuming property's percentage reference. Resolution returns a
specific missing-context status instead of guessing a value.

The length subset covers `px`, `pt`, `pc`, `in`, `cm`, `mm`, `q`, `em`, `rem`, the
`vw`/`vh`/`vmin`/`vmax` families with `sv`, `lv`, and `dv` variants, and
`cqw`/`cqh`/`cqi`/`cqb`/`cqmin`/`cqmax`. Specialized viewport dimensions can be supplied
separately and otherwise use the declared default viewport. Container units use the query
container dimensions and fall back to the small viewport dimensions when no eligible
container is supplied. The current static subset treats inline and block container sizes as
explicit context; it does not infer writing mode. Font metric units such as `ex`, `cap`,
`ch`, `ic`, and `lh`, wider Color 4 spaces, relative colors, color interpolation,
multi-keyword display values, and `visibility: force-hidden` remain explicit grammar gaps
even where the full conversion package may already render some of them through its retained
implementation.

`HtmlCssMathOptions` bounds input, tokens, nesting, operations, and comparison arguments.
Limit exhaustion throws `HtmlCssMathLimitException`; cancellation publishes no partial tree.
The parser requires CSS whitespace around binary `+` and `-`, keeps unitless zero valid as a
standalone length, and rejects unitless zero as a length inside math expressions. Division by
zero remains a typed expression and resolves with `NonFiniteValue`, allowing property and
rendering consumers to apply their documented fallback.

Parse and match the selected selector subset directly against an owned element. Namespace
prefixes are stylesheet-scoped, so the same immutable context can be passed to every selector
parsed from that sheet:

```csharp
HtmlCssStyleSheet sheet = HtmlCssSyntaxParser.ParseStyleSheet(
    "@namespace svg url('http://www.w3.org/2000/svg');");
var selectorOptions = new HtmlCssSelectorOptions {
    Namespaces = HtmlCssNamespaceContext.FromStyleSheet(sheet)
};
HtmlCssSelectorListParseResult parsedSelectors = HtmlCssSelectorParser.ParseList(
    "main > article.card:last-child:is([data-state='READY' i], .queued), svg|a:first-child",
    selectorOptions);
if (parsedSelectors.IsSupported) {
    bool matches = parsedSelectors.SelectorList!.Matches(article);
    Console.WriteLine(parsedSelectors.SelectorList.Selectors[0].Specificity);
}
```

The owned selector slice covers type, universal, id, class, namespace-qualified type and
attribute selectors, attribute comparison modifiers, and descendant, child, adjacent-sibling,
and general-sibling combinators. It implements `:root`, `:empty`, first/last/only child and
of-type forms, unfiltered `:nth-child()`, `:nth-last-child()`, `:nth-of-type()`, and
`:nth-last-of-type()` An+B expressions, a single basic identifier `:lang()` range, plus
`:is()`, `:where()`, and `:not()`. Logical
specificity follows Selectors Level 4: `:where()` contributes zero while `:is()` and `:not()`
use their most specific argument. `:is()` and `:where()` discard malformed list members;
`:not()` keeps its strict list grammar. In logical pseudo-class arguments, an implicit
universal on the final subject compound is not constrained by a stylesheet default namespace;
preceding compounds still use that default namespace.

`Parse` retains the single-complex-selector contract and classifies a top-level comma as
`Unsupported`. `ParseList` parses a comma-separated list atomically; every member must be in
the owned subset before a match can run. Both results distinguish malformed syntax from valid
syntax that needs a provider. Dynamic state pseudo-classes, `:has()`, filtered
`:nth-child(... of S)`, pseudo-elements, and nesting remain explicit provider cases. Default
namespaces apply to explicit and implicit type/universal selectors but not to unprefixed
attributes. `:empty` follows deployed browser text-node semantics: whitespace text makes an
element non-empty, while comments do not. `HtmlCssSelectorOptions` bounds source length, token count, compounds, simple
selectors, list members, and logical nesting. Parsing and matching observe cancellation, and
limit failures publish no partial selector.

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
