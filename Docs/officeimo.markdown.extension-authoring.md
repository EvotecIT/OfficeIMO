# OfficeIMO.Markdown Extension Authoring

This guide captures the current public extension surface for `OfficeIMO.Markdown` after the recent parser, AST, and renderer seam work.

Runnable sample:

- `OfficeIMO.Examples/Markdown/Markdown07_Custom_Extensions.cs`
- `OfficeIMO.Examples/Markdown/Markdown08_Custom_Block_Parsers.cs`
- `OfficeIMO.Examples/Markdown/Markdown09_Custom_Markdown_Write_Overrides.cs`

## What You Can Extend Today

The package now has public seams for:

- block parser registration with `MarkdownBlockParserExtension`
- fenced block registration with `MarkdownFencedBlockExtension`
- inline parser registration with `MarkdownInlineParserExtension`
- post-parse semantic transforms with `IMarkdownDocumentTransform`
- syntax-tree participation with:
  - `ISyntaxMarkdownInline`
  - `ISyntaxMarkdownBlock`
  - `ISyntaxMarkdownBlockWithContext`
- nested block traversal with:
  - `IChildMarkdownBlockContainer`
- HTML rendering with:
  - `IMarkdownBlock.RenderHtml()`
  - `IContextualHtmlMarkdownBlock`
  - `IContextualHtmlMarkdownInline`
- Markdown rendering with:
  - `MarkdownBlockMarkdownRenderExtension`
  - `MarkdownWriteContext`
- markdown/html inline participation with:
  - `IRenderableMarkdownInline`
  - `IPlainTextMarkdownInline`
  - `IInlineContainerMarkdownInline`

That means an external extension can now:

- parse custom blocks or inline tokens
- emit typed semantic objects into `MarkdownDoc`
- participate in `ParseWithSyntaxTree(...)`
- expose named AST nodes through `MarkdownSyntaxNode.CustomKind`
- render with normal HTML options/body context

## Recommended Mental Model

Use the public layers like this:

1. Parsing
   Register block, fenced-block, or inline parsers on `MarkdownReaderOptions`.
2. Semantic model
   Return real block/inline objects instead of string placeholders.
3. Syntax tree
   Implement `ISyntaxMarkdownInline` or `ISyntaxMarkdownBlockWithContext` when you want more than the default `Unknown` node fallback.
4. Rendering
   Implement `IMarkdownBlock.RenderHtml()` for simple self-contained output.
   Implement `IContextualHtmlMarkdownBlock` when rendering depends on `HtmlOptions` or surrounding blocks.

## Inline Extension Pattern

This is the minimum useful pattern for a custom inline:

```csharp
using OfficeIMO.Markdown;
using System.Text;

public sealed class DoubleBraceInline :
    MarkdownInline,
    IRenderableMarkdownInline,
    IContextualHtmlMarkdownInline,
    IPlainTextMarkdownInline,
    IInlineContainerMarkdownInline,
    ISyntaxMarkdownInline {
    public DoubleBraceInline(InlineSequence inlines) {
        Inlines = inlines ?? new InlineSequence();
    }

    public InlineSequence Inlines { get; }

    public string RenderMarkdown() => "{{" + Inlines.RenderMarkdown() + "}}";

    public string RenderHtml() => "<span data-inline=\"double-brace\">" + Inlines.RenderHtml() + "</span>";

    string IContextualHtmlMarkdownInline.RenderHtml(HtmlOptions options) =>
        "<span data-inline=\"double-brace\" data-title=\""
        + System.Net.WebUtility.HtmlEncode(options.Title)
        + "\">"
        + Inlines.RenderHtml()
        + "</span>";

    public void AppendPlainText(StringBuilder sb) => InlinePlainText.AppendPlainText(sb, Inlines);

    InlineSequence? IInlineContainerMarkdownInline.NestedInlines => Inlines;

    public MarkdownSyntaxNode BuildSyntaxNode(MarkdownInlineSyntaxBuilderContext context, MarkdownSourceSpan? span) {
        var children = context.BuildChildren(Inlines);
        return new MarkdownSyntaxNode(
            MarkdownSyntaxKind.Unknown,
            span ?? context.GetAggregateSpan(children),
            literal: RenderMarkdown(),
            children: children,
            associatedObject: this,
            customKind: "double-brace");
    }
}

public static class DoubleBraceExtension {
    public static void Add(MarkdownReaderOptions options) {
        options.InlineParserExtensions.Add(new MarkdownInlineParserExtension(
            "DoubleBrace",
            TryParse));
    }

    private static bool TryParse(MarkdownInlineParserContext context, out MarkdownInlineParseResult result) {
        result = default;
        if (context.CurrentChar != '{'
            || context.Position + 1 >= context.Text.Length
            || context.Text[context.Position + 1] != '{') {
            return false;
        }

        var closing = context.Text.IndexOf("}}", context.Position + 2, StringComparison.Ordinal);
        if (closing < 0) {
            return false;
        }

        var innerLength = closing - (context.Position + 2);
        var nested = context.ParseNestedInlines(2, innerLength);
        result = new MarkdownInlineParseResult(
            new DoubleBraceInline(nested),
            closing + 2 - context.Position);
        return true;
    }
}
```

Why each interface matters:

- `IRenderableMarkdownInline`
  lets the node participate in markdown and HTML output
- `IContextualHtmlMarkdownInline`
  lets HTML output react to the active `HtmlOptions` during real document rendering
- `IPlainTextMarkdownInline`
  keeps plain-text extraction, heading text, and compatibility views working
- `IInlineContainerMarkdownInline`
  lets the object binder and syntax builder descend into nested children
- `ISyntaxMarkdownInline`
  lets the extension publish a stable AST node instead of a generic fallback

## Block Extension Pattern

For fenced blocks, prefer returning a custom block object from `MarkdownFencedBlockExtension`.

```csharp
using OfficeIMO.Markdown;

public sealed class VendorChartBlock :
    MarkdownBlock,
    IMarkdownBlock,
    ISyntaxMarkdownBlockWithContext,
    IContextualHtmlMarkdownBlock {
    public VendorChartBlock(string language, string payload) {
        Language = language ?? string.Empty;
        Payload = payload ?? string.Empty;
    }

    public string Language { get; }
    public string Payload { get; }

    string IMarkdownBlock.RenderMarkdown() => $"```{Language}\n{Payload}\n```";

    string IMarkdownBlock.RenderHtml() =>
        $"<pre><code class=\"language-{System.Net.WebUtility.HtmlEncode(Language)}\">{System.Net.WebUtility.HtmlEncode(Payload)}</code></pre>";

    string IContextualHtmlMarkdownBlock.RenderHtml(MarkdownBodyRenderContext context) =>
        $"<div data-vendor-chart=\"true\" data-title=\"{System.Net.WebUtility.HtmlEncode(context.Options.Title)}\">{System.Net.WebUtility.HtmlEncode(Payload)}</div>";

    public MarkdownSyntaxNode BuildSyntaxNode(MarkdownBlockSyntaxBuilderContext context, MarkdownSourceSpan? span) {
        var payloadNode = context.BuildInlineContainerNode(
            MarkdownSyntaxKind.Paragraph,
            new InlineSequence().Text(Payload),
            literal: Payload);

        var children = new[] {
            new MarkdownSyntaxNode(MarkdownSyntaxKind.CodeFenceInfo, literal: Language),
            payloadNode
        };

        return new MarkdownSyntaxNode(
            MarkdownSyntaxKind.Unknown,
            span ?? context.GetAggregateSpan(children),
            literal: context.NormalizeLiteralLineEndings(((IMarkdownBlock)this).RenderMarkdown()),
            children: children,
            associatedObject: this,
            customKind: "vendor-chart");
    }
}

var options = new MarkdownReaderOptions();
options.FencedBlockExtensions.Add(new MarkdownFencedBlockExtension(
    "Vendor charts",
    new[] { "vendor-chart" },
    context => new VendorChartBlock(context.Language, context.Content)));
```

Use:

- `ISyntaxMarkdownBlock`
  when the syntax node is simple and self-contained
- `ISyntaxMarkdownBlockWithContext`
  when you want helper methods for child blocks, inline children, aggregate spans, or literal normalization
- `IContextualHtmlMarkdownBlock`
  when output depends on `HtmlOptions`, heading catalog behavior, or surrounding blocks

`MarkdownBodyRenderContext` now also exposes public helper methods for extension authors:

- `GetBlockIndex(...)`
  find the current top-level block position
- `GetHeadingAnchor(...)`
  resolve the active anchor slug for a heading block
- `GetPrecedingHeadingAnchor(...)`
  resolve the title heading anchor for section-scoped rendering
- `BuildTocEntries(...)`
  generate TOC-style heading entries without depending on internal catalog types

`HtmlOptions.BlockRenderExtensions` also supports context-aware override registrations now. When you need an external override to win before the block's own `IContextualHtmlMarkdownBlock` implementation, register a `MarkdownBlockHtmlRenderExtension` with `MarkdownBlockHtmlRenderExtension.CreateContextual(...)`.

`MarkdownWriteOptions.BlockRenderExtensions` now has the same pattern through `MarkdownBlockMarkdownRenderExtension.CreateContextual(...)`, with `MarkdownWriteContext` exposing:

- `GetBlockIndex(...)`
- `GetHeadingAnchor(...)`
- `GetPrecedingHeadingAnchor(...)`
- `BuildTocEntries(...)`

See `OfficeIMO.Examples/Markdown/Markdown09_Custom_Markdown_Write_Overrides.cs` for a runnable example that writes both default markdown and a customized markdown profile using a contextual TOC override plus a legacy-compatible callout override.

## Delegate Block Parser Pattern

For non-fenced syntax, the public block parser API now has the same delegate-style ergonomics as inline parsing.

```csharp
var options = MarkdownReaderOptions.CreatePortableProfile();
options.BlockParserExtensions.Add(new MarkdownBlockParserExtension(
    "Panel blocks",
    MarkdownBlockParserPlacement.BeforeParagraphs,
    TryParsePanel));

static bool TryParsePanel(MarkdownBlockParserContext context, out MarkdownBlockParseResult result) {
    result = default;
    if (!context.CurrentLine.TrimStart().StartsWith(":::panel ", StringComparison.Ordinal)) {
        return false;
    }

    var title = context.CurrentLine.Trim().Substring(":::panel".Length).Trim();
    var closingOffset = -1;
    for (var offset = 1; context.TryGetLine(offset, out var line); offset++) {
        if (string.Equals(line.Trim(), ":::", StringComparison.Ordinal)) {
            closingOffset = offset;
            break;
        }
    }

    if (closingOffset < 0) {
        return false;
    }

    var childBlocks = closingOffset > 1
        ? context.ParseNestedBlocks(1, closingOffset - 1)
        : Array.Empty<IMarkdownBlock>();
    result = new MarkdownBlockParseResult(new PanelBlock(title, childBlocks), closingOffset + 1);
    return true;
}
```

Key pieces:

- `MarkdownBlockParserContext`
  exposes the current line, surrounding source lines, reader options, shared reader state, and the in-progress document
- `ParseNestedBlocks(...)`
  lets custom block syntax reuse the core parser for nested markdown while preserving source spans
- `MarkdownBlockParseResult`
  returns the blocks produced by the parser plus the number of consumed source lines
- `IChildMarkdownBlockContainer`
  lets external custom blocks expose nested child blocks so traversal, rewriters, source-span binding, and syntax-tree generation can descend into them

## Source Spans and AST Notes

- Inline source spans are assigned by the reader when your inline parser returns a node through `MarkdownInlineParseResult`.
- Block source spans are assigned when the block is part of the parsed document and its syntax node `AssociatedObject` points back to the block.
- Use `customKind` on `MarkdownSyntaxNode` for extension-specific AST identity without forcing enum changes in `MarkdownSyntaxKind`.

Recommended practice:

- keep `Kind = MarkdownSyntaxKind.Unknown` for extension-owned nodes unless the node truly matches an existing core syntax kind
- use a stable `CustomKind` string such as `vendor-chart`, `double-brace`, or `my-company.alert`

## Choosing Between Extension Types

- Use `MarkdownBlockParserExtension` when the syntax is block-structured and does not naturally fit fenced code.
- Use `MarkdownFencedBlockExtension` when a fenced language token is the right contract boundary.
- Use `MarkdownInlineParserExtension` for local token syntax inside paragraphs/headings/list items.
- Use `IMarkdownDocumentTransform` for AST upgrades that should happen after parsing, not during token recognition.

## Current Limits

The extension surface is much stronger than before, but a few things are still intentionally limited:

- `MarkdownSyntaxKind` remains a fixed enum, so extension node identity should flow through `CustomKind`
- inline HTML rendering does not yet have a separate override registry beyond the inline node itself
- block HTML override extensions still operate by block type, not by syntax-node shape
- the syntax tree is semantic-friendly, but it is still not a fully lossless token stream

## Suggested Pattern For Real Packages

For a production extension package, keep code organized like this:

- `Reader`
  registration helpers and parser delegates
- `Blocks` / `Inlines`
  semantic nodes
- `Rendering`
  HTML helper methods or shared rendering utilities
- `Docs`
  one small README that shows registration and expected AST shape

The important part is to keep your semantic node type as the center. Parsing, AST shape, and rendering should all flow from that one type instead of from disconnected string-processing helpers.
