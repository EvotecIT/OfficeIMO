using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite;

public class Markdown_Semantic_Fenced_Block_Tests {
    [Fact]
    public void Parse_Uses_FencedBlockExtension_To_Create_SemanticFencedBlock() {
        var options = CreateSemanticOptions("ix-chart", MarkdownSemanticKinds.Chart);
        var markdown = """
```ix-chart
{"type":"bar"}
```
_Chart caption_
""";

        var doc = OfficeIMO.Markdown.MarkdownReader.Parse(markdown, options);

        var block = Assert.IsType<SemanticFencedBlock>(Assert.Single(doc.Blocks));
        Assert.Equal(MarkdownSemanticKinds.Chart, block.SemanticKind);
        Assert.Equal("ix-chart", block.Language);
        Assert.Equal("{\"type\":\"bar\"}", block.Content);
        Assert.Equal("Chart caption", block.Caption);
    }

    [Fact]
    public void Parse_Uses_FencedBlockExtension_For_Nested_Quoted_Fences() {
        var options = CreateSemanticOptions("ix-chart", MarkdownSemanticKinds.Chart);
        var markdown = """
> ```ix-chart
> {"type":"bar"}
> ```
""";

        var doc = OfficeIMO.Markdown.MarkdownReader.Parse(markdown, options);

        var quote = Assert.IsType<QuoteBlock>(Assert.Single(doc.Blocks));
        var block = Assert.IsType<SemanticFencedBlock>(Assert.Single(quote.ChildBlocks));
        Assert.Equal(MarkdownSemanticKinds.Chart, block.SemanticKind);
        Assert.Equal("ix-chart", block.Language);
    }

    [Fact]
    public void ParseWithSyntaxTree_Captures_SemanticFencedBlock_Structure() {
        var options = CreateSemanticOptions("ix-chart", MarkdownSemanticKinds.Chart);
        var markdown = """
```ix-chart
{"type":"bar"}
```
""";

        var result = OfficeIMO.Markdown.MarkdownReader.ParseWithSyntaxTree(markdown, options);

        var block = Assert.Single(result.SyntaxTree.Children);
        Assert.Equal(MarkdownSyntaxKind.SemanticFencedBlock, block.Kind);
        Assert.Equal(5, block.Children.Count);
        Assert.Equal(MarkdownSyntaxKind.FenceSemanticKind, block.Children[0].Kind);
        Assert.Equal(MarkdownSemanticKinds.Chart, block.Children[0].Literal);
        Assert.Equal(MarkdownSyntaxKind.CodeFenceOpening, block.Children[1].Kind);
        Assert.Equal("```", block.Children[1].Literal);
        Assert.Equal(MarkdownSyntaxKind.CodeFenceInfo, block.Children[2].Kind);
        Assert.Equal("ix-chart", block.Children[2].Literal);
        Assert.Equal(MarkdownSyntaxKind.CodeContent, block.Children[3].Kind);
        Assert.Equal("{\"type\":\"bar\"}", block.Children[3].Literal);
        Assert.Equal(MarkdownSyntaxKind.CodeFenceClosing, block.Children[4].Kind);
        Assert.Equal("```", block.Children[4].Literal);
    }

    [Fact]
    public void ParseWithSyntaxTree_Assigns_SourceSpan_To_SemanticFencedBlock_ObjectModel() {
        var options = CreateSemanticOptions("ix-chart", MarkdownSemanticKinds.Chart);
        var markdown = """
```ix-chart
{"type":"bar"}
```
""";

        var result = OfficeIMO.Markdown.MarkdownReader.ParseWithSyntaxTree(markdown, options);

        var block = Assert.IsType<SemanticFencedBlock>(Assert.Single(result.Document.Blocks));
        Assert.Equal(new MarkdownSourceSpan(1, 1, 3, 3), block.SourceSpan);
    }

    [Fact]
    public void ParseWithSyntaxTree_Projects_SemanticFencedBlock_Metadata_To_GenericAttributes() {
        var options = CreateSemanticOptions("ix-chart", MarkdownSemanticKinds.Chart);
        var markdown = """
```ix-chart {#chart .wide data-panel=main pinned}
{"type":"bar"}
```
""";

        var result = OfficeIMO.Markdown.MarkdownReader.ParseWithSyntaxTree(markdown, options);

        var block = Assert.IsType<SemanticFencedBlock>(Assert.Single(result.Document.Blocks));
        Assert.Equal("chart", block.Attributes.ElementId);
        Assert.Equal(new[] { "wide" }, block.Attributes.Classes);
        Assert.Equal("main", block.Attributes.GetAttribute("data-panel"));
        Assert.Equal("true", block.Attributes.GetAttribute("pinned"));

        var syntaxBlock = Assert.Single(result.SyntaxTree.Children);
        Assert.Equal("chart", syntaxBlock.Attributes.ElementId);
        Assert.Equal(new[] { "wide" }, syntaxBlock.Attributes.Classes);
        Assert.Equal("main", syntaxBlock.Attributes.GetAttribute("data-panel"));
        Assert.Equal("true", syntaxBlock.Attributes.GetAttribute("pinned"));

        var native = MarkdownNativeDocument.Parse(markdown, options);
        var visual = Assert.IsType<MarkdownNativeVisualBlock>(Assert.Single(native.Blocks));
        var attributes = Assert.Single(native.EnumerateBlockSourceFields("attributes"));

        Assert.Same(visual, attributes.Block);
        Assert.Equal("{#chart .wide data-panel=main pinned}", attributes.Value);
        Assert.Equal(new MarkdownSourceSpan(1, 13, 1, 49), attributes.SourceSpan);
        Assert.Equal("attributes", native.FindBlockSourceFieldAtPosition(1, 13)?.Name);

        var snapshot = Assert.Single(native.ToSnapshot().Blocks);
        var snapshotAttributes = Assert.Single(snapshot.EnumerateSourceFields("attributes"));
        Assert.Equal("{#chart .wide data-panel=main pinned}", snapshot.Fields["attributes"]);
        Assert.Equal(13, snapshot.FieldSourceSpans["attributes"]!.StartColumn);
        Assert.Equal(49, snapshot.FieldSourceSpans["attributes"]!.EndColumn);
        Assert.Equal("{#chart .wide data-panel=main pinned}", snapshotAttributes.Value);
        Assert.Equal(13, snapshotAttributes.SourceSpan.StartColumn);
        Assert.Equal(49, snapshotAttributes.SourceSpan.EndColumn);
    }

    [Fact]
    public void SemanticFencedBlock_Native_SourceFields_Expose_Fence_Token_Values() {
        var options = CreateSemanticOptions("ix-chart", MarkdownSemanticKinds.Chart);
        const string markdown = "~~~~ix-chart {#chart .wide}\n{\"type\":\"bar\"}\n~~~~\n";

        var native = MarkdownNativeDocument.Parse(markdown, options);
        var visual = Assert.IsType<MarkdownNativeVisualBlock>(Assert.Single(native.Blocks));

        Assert.Equal("~~~~", visual.OpeningFence);
        Assert.Equal("~~~~", visual.ClosingFence);

        var openingFence = Assert.Single(native.EnumerateBlockSourceFields("openingFence"));
        Assert.Same(visual, openingFence.Block);
        Assert.Equal("~~~~", openingFence.Value);
        Assert.Equal(new MarkdownSourceSpan(1, 1, 1, 4), openingFence.SourceSpan);

        var infoString = Assert.Single(native.EnumerateBlockSourceFields("infoString"));
        Assert.Equal("ix-chart {#chart .wide}", infoString.Value);
        Assert.Equal(new MarkdownSourceSpan(1, 5, 1, 27), infoString.SourceSpan);

        var attributes = Assert.Single(native.EnumerateBlockSourceFields("attributes"));
        Assert.Equal("{#chart .wide}", attributes.Value);
        Assert.Equal(new MarkdownSourceSpan(1, 14, 1, 27), attributes.SourceSpan);

        var content = Assert.Single(native.EnumerateBlockSourceFields("content"));
        Assert.Equal("{\"type\":\"bar\"}", content.Value);
        Assert.Equal(new MarkdownSourceSpan(2, 1, 2, 14), content.SourceSpan);

        var closingFence = Assert.Single(native.EnumerateBlockSourceFields("closingFence"));
        Assert.Equal("~~~~", closingFence.Value);
        Assert.Equal(new MarkdownSourceSpan(3, 1, 3, 4), closingFence.SourceSpan);

        Assert.Equal("openingFence", native.FindBlockSourceFieldAtPosition(1, 2)?.Name);
        Assert.Equal("closingFence", native.FindBlockSourceFieldAtPosition(3, 3)?.Name);

        var snapshot = Assert.Single(native.ToSnapshot().Blocks);
        Assert.Equal("~~~~", snapshot.Fields["openingFence"]);
        Assert.Equal("~~~~", snapshot.Fields["closingFence"]);
        Assert.Contains(snapshot.SourceFields, field =>
            field.Name == "openingFence"
            && field.Value == "~~~~"
            && field.SourceSpan.StartColumn == 1
            && field.SourceSpan.EndColumn == 4);
        Assert.Contains(snapshot.SourceFields, field =>
            field.Name == "closingFence"
            && field.Value == "~~~~"
            && field.SourceSpan.StartLine == 3
            && field.SourceSpan.StartColumn == 1
            && field.SourceSpan.EndColumn == 4);

        var openingEdited = native.CreateReplaceEdit(openingFence, "```").Apply(native.SourceMarkdown);
        Assert.StartsWith("```ix-chart {#chart .wide}", openingEdited.Replace("\r\n", "\n"), System.StringComparison.Ordinal);

        var closingEdited = native.CreateReplaceEdit(closingFence, "```").Apply(native.SourceMarkdown);
        Assert.Contains("\n```\n", closingEdited.Replace("\r\n", "\n"), System.StringComparison.Ordinal);
    }

    [Fact]
    public void Native_SourceFields_Target_Tab_Padded_Fence_Info_String() {
        const string markdown = "```\tcsharp\nbody\n```\n";

        var native = MarkdownNativeDocument.Parse(markdown);
        var infoString = Assert.Single(native.EnumerateBlockSourceFields("infoString"));

        Assert.Equal("csharp", infoString.Value);
        Assert.Equal(new MarkdownSourceSpan(1, 5, 1, 10), infoString.SourceSpan);
        Assert.Equal("```\tfsharp\nbody\n```\n", native.CreateReplaceEdit(infoString, "fsharp").Apply(native.SourceMarkdown));
    }

    [Fact]
    public void SemanticFencedBlock_RenderHtml_Projects_GenericAttributes_To_Default_Pre() {
        var options = CreateSemanticOptions("ix-chart", MarkdownSemanticKinds.Chart);
        var markdown = """
```ix-chart {#chart .wide data-panel=main pinned}
{"type":"bar"}
```
""";

        var document = OfficeIMO.Markdown.MarkdownReader.Parse(markdown, options);
        var html = document.ToHtmlFragment(new HtmlOptions {
            Style = HtmlStyle.Plain,
            CssDelivery = CssDelivery.None,
            BodyClass = null
        });

        Assert.Contains("<pre id=\"chart\" class=\"wide\" data-panel=\"main\" pinned=\"true\"><code class=\"language-ix-chart\">", html, StringComparison.Ordinal);
    }

    [Fact]
    public void ParseWithSyntaxTree_Uses_Custom_Fenced_Block_Syntax_Node_For_External_Block_Extensions() {
        var options = new MarkdownReaderOptions();
        options.FencedBlockExtensions.Add(new MarkdownFencedBlockExtension(
            "Custom AST block",
            new[] { "ix-custom" },
            context => new CustomSyntaxFencedBlock(context.Language, context.Content)));
        var markdown = """
```ix-custom {#custom-panel .wide pinned}
hello
```
""";

        var result = OfficeIMO.Markdown.MarkdownReader.ParseWithSyntaxTree(markdown, options);

        var syntaxBlock = Assert.Single(result.SyntaxTree.Children);
        Assert.Equal(MarkdownSyntaxKind.Unknown, syntaxBlock.Kind);
        Assert.Equal("custom-fenced-block", syntaxBlock.CustomKind);
        Assert.Equal("```ix-custom\nhello\n```", syntaxBlock.Literal);
        Assert.Equal(new[] {
            MarkdownSyntaxKind.CodeFenceInfo,
            MarkdownSyntaxKind.Paragraph
        }, syntaxBlock.Children.Select(child => child.Kind).ToArray());
        Assert.Equal("hello", syntaxBlock.Children[1].Literal);
        Assert.Equal("custom-panel", syntaxBlock.Attributes.ElementId);
        Assert.Equal(new[] { "wide" }, syntaxBlock.Attributes.Classes);
        Assert.Equal("true", syntaxBlock.Attributes.GetAttribute("pinned"));

        var block = Assert.IsType<CustomSyntaxFencedBlock>(Assert.Single(result.Document.Blocks));
        Assert.Equal(new MarkdownSourceSpan(1, 1, 3, 3), block.SourceSpan);
        Assert.Same(block, syntaxBlock.AssociatedObject);
        Assert.Equal("custom-panel", block.Attributes.ElementId);
        Assert.Equal(new[] { "wide" }, block.Attributes.Classes);
        Assert.Equal("true", block.Attributes.GetAttribute("pinned"));
    }

    [Fact]
    public void Custom_Fenced_Block_Can_Render_Html_With_Public_Body_Render_Context() {
        var options = new MarkdownReaderOptions();
        options.FencedBlockExtensions.Add(new MarkdownFencedBlockExtension(
            "Custom AST block",
            new[] { "ix-custom" },
            context => new CustomSyntaxFencedBlock(context.Language, context.Content)));
        var markdown = """
```ix-custom
hello
```
""";

        var document = OfficeIMO.Markdown.MarkdownReader.Parse(markdown, options);
        var html = document.ToHtmlFragment(new HtmlOptions {
            Kind = HtmlKind.Fragment,
            Title = "custom-title"
        });

        Assert.Contains("data-title=\"custom-title\"", html, StringComparison.Ordinal);
        Assert.Contains("data-block-count=\"1\"", html, StringComparison.Ordinal);
        Assert.Contains(">hello<", html, StringComparison.Ordinal);
    }

    [Fact]
    public void SemanticFencedBlock_RenderHtml_Uses_SemanticRenderer_And_RoundTrips_Markdown() {
        var block = new SemanticFencedBlock(MarkdownSemanticKinds.Chart, "ix-chart", "{\"type\":\"bar\"}", "Chart caption");
        var doc = MarkdownDoc.Create().Add(block);
        var html = doc.ToHtmlFragment(new HtmlOptions {
            Kind = HtmlKind.Fragment,
            SemanticFencedBlockHtmlRenderer = static (semanticBlock, _) =>
                $"<div class=\"semantic-block\" data-kind=\"{semanticBlock.SemanticKind}\" data-language=\"{semanticBlock.Language}\"></div>"
        });

        Assert.Contains("class=\"semantic-block\"", html, StringComparison.Ordinal);
        Assert.Contains($"data-kind=\"{MarkdownSemanticKinds.Chart}\"", html, StringComparison.Ordinal);
        Assert.Contains("data-language=\"ix-chart\"", html, StringComparison.Ordinal);
        var markdown = ((IMarkdownBlock)block).RenderMarkdown().Replace("\r\n", "\n");
        Assert.Equal("```ix-chart\n{\"type\":\"bar\"}\n```\n_Chart caption_", markdown);
    }

    [Fact]
    public void SemanticFencedBlock_RenderHtml_Falls_Back_To_CodeBlockHtmlRenderer_When_Needed() {
        var block = new SemanticFencedBlock("note", "ix-note {#note .wide pinned}", "hello");
        var doc = MarkdownDoc.Create().Add(block);
        var html = doc.ToHtmlFragment(new HtmlOptions {
            Kind = HtmlKind.Fragment,
            CodeBlockHtmlRenderer = static (codeBlock, _) =>
                $"<aside class=\"code-fallback {string.Join(" ", codeBlock.Attributes.Classes)}\" id=\"{codeBlock.Attributes.ElementId}\" data-pinned=\"{codeBlock.Attributes.GetAttribute("pinned")}\" data-language=\"{codeBlock.Language}\">{System.Net.WebUtility.HtmlEncode(codeBlock.Content)}</aside>"
        });

        Assert.Contains("class=\"code-fallback wide\"", html, StringComparison.Ordinal);
        Assert.Contains("id=\"note\"", html, StringComparison.Ordinal);
        Assert.Contains("data-pinned=\"true\"", html, StringComparison.Ordinal);
        Assert.Contains("data-language=\"ix-note\"", html, StringComparison.Ordinal);
        Assert.Contains(">hello<", html, StringComparison.Ordinal);
    }

    private static MarkdownReaderOptions CreateSemanticOptions(string language, string semanticKind) {
        var options = new MarkdownReaderOptions();
        options.FencedBlockExtensions.Add(new MarkdownFencedBlockExtension(
            "Semantic AST",
            new[] { language },
            context => new SemanticFencedBlock(semanticKind, context.InfoString, context.Content, context.Caption)));
        return options;
    }

    private sealed class CustomSyntaxFencedBlock(string language, string content) : MarkdownBlock, IMarkdownBlock, ISyntaxMarkdownBlockWithContext, IContextualHtmlMarkdownBlock {
        public string Language { get; } = language ?? string.Empty;
        public string Content { get; } = content ?? string.Empty;

        string IMarkdownBlock.RenderMarkdown() => $"```{Language}\n{Content}\n```";

        string IMarkdownBlock.RenderHtml() =>
            $"<pre><code class=\"language-{System.Net.WebUtility.HtmlEncode(Language)}\">{System.Net.WebUtility.HtmlEncode(Content)}</code></pre>";

        string IContextualHtmlMarkdownBlock.RenderHtml(MarkdownBodyRenderContext context) =>
            $"<div data-custom-block=\"true\" data-title=\"{System.Net.WebUtility.HtmlEncode(context.Options.Title)}\" data-block-count=\"{context.Blocks.Count}\">{System.Net.WebUtility.HtmlEncode(Content)}</div>";

        MarkdownSyntaxNode ISyntaxMarkdownBlockWithContext.BuildSyntaxNode(MarkdownBlockSyntaxBuilderContext context, MarkdownSourceSpan? span) {
            var contentInlines = new InlineSequence().Text(Content);
            var children = new[] {
                new MarkdownSyntaxNode(MarkdownSyntaxKind.CodeFenceInfo, literal: Language),
                context.BuildInlineContainerNode(MarkdownSyntaxKind.Paragraph, contentInlines, literal: Content)
            };

            return new MarkdownSyntaxNode(
                MarkdownSyntaxKind.Unknown,
                span ?? context.GetAggregateSpan(children),
                literal: context.NormalizeLiteralLineEndings(((IMarkdownBlock)this).RenderMarkdown()),
                children: children,
                associatedObject: this,
                customKind: "custom-fenced-block");
        }
    }
}
