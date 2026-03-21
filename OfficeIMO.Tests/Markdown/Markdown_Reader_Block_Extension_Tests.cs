using System.Text;
using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite;

public class Markdown_Reader_Block_Extension_Tests {
    [Fact]
    public void Parse_Uses_Delegate_Based_Block_Parser_Extension_With_Nested_Blocks() {
        var markdown = """
:::panel Ops Notes
Paragraph line

- item
:::
""";

        var document = MarkdownReader.Parse(markdown, CreateOptions());

        var panel = Assert.IsType<PanelBlock>(Assert.Single(document.Blocks));
        Assert.Equal("Ops Notes", panel.Title);
        Assert.Equal(2, panel.ChildBlocks.Count);
        Assert.IsType<ParagraphBlock>(panel.ChildBlocks[0]);
        Assert.IsType<UnorderedListBlock>(panel.ChildBlocks[1]);

        Assert.Equal(2, document.DescendantsOfType<ParagraphBlock>().Count());
        var nestedParagraph = Assert.IsType<ParagraphBlock>(panel.ChildBlocks[0]);
        Assert.Equal(new MarkdownSourceSpan(2, 1, 2, 14), nestedParagraph.SourceSpan);
    }

    [Fact]
    public void ParseWithSyntaxTree_Captures_Custom_Block_Parser_Ast_And_SourceSpans() {
        var markdown = """
:::panel Ops Notes
Paragraph line

- item
:::
""";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, CreateOptions());

        var panelSyntax = Assert.Single(result.SyntaxTree.Children);
        Assert.Equal(MarkdownSyntaxKind.Unknown, panelSyntax.Kind);
        Assert.Equal("panel-block", panelSyntax.CustomKind);
        Assert.Equal(new MarkdownSourceSpan(1, 1, 5, 3), panelSyntax.SourceSpan);

        Assert.Equal(new[] {
            MarkdownSyntaxKind.Paragraph,
            MarkdownSyntaxKind.Paragraph,
            MarkdownSyntaxKind.UnorderedList
        }, panelSyntax.Children.Select(child => child.Kind).ToArray());

        var panel = Assert.IsType<PanelBlock>(Assert.Single(result.Document.Blocks));
        Assert.Equal(new MarkdownSourceSpan(1, 1, 5, 3), panel.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(2, 1, 2, 14), ((MarkdownObject)panel.ChildBlocks[0]).SourceSpan);
        Assert.Same(panel, panelSyntax.AssociatedObject);
    }

    [Fact]
    public void Custom_Block_Can_Render_Html_With_Public_Heading_Context_Helpers() {
        var markdown = """
# Intro

:::panel Ops Notes
Paragraph line
:::

## Child
### Deep
""";

        var document = MarkdownReader.Parse(markdown, CreateOptions());
        var html = document.ToHtmlFragment(new HtmlOptions {
            Kind = HtmlKind.Fragment,
            Title = "panel-title"
        });

        Assert.Contains("data-title=\"panel-title\"", html, StringComparison.Ordinal);
        Assert.Contains("data-block-index=\"1\"", html, StringComparison.Ordinal);
        Assert.Contains("data-parent-anchor=\"intro\"", html, StringComparison.Ordinal);
        Assert.Contains("href=\"#child\"", html, StringComparison.Ordinal);
        Assert.Contains("href=\"#deep\"", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Html_Block_Render_Extension_Can_Use_Public_Body_Context_And_Override_Contextual_Block_Rendering() {
        var markdown = """
# Intro

:::panel Ops Notes
Paragraph line
:::

## Child
""";

        var document = MarkdownReader.Parse(markdown, CreateOptions());
        var htmlOptions = new HtmlOptions {
            Kind = HtmlKind.Fragment,
            Title = "override-title"
        };
        htmlOptions.BlockRenderExtensions.Add(MarkdownBlockHtmlRenderExtension.CreateContextual(
            "panel-override",
            typeof(PanelBlock),
            static (block, context) => {
                if (block is not PanelBlock panel) {
                    return null;
                }

                var blockIndex = context.GetBlockIndex(panel);
                var titleAnchor = context.GetPrecedingHeadingAnchor(blockIndex, new TocOptions {
                    Scope = TocScope.PreviousHeading,
                    IncludeTitle = true,
                    MinLevel = 2,
                    MaxLevel = 6
                });
                return $"<aside data-panel-override=\"true\" data-title=\"{System.Net.WebUtility.HtmlEncode(context.Options.Title)}\" data-block-index=\"{blockIndex}\" data-parent-anchor=\"{System.Net.WebUtility.HtmlEncode(titleAnchor)}\">{System.Net.WebUtility.HtmlEncode(panel.Title)}</aside>";
            }));

        var html = document.ToHtmlFragment(htmlOptions);

        Assert.Contains("data-panel-override=\"true\"", html, StringComparison.Ordinal);
        Assert.Contains("data-title=\"override-title\"", html, StringComparison.Ordinal);
        Assert.Contains("data-block-index=\"1\"", html, StringComparison.Ordinal);
        Assert.Contains("data-parent-anchor=\"intro\"", html, StringComparison.Ordinal);
        Assert.DoesNotContain("data-panel=\"true\"", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Html_Block_Render_Extension_Legacy_Constructor_Still_Uses_Options_And_Applies() {
        var markdown = """
:::panel Ops Notes
Paragraph line
:::
""";

        var document = MarkdownReader.Parse(markdown, CreateOptions());
        var htmlOptions = new HtmlOptions {
            Kind = HtmlKind.Fragment,
            Title = "legacy-title"
        };
        htmlOptions.BlockRenderExtensions.Add(new MarkdownBlockHtmlRenderExtension(
            "panel-legacy-override",
            typeof(PanelBlock),
            static (block, options) => {
                if (block is not PanelBlock panel) {
                    return null;
                }

                return $"<aside data-legacy-panel=\"true\" data-title=\"{System.Net.WebUtility.HtmlEncode(options.Title)}\">{System.Net.WebUtility.HtmlEncode(panel.Title)}</aside>";
            }));

        var html = document.ToHtmlFragment(htmlOptions);

        Assert.Contains("data-legacy-panel=\"true\"", html, StringComparison.Ordinal);
        Assert.Contains("data-title=\"legacy-title\"", html, StringComparison.Ordinal);
    }

    private static MarkdownReaderOptions CreateOptions() {
        var options = MarkdownReaderOptions.CreatePortableProfile();
        options.BlockParserExtensions.Add(new MarkdownBlockParserExtension(
            "panel-block",
            MarkdownBlockParserPlacement.BeforeParagraphs,
            TryParsePanelBlock));
        return options;
    }

    private static bool TryParsePanelBlock(MarkdownBlockParserContext context, out MarkdownBlockParseResult result) {
        result = default;
        const string prefix = ":::panel";
        var trimmed = context.CurrentLine.Trim();
        if (!trimmed.StartsWith(prefix, StringComparison.Ordinal)) {
            return false;
        }

        var title = trimmed.Length == prefix.Length ? string.Empty : trimmed.Substring(prefix.Length).Trim();
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

        var nestedBlocks = closingOffset > 1
            ? context.ParseNestedBlocks(1, closingOffset - 1)
            : Array.Empty<IMarkdownBlock>();
        result = new MarkdownBlockParseResult(new PanelBlock(title, nestedBlocks), closingOffset + 1);
        return true;
    }

    private sealed class PanelBlock : MarkdownBlock, IMarkdownBlock, ISyntaxMarkdownBlockWithContext, IContextualHtmlMarkdownBlock, IChildMarkdownBlockContainer {
        public PanelBlock(string title, IReadOnlyList<IMarkdownBlock> childBlocks) {
            Title = title ?? string.Empty;
            ChildBlocks = childBlocks ?? Array.Empty<IMarkdownBlock>();
        }

        public string Title { get; }
        public IReadOnlyList<IMarkdownBlock> ChildBlocks { get; }

        string IMarkdownBlock.RenderMarkdown() {
            var body = string.Join("\n\n", ChildBlocks.Select(block => block.RenderMarkdown().TrimEnd()));
            return string.IsNullOrWhiteSpace(body)
                ? $":::panel {Title}\n:::"
                : $":::panel {Title}\n{body}\n:::";
        }

        string IMarkdownBlock.RenderHtml() {
            var body = string.Concat(ChildBlocks.Select(block => block.RenderHtml()));
            return $"<section data-panel=\"true\"><h2>{System.Net.WebUtility.HtmlEncode(Title)}</h2>{body}</section>";
        }

        string IContextualHtmlMarkdownBlock.RenderHtml(MarkdownBodyRenderContext context) {
            var body = string.Concat(ChildBlocks.Select(block => block.RenderHtml()));
            var blockIndex = context.GetBlockIndex(this);
            var tocOptions = new TocOptions {
                Scope = TocScope.PreviousHeading,
                IncludeTitle = true,
                MinLevel = 2,
                MaxLevel = 6
            };
            var titleAnchor = context.GetPrecedingHeadingAnchor(blockIndex, tocOptions);
            var entries = context.BuildTocEntries(blockIndex, tocOptions, titleAnchor);
            var nav = string.Concat(entries.Select(entry =>
                $"<a href=\"#{System.Net.WebUtility.HtmlEncode(entry.Anchor)}\">{System.Net.WebUtility.HtmlEncode(entry.Text)}</a>"));
            return $"<section data-panel=\"true\" data-title=\"{System.Net.WebUtility.HtmlEncode(context.Options.Title)}\" data-block-index=\"{blockIndex}\" data-parent-anchor=\"{System.Net.WebUtility.HtmlEncode(titleAnchor)}\"><h2>{System.Net.WebUtility.HtmlEncode(Title)}</h2>{nav}{body}</section>";
        }

        public MarkdownSyntaxNode BuildSyntaxNode(MarkdownBlockSyntaxBuilderContext context, MarkdownSourceSpan? span) {
            var titleNode = context.BuildInlineContainerNode(
                MarkdownSyntaxKind.Paragraph,
                new InlineSequence().Text(Title),
                literal: Title);
            var childNodes = context.BuildChildSyntaxNodes(ChildBlocks);
            var children = new List<MarkdownSyntaxNode>(childNodes.Count + 1) { titleNode };
            for (int i = 0; i < childNodes.Count; i++) {
                children.Add(childNodes[i]);
            }

            return new MarkdownSyntaxNode(
                MarkdownSyntaxKind.Unknown,
                span ?? context.GetAggregateSpan(children),
                literal: context.NormalizeLiteralLineEndings(((IMarkdownBlock)this).RenderMarkdown()),
                children: children,
                associatedObject: this,
                customKind: "panel-block");
        }
    }
}
