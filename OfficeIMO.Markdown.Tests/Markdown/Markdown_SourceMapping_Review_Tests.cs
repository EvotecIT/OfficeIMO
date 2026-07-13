using System;
using System.Linq;
using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite;

public sealed class Markdown_SourceMapping_Review_Tests {
    [Fact]
    public void DetailsSummary_Inlines_Are_SourceBacked() {
        const string markdown = "<details><summary>[docs](url)</summary>\nbody\n</details>";
        var native = MarkdownNativeDocument.Parse(markdown, new MarkdownReaderOptions {
            PreserveTrivia = true
        });

        var details = Assert.IsType<MarkdownNativeDetailsBlock>(Assert.Single(native.Blocks));
        var link = Assert.Single(details.SummaryInlineRuns, inline => inline.Kind == MarkdownNativeInlineKind.Link);

        Assert.Equal(new MarkdownSourceSpan(1, 19, 1, 29), link.SourceSpan);
        Assert.True(native.TryCreateOriginalSourceSlice(link, out var slice));
        Assert.Equal("[docs](url)", slice.Text);
    }

    [Fact]
    public void InlineGenericAttributes_Do_Not_Extend_Native_Content_SourceSpan() {
        var options = MarkdownReaderOptions.CreatePortableProfile();
        options.GenericAttributes = true;
        options.PreserveTrivia = true;
        var native = MarkdownNativeDocument.Parse("**hot**{#id}\n", options);

        var paragraph = Assert.IsType<MarkdownNativeParagraphBlock>(Assert.Single(native.Blocks));
        var strong = Assert.Single(paragraph.InlineRuns, inline => inline.Kind == MarkdownNativeInlineKind.Strong);

        Assert.Equal(new MarkdownSourceSpan(1, 3, 1, 5), strong.SourceSpan);
        Assert.True(native.TryCreateOriginalSourceSlice(strong, out var slice));
        Assert.Equal("hot", slice.Text);
        Assert.Equal("**warm**{#id}\n", native.CreateReplaceEdit(strong, "warm").Apply(native.SourceMarkdown));
    }

    [Fact]
    public void UnterminatedHtmlComment_Does_Not_Expose_Synthetic_Marker_SourceFields() {
        const string markdown = "<!-- unfinished";

        var result = OfficeIMO.Markdown.MarkdownReader.ParseWithSyntaxTree(markdown, new MarkdownReaderOptions {
            PreserveTrivia = true
        });
        var native = MarkdownNativeDocument.Parse(markdown, new MarkdownReaderOptions {
            PreserveTrivia = true
        });

        var comment = Assert.Single(result.SyntaxTree.Children);
        Assert.Equal(MarkdownSyntaxKind.HtmlComment, comment.Kind);
        Assert.Empty(comment.Children);

        var html = Assert.IsType<MarkdownNativeHtmlBlock>(Assert.Single(native.Blocks));
        Assert.Null(html.OpeningMarkerSourceSpan);
        Assert.Null(html.BodySourceSpan);
        Assert.Null(html.ClosingMarkerSourceSpan);
        Assert.DoesNotContain(html.EnumerateSourceFields(), field => field.Name.StartsWith("htmlComment", StringComparison.Ordinal));
    }
}
