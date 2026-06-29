using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite;

public class Markdown_GenericAttributes_Syntax_Tests {
    [Fact]
    public void ParseWithSyntaxTree_Captures_Block_GenericAttribute_Tokens() {
        const string markdown = "# Heading {#title .hero}\n\nAlpha paragraph {#intro .lead}\n";
        var options = new MarkdownReaderOptions {
            GenericAttributes = true,
            PreserveTrivia = true
        };

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, options);

        MarkdownInvariantAssert.SyntaxTreeIsWellFormed(result.FinalSyntaxTree);
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);

        var heading = result.FinalSyntaxTree.Children[0];
        var paragraph = result.FinalSyntaxTree.Children[1];
        var headingAttributes = Assert.Single(heading.Children, node => node.Kind == MarkdownSyntaxKind.GenericAttributeBlock);
        var paragraphAttributes = Assert.Single(paragraph.Children, node => node.Kind == MarkdownSyntaxKind.GenericAttributeBlock);

        Assert.Equal("{#title .hero}", headingAttributes.Literal);
        Assert.Equal(new MarkdownSourceSpan(1, 11, 1, 24), headingAttributes.SourceSpan);
        Assert.True(heading.SourceSpan!.Value.Contains(headingAttributes.SourceSpan!.Value));

        Assert.Equal("{#intro .lead}", paragraphAttributes.Literal);
        Assert.Equal(new MarkdownSourceSpan(3, 17, 3, 30), paragraphAttributes.SourceSpan);
        Assert.True(paragraph.SourceSpan!.Value.Contains(paragraphAttributes.SourceSpan!.Value));

        Assert.Equal(MarkdownSyntaxKind.GenericAttributeBlock, result.FindDeepestFinalNodeAtPosition(1, 20)!.Kind);
        Assert.Equal(MarkdownSyntaxKind.GenericAttributeBlock, result.FindDeepestFinalNodeAtPosition(3, 20)!.Kind);

        Assert.True(result.TryCreateOriginalSourceSlice(headingAttributes, out var headingSlice));
        Assert.Equal("{#title .hero}", headingSlice.Text);
        Assert.True(result.TryCreateOriginalSourceSlice(paragraphAttributes, out var paragraphSlice));
        Assert.Equal("{#intro .lead}", paragraphSlice.Text);
    }

    [Fact]
    public void Paragraph_GenericAttributes_Preserve_Consumed_Separator_Whitespace() {
        const string markdown = "Alpha paragraph  {#intro .lead}\n";
        var options = new MarkdownReaderOptions {
            GenericAttributes = true
        };

        var document = MarkdownReader.Parse(markdown, options);
        var paragraph = Assert.IsType<ParagraphBlock>(Assert.Single(document.Blocks));

        Assert.Equal("  ", paragraph.GenericAttributeConsumedWhitespace);
        Assert.Equal("Alpha paragraph  {#intro .lead}", ((IMarkdownBlock)paragraph).RenderMarkdown());
        Assert.Equal(
            "<p id=\"intro\" class=\"lead\">Alpha paragraph  </p>",
            document.ToHtmlFragment(new HtmlOptions {
                Style = HtmlStyle.Plain,
                CssDelivery = CssDelivery.None,
                BodyClass = null,
                EscapeNonAsciiText = false
            }));
    }

    [Fact]
    public void ParseWithSyntaxTree_Captures_Inline_GenericAttribute_Tokens_Without_Duplicating_Native_Metadata() {
        const string markdown = "See [docs](old.md){#docs .primary} now\n";
        var options = new MarkdownReaderOptions {
            GenericAttributes = true,
            PreserveTrivia = true
        };

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, options);

        MarkdownInvariantAssert.SyntaxTreeIsWellFormed(result.FinalSyntaxTree);
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);

        var link = Assert.Single(result.FinalSyntaxTree.Descendants(), node => node.Kind == MarkdownSyntaxKind.InlineLink);
        var attributes = Assert.Single(link.Children, node => node.Kind == MarkdownSyntaxKind.GenericAttributeBlock);

        Assert.Equal("{#docs .primary}", attributes.Literal);
        Assert.Equal(new MarkdownSourceSpan(1, 19, 1, 34), attributes.SourceSpan);
        Assert.True(link.SourceSpan!.Value.Contains(attributes.SourceSpan!.Value));
        Assert.Equal(MarkdownSyntaxKind.GenericAttributeBlock, result.FindDeepestFinalNodeAtPosition(1, 23)!.Kind);

        Assert.True(result.TryCreateOriginalSourceSlice(attributes, out var slice));
        Assert.Equal("{#docs .primary}", slice.Text);

        var native = MarkdownNativeDocument.Parse(markdown, options);
        var nativeParagraph = Assert.IsType<MarkdownNativeParagraphBlock>(Assert.Single(native.Blocks));
        var nativeLink = Assert.Single(nativeParagraph.InlineRuns, inline => inline.Kind == MarkdownNativeInlineKind.Link);
        var nativeAttributes = Assert.Single(nativeLink.Metadata, metadata => metadata.Name == "attributes");

        Assert.Equal("{#docs .primary}", nativeAttributes.Value);
        Assert.Equal(new MarkdownSourceSpan(1, 19, 1, 34), nativeAttributes.SourceSpan);
    }

    [Fact]
    public void FootnoteReference_GenericAttributes_Are_Consumed_Without_Metadata() {
        const string markdown = "See note[^a]{#ref .wide}\n\n[^a]: Footnote\n";
        var options = new MarkdownReaderOptions {
            Footnotes = true,
            GenericAttributes = true,
            PreserveTrivia = true
        };

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, options);

        MarkdownInvariantAssert.SyntaxTreeIsWellFormed(result.FinalSyntaxTree);
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);
        Assert.DoesNotContain(
            result.FinalSyntaxTree.Descendants(),
            node => node.Kind == MarkdownSyntaxKind.GenericAttributeBlock);

        var native = MarkdownNativeDocument.Parse(markdown, options);
        var paragraph = Assert.IsType<MarkdownNativeParagraphBlock>(native.Blocks[0]);
        var footnote = Assert.Single(paragraph.InlineRuns, inline => inline.Kind == MarkdownNativeInlineKind.FootnoteRef);

        Assert.DoesNotContain(footnote.Metadata, metadata => metadata.Name == "attributes");
        Assert.DoesNotContain(
            "{#ref .wide}",
            result.Document.ToHtmlFragment(new HtmlOptions {
                Style = HtmlStyle.Plain,
                CssDelivery = CssDelivery.None,
                BodyClass = null,
                GitHubFootnoteHtml = true
            }),
            StringComparison.Ordinal);
    }

    [Fact]
    public void ParseWithSyntaxTree_Keeps_Blockquote_Block_GenericAttributes_Literal() {
        const string markdown = "> quote {#q .lead}\n> # Heading {#h .wide}\n";
        var options = new MarkdownReaderOptions {
            GenericAttributes = true,
            PreserveTrivia = true
        };

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, options);

        MarkdownInvariantAssert.SyntaxTreeIsWellFormed(result.FinalSyntaxTree);
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);

        Assert.DoesNotContain(
            result.FinalSyntaxTree.Descendants(),
            node => node.Kind == MarkdownSyntaxKind.GenericAttributeBlock);

        var html = result.Document.ToHtmlFragment(new HtmlOptions {
            Style = HtmlStyle.Plain,
            CssDelivery = CssDelivery.None,
            BodyClass = null,
            EscapeNonAsciiText = false
        });

        Assert.Contains("quote {#q .lead}", html, StringComparison.Ordinal);
        Assert.Contains("Heading {#h .wide}", html, StringComparison.Ordinal);
        Assert.DoesNotContain("id=\"q\"", html, StringComparison.Ordinal);
        Assert.DoesNotContain("id=\"h\"", html, StringComparison.Ordinal);

        var native = MarkdownNativeDocument.Parse(markdown, options);
        Assert.Empty(native.EnumerateBlockSourceFields("attributes"));
    }

    [Fact]
    public void ParseWithSyntaxTree_Captures_ListItem_GenericAttribute_Tokens() {
        const string markdown = "- item {#li .selected}\n";
        var options = new MarkdownReaderOptions {
            GenericAttributes = true,
            PreserveTrivia = true
        };

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, options);

        MarkdownInvariantAssert.SyntaxTreeIsWellFormed(result.FinalSyntaxTree);
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);

        var listItem = Assert.Single(result.FinalSyntaxTree.Descendants(), node => node.Kind == MarkdownSyntaxKind.ListItem);
        var attributes = Assert.Single(listItem.Children, node => node.Kind == MarkdownSyntaxKind.GenericAttributeBlock);

        Assert.Equal("{#li .selected}", attributes.Literal);
        Assert.Equal(new MarkdownSourceSpan(1, 8, 1, 22), attributes.SourceSpan);
        Assert.True(listItem.SourceSpan!.Value.Contains(attributes.SourceSpan!.Value));
        Assert.Equal(MarkdownSyntaxKind.GenericAttributeBlock, result.FindDeepestFinalNodeAtPosition(1, 12)!.Kind);

        Assert.True(result.TryCreateOriginalSourceSlice(attributes, out var slice));
        Assert.Equal("{#li .selected}", slice.Text);
    }
}
