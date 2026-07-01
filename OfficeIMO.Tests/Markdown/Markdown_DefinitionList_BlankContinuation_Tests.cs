using OfficeIMO.Markdown;
using MarkdigMarkdown = Markdig.Markdown;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite;

public sealed class Markdown_DefinitionList_BlankContinuation_Tests {
    [Fact]
    public void DefinitionList_BlankSeparatedIndentedParagraph_StripsContainerIndent_AndPreservesLazyTail() {
        const string markdown = """
Term
:   First paragraph

    Second paragraph
lazy continuation
""";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, CreateMarkdigDefinitionListReaderOptions());
        var definitionList = Assert.IsType<DefinitionListBlock>(Assert.Single(result.Document.Blocks));
        var definition = Assert.Single(Assert.Single(definitionList.Groups).Definitions);
        var firstParagraph = Assert.IsType<ParagraphBlock>(definition.Blocks[0]);
        var secondParagraph = Assert.IsType<ParagraphBlock>(definition.Blocks[1]);
        var syntaxGroup = Assert.Single(result.SyntaxTree.Children).Children[0];
        var definitionValue = syntaxGroup.Children.Single(child => child.Kind == MarkdownSyntaxKind.DefinitionValue);
        var paragraphSyntax = definitionValue.Children.Where(child => child.Kind == MarkdownSyntaxKind.Paragraph).ToArray();
        var written = NormalizeMarkdown(result.Document.ToMarkdown());
        var reparsed = MarkdownReader.Parse(written, CreateMarkdigDefinitionListReaderOptions());
        var office = result.Document.ToHtmlFragment(CreateMarkdigDefinitionListHtmlOptions());
        var reparsedOffice = reparsed.ToHtmlFragment(CreateMarkdigDefinitionListHtmlOptions());
        var markdig = MarkdigMarkdown.ToHtml(markdown, CreateMarkdigDefinitionListPipeline());

        Assert.Equal("First paragraph", firstParagraph.Inlines.RenderMarkdown());
        Assert.Equal("Second paragraph\nlazy continuation", secondParagraph.Inlines.RenderMarkdown());
        Assert.Equal(2, paragraphSyntax.Length);
        Assert.Equal(new MarkdownSourceSpan(2, 5, 5, 17), definitionValue.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(2, 5, 2, 19), paragraphSyntax[0].SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(4, 5, 5, 17), paragraphSyntax[1].SourceSpan);
        Assert.Equal(
            new[] {
                MarkdownSyntaxKind.InlineText,
                MarkdownSyntaxKind.InlineSoftBreak,
                MarkdownSyntaxKind.InlineText
            },
            paragraphSyntax[1].Children.Select(child => child.Kind).ToArray());
        Assert.Equal("Term\n:   First paragraph\n\n    Second paragraph\n    lazy continuation", written);
        Assert.Equal(NormalizeHtml(markdig), NormalizeHtml(office));
        Assert.Equal(NormalizeHtml(markdig), NormalizeHtml(reparsedOffice));

        var native = MarkdownNativeDocument.Parse(markdown, CreateMarkdigDefinitionListReaderOptions());
        var definitionBody = Assert.Single(native.EnumerateBlockSourceFields("definitionBody"));
        var blankLine = Assert.Single(native.EnumerateBlockSourceFields("definitionBlankLine"));
        var continuationIndent = Assert.Single(native.EnumerateBlockSourceFields("definitionContinuationIndent"));
        var nativeDefinitionList = Assert.IsType<MarkdownNativeDefinitionListBlock>(Assert.Single(native.Blocks));
        var nativeDefinition = Assert.Single(Assert.Single(nativeDefinitionList.Groups).Definitions);
        var nativeParagraphs = nativeDefinition.Children.OfType<MarkdownNativeParagraphBlock>().ToArray();

        Assert.Equal("First paragraph\n\nSecond paragraph\nlazy continuation", definitionBody.Value!.Replace("\r\n", "\n"));
        Assert.Equal(new MarkdownSourceSpan(2, 5, 5, 17), definitionBody.SourceSpan);
        Assert.Equal(string.Empty, blankLine.Value);
        Assert.Equal(new MarkdownSourceSpan(3, 1, 3, 1), blankLine.SourceSpan);
        Assert.Same(nativeDefinitionList, blankLine.Block);
        Assert.Equal(0, blankLine.Index);
        var foundBlankLine = native.FindBlockSourceFieldAtPosition(3, 1);
        Assert.NotNull(foundBlankLine);
        Assert.Equal(blankLine.Name, foundBlankLine!.Name);
        Assert.Equal(blankLine.SourceSpan, foundBlankLine.SourceSpan);
        Assert.Equal(new[] { new MarkdownSourceSpan(3, 1, 3, 1) }, nativeDefinition.BlankLineSourceSpans);
        Assert.Null(continuationIndent.Value);
        Assert.Equal(new MarkdownSourceSpan(4, 1, 4, 4), continuationIndent.SourceSpan);
        Assert.Same(nativeDefinitionList, continuationIndent.Block);
        Assert.Equal(0, continuationIndent.Index);
        var foundContinuationIndent = native.FindBlockSourceFieldAtPosition(4, 1);
        Assert.NotNull(foundContinuationIndent);
        Assert.Equal(continuationIndent.Name, foundContinuationIndent!.Name);
        Assert.Equal(continuationIndent.SourceSpan, foundContinuationIndent.SourceSpan);
        Assert.Equal(new[] { new MarkdownSourceSpan(4, 1, 4, 4) }, nativeDefinition.ContinuationIndentSourceSpans);
        Assert.Equal(new[] { "First paragraph", "Second paragraph\nlazy continuation" }, nativeParagraphs.Select(paragraph => paragraph.Text).ToArray());
        Assert.Contains(nativeParagraphs[1].InlineRuns, inline => inline.Kind == MarkdownNativeInlineKind.SoftBreak);
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);
    }

    [Fact]
    public void DefinitionList_BlankSeparatedNestedListLazyTail_PreservesListItemSoftBreak() {
        const string markdown = """
Term
:   First

    - item
lazy continuation
""";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, CreateMarkdigDefinitionListReaderOptions());
        var definitionList = Assert.IsType<DefinitionListBlock>(Assert.Single(result.Document.Blocks));
        var definition = Assert.Single(Assert.Single(definitionList.Groups).Definitions);
        var firstParagraph = Assert.IsType<ParagraphBlock>(definition.Blocks[0]);
        var nestedList = Assert.IsType<UnorderedListBlock>(definition.Blocks[1]);
        var item = Assert.Single(nestedList.Items);
        var syntaxGroup = Assert.Single(result.SyntaxTree.Children).Children[0];
        var definitionValue = syntaxGroup.Children.Single(child => child.Kind == MarkdownSyntaxKind.DefinitionValue);
        var listSyntax = definitionValue.Children.Single(child => child.Kind == MarkdownSyntaxKind.UnorderedList);
        var listItemSyntax = listSyntax.Children.Single(child => child.Kind == MarkdownSyntaxKind.ListItem);
        var itemParagraphSyntax = listItemSyntax.Children.Single(child => child.Kind == MarkdownSyntaxKind.Paragraph);
        var written = NormalizeMarkdown(result.Document.ToMarkdown());
        var reparsed = MarkdownReader.Parse(written, CreateMarkdigDefinitionListReaderOptions());
        var office = result.Document.ToHtmlFragment(CreateMarkdigDefinitionListHtmlOptions());
        var reparsedOffice = reparsed.ToHtmlFragment(CreateMarkdigDefinitionListHtmlOptions());
        var markdig = MarkdigMarkdown.ToHtml(markdown, CreateMarkdigDefinitionListPipeline());

        Assert.Equal("First", firstParagraph.Inlines.RenderMarkdown());
        Assert.Equal("item\nlazy continuation", item.Content.RenderMarkdown().Replace("\r\n", "\n"));
        Assert.Equal(
            new[] {
                MarkdownSyntaxKind.InlineText,
                MarkdownSyntaxKind.InlineSoftBreak,
                MarkdownSyntaxKind.InlineText
            },
            itemParagraphSyntax.Children.Select(child => child.Kind).ToArray());
        Assert.Equal(new MarkdownSourceSpan(2, 5, 5, 17), definitionValue.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(4, 5, 5, 17), listSyntax.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(4, 5, 5, 17), listItemSyntax.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(4, 7, 5, 17), itemParagraphSyntax.SourceSpan);
        Assert.Equal("Term\n:   First\n    - item\n      lazy continuation", written);
        Assert.Equal(NormalizeHtml(markdig), NormalizeHtml(office));
        Assert.Equal(NormalizeHtml(markdig), NormalizeHtml(reparsedOffice));

        var native = MarkdownNativeDocument.Parse(markdown, CreateMarkdigDefinitionListReaderOptions());
        var definitionBody = Assert.Single(native.EnumerateBlockSourceFields("definitionBody"));
        var blankLine = Assert.Single(native.EnumerateBlockSourceFields("definitionBlankLine"));
        var continuationIndent = Assert.Single(native.EnumerateBlockSourceFields("definitionContinuationIndent"));
        var nativeDefinitionList = Assert.IsType<MarkdownNativeDefinitionListBlock>(Assert.Single(native.Blocks));
        var nativeDefinition = Assert.Single(Assert.Single(nativeDefinitionList.Groups).Definitions);
        var nativeList = Assert.IsType<MarkdownNativeListBlock>(nativeDefinition.Children[1]);
        var nativeItem = Assert.Single(nativeList.Items);
        var nativeParagraph = Assert.IsType<MarkdownNativeParagraphBlock>(Assert.Single(nativeItem.Children));

        Assert.Equal("First\n\n- item\n  lazy continuation", definitionBody.Value!.Replace("\r\n", "\n"));
        Assert.Equal(new MarkdownSourceSpan(2, 5, 5, 17), definitionBody.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(3, 1, 3, 1), blankLine.SourceSpan);
        var foundBlankLine = native.FindBlockSourceFieldAtPosition(3, 1);
        Assert.NotNull(foundBlankLine);
        Assert.Equal(blankLine.Name, foundBlankLine!.Name);
        Assert.Equal(blankLine.SourceSpan, foundBlankLine.SourceSpan);
        Assert.Equal(new[] { new MarkdownSourceSpan(3, 1, 3, 1) }, nativeDefinition.BlankLineSourceSpans);
        Assert.Equal(new MarkdownSourceSpan(4, 1, 4, 4), continuationIndent.SourceSpan);
        var foundContinuationIndent = native.FindBlockSourceFieldAtPosition(4, 1);
        Assert.NotNull(foundContinuationIndent);
        Assert.Equal(continuationIndent.Name, foundContinuationIndent!.Name);
        Assert.Equal(continuationIndent.SourceSpan, foundContinuationIndent.SourceSpan);
        Assert.Equal(new[] { new MarkdownSourceSpan(4, 1, 4, 4) }, nativeDefinition.ContinuationIndentSourceSpans);
        Assert.Equal("item\nlazy continuation", nativeParagraph.Text.Replace("\r\n", "\n"));
        Assert.Contains(nativeParagraph.InlineRuns, inline => inline.Kind == MarkdownNativeInlineKind.SoftBreak);
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);
    }

    [Fact]
    public void DefinitionList_BlankSeparatedNestedBlockquoteLazyTail_PreservesQuoteSoftBreak() {
        const string markdown = """
Term
:   First

    > quote
lazy continuation
""";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, CreateMarkdigDefinitionListReaderOptions());
        var definitionList = Assert.IsType<DefinitionListBlock>(Assert.Single(result.Document.Blocks));
        var definition = Assert.Single(Assert.Single(definitionList.Groups).Definitions);
        var firstParagraph = Assert.IsType<ParagraphBlock>(definition.Blocks[0]);
        var quote = Assert.IsType<QuoteBlock>(definition.Blocks[1]);
        var quoteParagraph = Assert.IsType<ParagraphBlock>(Assert.Single(quote.ChildBlocks));
        var syntaxGroup = Assert.Single(result.SyntaxTree.Children).Children[0];
        var definitionValue = syntaxGroup.Children.Single(child => child.Kind == MarkdownSyntaxKind.DefinitionValue);
        var quoteSyntax = definitionValue.Children.Single(child => child.Kind == MarkdownSyntaxKind.Quote);
        var quoteParagraphSyntax = quoteSyntax.Children.Single(child => child.Kind == MarkdownSyntaxKind.Paragraph);
        var written = NormalizeMarkdown(result.Document.ToMarkdown());
        var reparsed = MarkdownReader.Parse(written, CreateMarkdigDefinitionListReaderOptions());
        var office = result.Document.ToHtmlFragment(CreateMarkdigDefinitionListHtmlOptions());
        var reparsedOffice = reparsed.ToHtmlFragment(CreateMarkdigDefinitionListHtmlOptions());
        var markdig = MarkdigMarkdown.ToHtml(markdown, CreateMarkdigDefinitionListPipeline());

        Assert.Equal("First", firstParagraph.Inlines.RenderMarkdown());
        Assert.Equal("quote\nlazy continuation", quoteParagraph.Inlines.RenderMarkdown().Replace("\r\n", "\n"));
        Assert.Equal(
            new[] {
                MarkdownSyntaxKind.InlineText,
                MarkdownSyntaxKind.InlineSoftBreak,
                MarkdownSyntaxKind.InlineText
            },
            quoteParagraphSyntax.Children.Select(child => child.Kind).ToArray());
        Assert.Equal(new MarkdownSourceSpan(2, 5, 5, 17), definitionValue.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(4, 5, 5, 17), quoteSyntax.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(4, 7, 5, 17), quoteParagraphSyntax.SourceSpan);
        Assert.Equal("Term\n:   First\n    > quote\n    > lazy continuation", written);
        Assert.Equal(NormalizeHtml(markdig), NormalizeHtml(office));
        Assert.Equal(NormalizeHtml(markdig), NormalizeHtml(reparsedOffice));

        var native = MarkdownNativeDocument.Parse(markdown, CreateMarkdigDefinitionListReaderOptions());
        var definitionBody = Assert.Single(native.EnumerateBlockSourceFields("definitionBody"));
        var blankLine = Assert.Single(native.EnumerateBlockSourceFields("definitionBlankLine"));
        var continuationIndent = Assert.Single(native.EnumerateBlockSourceFields("definitionContinuationIndent"));
        var nativeDefinitionList = Assert.IsType<MarkdownNativeDefinitionListBlock>(Assert.Single(native.Blocks));
        var nativeDefinition = Assert.Single(Assert.Single(nativeDefinitionList.Groups).Definitions);
        var nativeQuote = Assert.IsType<MarkdownNativeQuoteBlock>(nativeDefinition.Children[1]);
        var nativeParagraph = Assert.IsType<MarkdownNativeParagraphBlock>(Assert.Single(nativeQuote.Children));

        Assert.Equal("First\n\n> quote\n> lazy continuation", definitionBody.Value!.Replace("\r\n", "\n"));
        Assert.Equal(new MarkdownSourceSpan(2, 5, 5, 17), definitionBody.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(3, 1, 3, 1), blankLine.SourceSpan);
        var foundBlankLine = native.FindBlockSourceFieldAtPosition(3, 1);
        Assert.NotNull(foundBlankLine);
        Assert.Equal(blankLine.Name, foundBlankLine!.Name);
        Assert.Equal(blankLine.SourceSpan, foundBlankLine.SourceSpan);
        Assert.Equal(new[] { new MarkdownSourceSpan(3, 1, 3, 1) }, nativeDefinition.BlankLineSourceSpans);
        Assert.Equal(new MarkdownSourceSpan(4, 1, 4, 4), continuationIndent.SourceSpan);
        var foundContinuationIndent = native.FindBlockSourceFieldAtPosition(4, 1);
        Assert.NotNull(foundContinuationIndent);
        Assert.Equal(continuationIndent.Name, foundContinuationIndent!.Name);
        Assert.Equal(continuationIndent.SourceSpan, foundContinuationIndent.SourceSpan);
        Assert.Equal(new[] { new MarkdownSourceSpan(4, 1, 4, 4) }, nativeDefinition.ContinuationIndentSourceSpans);
        Assert.Equal("quote\nlazy continuation", nativeParagraph.Text.Replace("\r\n", "\n"));
        Assert.Contains(nativeParagraph.InlineRuns, inline => inline.Kind == MarkdownNativeInlineKind.SoftBreak);
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);
    }

    private static Markdig.MarkdownPipeline CreateMarkdigDefinitionListPipeline() {
        var builder = new Markdig.MarkdownPipelineBuilder();
        Markdig.MarkdownExtensions.UseDefinitionLists(builder);
        return builder.Build();
    }

    private static MarkdownReaderOptions CreateMarkdigDefinitionListReaderOptions() {
        var options = MarkdownReaderOptions.CreateCommonMarkProfile();
        options.DefinitionLists = true;
        return options;
    }

    private static HtmlOptions CreateMarkdigDefinitionListHtmlOptions() => new() {
        Style = HtmlStyle.Plain,
        CssDelivery = CssDelivery.None,
        BodyClass = null,
        AutoHeadingIdentifiers = false
    };

    private static string NormalizeHtml(string html) {
        if (string.IsNullOrWhiteSpace(html)) {
            return string.Empty;
        }

        var compact = html
            .Replace("\r\n", "\n")
            .Replace('\r', '\n')
            .Replace("> <", "><")
            .Trim();
        var sb = new System.Text.StringBuilder(compact.Length);
        bool lastWasWhitespace = false;
        for (int i = 0; i < compact.Length; i++) {
            char ch = compact[i];
            if (char.IsWhiteSpace(ch)) {
                lastWasWhitespace = true;
                continue;
            }

            if (lastWasWhitespace && sb.Length > 0 && sb[sb.Length - 1] != '>') {
                sb.Append(' ');
            }

            lastWasWhitespace = false;
            sb.Append(ch);
        }

        return sb.ToString();
    }

    private static string NormalizeMarkdown(string markdown) {
        return markdown.Replace("\r\n", "\n").TrimEnd();
    }
}
