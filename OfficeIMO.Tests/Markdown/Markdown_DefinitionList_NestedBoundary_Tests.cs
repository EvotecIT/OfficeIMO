using OfficeIMO.Markdown;
using MarkdigMarkdown = Markdig.Markdown;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite;

public sealed class Markdown_DefinitionList_NestedBoundary_Tests {
    [Fact]
    public void DefinitionList_NestedBlockquoteBody_Merges_UnindentedBlockquoteContinuation() {
        const string markdown = """
Term
:   First
    > quote
> sibling quote
""";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, CreateMarkdigDefinitionListReaderOptions());
        var definitionList = Assert.IsType<DefinitionListBlock>(Assert.Single(result.Document.Blocks));
        var definition = Assert.Single(Assert.Single(definitionList.Groups).Definitions);
        var paragraph = Assert.IsType<ParagraphBlock>(definition.Blocks[0]);
        var quote = Assert.IsType<QuoteBlock>(definition.Blocks[1]);
        var quoteParagraph = Assert.IsType<ParagraphBlock>(Assert.Single(quote.ChildBlocks));
        var syntaxGroup = result.SyntaxTree.Children[0].Children[0];
        var definitionValue = syntaxGroup.Children.Single(child => child.Kind == MarkdownSyntaxKind.DefinitionValue);
        var written = NormalizeMarkdown(result.Document.ToMarkdown());
        var reparsed = MarkdownReader.Parse(written, CreateMarkdigDefinitionListReaderOptions());
        var office = result.Document.ToHtmlFragment(CreateMarkdigDefinitionListHtmlOptions());
        var reparsedOffice = reparsed.ToHtmlFragment(CreateMarkdigDefinitionListHtmlOptions());
        var markdig = MarkdigMarkdown.ToHtml(markdown, CreateMarkdigDefinitionListPipeline());

        Assert.Equal("First", paragraph.Inlines.RenderMarkdown());
        Assert.Equal("quote sibling quote", quoteParagraph.Inlines.RenderMarkdown());
        Assert.Equal(
            new[] {
                MarkdownSyntaxKind.Paragraph,
                MarkdownSyntaxKind.Quote
            },
            definitionValue.Children.Select(child => child.Kind).ToArray());
        Assert.Equal("Term\n:   First\n    > quote sibling quote", written);
        Assert.Equal(NormalizeHtml(markdig), NormalizeHtml(office));
        Assert.Equal(NormalizeHtml(markdig), NormalizeHtml(reparsedOffice));

        var native = MarkdownNativeDocument.Parse(markdown, CreateMarkdigDefinitionListReaderOptions());
        var definitionBody = Assert.Single(native.EnumerateBlockSourceFields("definitionBody"));

        Assert.Equal("First\n\n> quote sibling quote", definitionBody.Value!.Replace("\r\n", "\n"));
        Assert.Equal(new MarkdownSourceSpan(2, 5, 4, 15), definitionBody.SourceSpan);
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);
    }

    [Fact]
    public void DefinitionList_NestedBlockquoteBody_StopsBefore_UnindentedFencedCode() {
        const string markdown = """
Term
:   First
    > quote
```csharp
code
```
""";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, CreateMarkdigDefinitionListReaderOptions());
        Assert.Equal(2, result.Document.Blocks.Count);
        var definitionList = Assert.IsType<DefinitionListBlock>(result.Document.Blocks[0]);
        var trailingCode = Assert.IsType<CodeBlock>(result.Document.Blocks[1]);
        var definition = Assert.Single(Assert.Single(definitionList.Groups).Definitions);
        var quote = Assert.IsType<QuoteBlock>(definition.Blocks[1]);
        var syntaxGroup = result.SyntaxTree.Children[0].Children[0];
        var definitionValue = syntaxGroup.Children.Single(child => child.Kind == MarkdownSyntaxKind.DefinitionValue);
        var written = NormalizeMarkdown(result.Document.ToMarkdown());
        var reparsed = MarkdownReader.Parse(written, CreateMarkdigDefinitionListReaderOptions());
        var office = result.Document.ToHtmlFragment(CreateMarkdigDefinitionListHtmlOptions());
        var reparsedOffice = reparsed.ToHtmlFragment(CreateMarkdigDefinitionListHtmlOptions());
        var markdig = MarkdigMarkdown.ToHtml(markdown, CreateMarkdigDefinitionListPipeline());

        Assert.Equal("quote", Assert.IsType<ParagraphBlock>(Assert.Single(quote.ChildBlocks)).Inlines.RenderMarkdown());
        Assert.Equal("csharp", trailingCode.Language);
        Assert.Equal("code", trailingCode.Content);
        Assert.Equal(
            new[] {
                MarkdownSyntaxKind.Paragraph,
                MarkdownSyntaxKind.Quote
            },
            definitionValue.Children.Select(child => child.Kind).ToArray());
        Assert.Equal("Term\n:   First\n    > quote\n\n```csharp\ncode\n```", written);
        Assert.Equal(NormalizeHtml(markdig), NormalizeHtml(office));
        Assert.Equal(NormalizeHtml(markdig), NormalizeHtml(reparsedOffice));

        var native = MarkdownNativeDocument.Parse(markdown, CreateMarkdigDefinitionListReaderOptions());
        var definitionBody = Assert.Single(native.EnumerateBlockSourceFields("definitionBody"));

        Assert.Equal("First\n\n> quote", definitionBody.Value!.Replace("\r\n", "\n"));
        Assert.Equal(new MarkdownSourceSpan(2, 5, 3, 11), definitionBody.SourceSpan);
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);
    }

    [Fact]
    public void DefinitionList_NestedListBody_StopsBefore_UnindentedHtmlBlock() {
        const string markdown = """
Term
:   First
    - item
<div>
html
</div>
""";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, CreateMarkdigDefinitionListReaderOptions());
        Assert.Equal(2, result.Document.Blocks.Count);
        var definitionList = Assert.IsType<DefinitionListBlock>(result.Document.Blocks[0]);
        var trailingHtml = Assert.IsType<HtmlRawBlock>(result.Document.Blocks[1]);
        var definition = Assert.Single(Assert.Single(definitionList.Groups).Definitions);
        var nestedList = Assert.IsType<UnorderedListBlock>(definition.Blocks[1]);
        var syntaxGroup = result.SyntaxTree.Children[0].Children[0];
        var definitionValue = syntaxGroup.Children.Single(child => child.Kind == MarkdownSyntaxKind.DefinitionValue);
        var written = NormalizeMarkdown(result.Document.ToMarkdown());
        var reparsed = MarkdownReader.Parse(written, CreateMarkdigDefinitionListReaderOptions());
        var office = result.Document.ToHtmlFragment(CreateMarkdigDefinitionListHtmlOptions());
        var reparsedOffice = reparsed.ToHtmlFragment(CreateMarkdigDefinitionListHtmlOptions());
        var markdig = MarkdigMarkdown.ToHtml(markdown, CreateMarkdigDefinitionListPipeline());

        Assert.Equal("item", Assert.Single(nestedList.Items).Content.RenderMarkdown());
        Assert.Equal("<div>\nhtml\n</div>", trailingHtml.Html.Replace("\r\n", "\n"));
        Assert.Equal(
            new[] {
                MarkdownSyntaxKind.Paragraph,
                MarkdownSyntaxKind.UnorderedList
            },
            definitionValue.Children.Select(child => child.Kind).ToArray());
        Assert.Equal("Term\n:   First\n    - item\n\n<div>\nhtml\n</div>", written);
        Assert.Equal(NormalizeHtml(markdig), NormalizeHtml(office));
        Assert.Equal(NormalizeHtml(markdig), NormalizeHtml(reparsedOffice));

        var native = MarkdownNativeDocument.Parse(markdown, CreateMarkdigDefinitionListReaderOptions());
        var definitionBody = Assert.Single(native.EnumerateBlockSourceFields("definitionBody"));

        Assert.Equal("First\n\n- item", definitionBody.Value!.Replace("\r\n", "\n"));
        Assert.Equal(new MarkdownSourceSpan(2, 5, 3, 10), definitionBody.SourceSpan);
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);
    }

    [Fact]
    public void DefinitionList_NestedBlockquoteLazyReferenceDefinitionLookingLine_StaysLiteral() {
        const string markdown = """
Term
:   First
    > quote
[ref]: https://example.com
""";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, CreateMarkdigDefinitionListReaderOptions());
        var definitionList = Assert.IsType<DefinitionListBlock>(Assert.Single(result.Document.Blocks));
        var definition = Assert.Single(Assert.Single(definitionList.Groups).Definitions);
        var quote = Assert.IsType<QuoteBlock>(definition.Blocks[1]);
        var quoteParagraph = Assert.IsType<ParagraphBlock>(Assert.Single(quote.ChildBlocks));
        var written = NormalizeMarkdown(result.Document.ToMarkdown());
        var reparsed = MarkdownReader.Parse(written, CreateMarkdigDefinitionListReaderOptions());
        var office = result.Document.ToHtmlFragment(CreateMarkdigDefinitionListHtmlOptions());
        var reparsedOffice = reparsed.ToHtmlFragment(CreateMarkdigDefinitionListHtmlOptions());
        var markdig = MarkdigMarkdown.ToHtml(markdown, CreateMarkdigDefinitionListPipeline());

        Assert.DoesNotContain(quoteParagraph.Inlines.Nodes, inline => inline is LinkInline);
        Assert.Equal("quote\n\\[ref\\]: https://example.com", quoteParagraph.Inlines.RenderMarkdown().Replace("\r\n", "\n"));
        Assert.Equal("Term\n:   First\n    > quote\n    > \\[ref\\]: https://example.com", written);
        Assert.Equal(NormalizeHtml(markdig), NormalizeHtml(office));
        Assert.Equal(NormalizeHtml(markdig), NormalizeHtml(reparsedOffice));

        var native = MarkdownNativeDocument.Parse(markdown, CreateMarkdigDefinitionListReaderOptions());
        var definitionBody = Assert.Single(native.EnumerateBlockSourceFields("definitionBody"));

        Assert.Equal("First\n\n> quote\n> \\[ref\\]: https://example.com", definitionBody.Value!.Replace("\r\n", "\n"));
        Assert.Equal(new MarkdownSourceSpan(2, 5, 4, 26), definitionBody.SourceSpan);
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);
    }

    [Fact]
    public void DefinitionList_NestedListLazyOrderedStartTwo_StaysInsideDefinition_AsOrderedChildList() {
        const string markdown = """
Term
:   First
    - item
2. sibling
""";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, CreateMarkdigDefinitionListReaderOptions());
        var definition = Assert.Single(Assert.Single(Assert.IsType<DefinitionListBlock>(Assert.Single(result.Document.Blocks)).Groups).Definitions);
        var unordered = Assert.IsType<UnorderedListBlock>(definition.Blocks[1]);
        var ordered = Assert.IsType<OrderedListBlock>(definition.Blocks[2]);
        var definitionValue = result.SyntaxTree.Children[0].Children[0].Children.Single(child => child.Kind == MarkdownSyntaxKind.DefinitionValue);
        var written = NormalizeMarkdown(result.Document.ToMarkdown());
        var reparsedOffice = MarkdownReader.Parse(written, CreateMarkdigDefinitionListReaderOptions()).ToHtmlFragment(CreateMarkdigDefinitionListHtmlOptions());
        var office = result.Document.ToHtmlFragment(CreateMarkdigDefinitionListHtmlOptions());
        var markdig = MarkdigMarkdown.ToHtml(markdown, CreateMarkdigDefinitionListPipeline());

        Assert.Equal("item", Assert.Single(unordered.Items).Content.RenderMarkdown());
        Assert.Equal(2, ordered.Start);
        Assert.Equal("sibling", Assert.Single(ordered.Items).Content.RenderMarkdown());
        Assert.Equal(
            new[] {
                MarkdownSyntaxKind.Paragraph,
                MarkdownSyntaxKind.UnorderedList,
                MarkdownSyntaxKind.OrderedList
            },
            definitionValue.Children.Select(child => child.Kind).ToArray());
        Assert.Equal("Term\n:   First\n    - item\n\n    2. sibling", written);
        Assert.Equal(NormalizeHtml(markdig), NormalizeHtml(office));
        Assert.Equal(NormalizeHtml(markdig), NormalizeHtml(reparsedOffice));
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);
    }

    [Fact]
    public void DefinitionList_NestedOrderedLazyParenDelimiter_SplitsOrderedListsLikeMarkdig() {
        const string markdown = """
Term
:   First
    1. item
1) sibling
""";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, CreateMarkdigDefinitionListReaderOptions());
        var definition = Assert.Single(Assert.Single(Assert.IsType<DefinitionListBlock>(Assert.Single(result.Document.Blocks)).Groups).Definitions);
        var firstOrdered = Assert.IsType<OrderedListBlock>(definition.Blocks[1]);
        var secondOrdered = Assert.IsType<OrderedListBlock>(definition.Blocks[2]);
        var definitionValue = result.SyntaxTree.Children[0].Children[0].Children.Single(child => child.Kind == MarkdownSyntaxKind.DefinitionValue);
        var written = NormalizeMarkdown(result.Document.ToMarkdown());
        var reparsedOffice = MarkdownReader.Parse(written, CreateMarkdigDefinitionListReaderOptions()).ToHtmlFragment(CreateMarkdigDefinitionListHtmlOptions());
        var office = result.Document.ToHtmlFragment(CreateMarkdigDefinitionListHtmlOptions());
        var markdig = MarkdigMarkdown.ToHtml(markdown, CreateMarkdigDefinitionListPipeline());

        Assert.Equal("item", Assert.Single(firstOrdered.Items).Content.RenderMarkdown());
        Assert.Equal("sibling", Assert.Single(secondOrdered.Items).Content.RenderMarkdown());
        Assert.Equal(
            new[] {
                MarkdownSyntaxKind.Paragraph,
                MarkdownSyntaxKind.OrderedList,
                MarkdownSyntaxKind.OrderedList
            },
            definitionValue.Children.Select(child => child.Kind).ToArray());
        Assert.Equal("Term\n:   First\n    1. item\n\n    1. sibling", written);
        Assert.Equal(NormalizeHtml(markdig), NormalizeHtml(office));
        Assert.Equal(NormalizeHtml(markdig), NormalizeHtml(reparsedOffice));
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);
    }

    [Fact]
    public void DefinitionList_NestedListEscapedPipeLazyText_RendersEscapedPipeLikeMarkdig() {
        const string markdown = """
Term
:   First
    - item
A \| B | C
---|---
D | E
""";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, CreateMarkdigDefinitionListReaderOptions());
        var definition = Assert.Single(Assert.Single(Assert.IsType<DefinitionListBlock>(Assert.Single(result.Document.Blocks)).Groups).Definitions);
        var unordered = Assert.IsType<UnorderedListBlock>(definition.Blocks[1]);
        var written = NormalizeMarkdown(result.Document.ToMarkdown());
        var reparsedOffice = MarkdownReader.Parse(written, CreateMarkdigDefinitionListReaderOptions()).ToHtmlFragment(CreateMarkdigDefinitionListHtmlOptions());
        var office = result.Document.ToHtmlFragment(CreateMarkdigDefinitionListHtmlOptions());
        var markdig = MarkdigMarkdown.ToHtml(markdown, CreateMarkdigDefinitionListPipeline());

        Assert.Equal("item\nA \\| B \\| C\n---\\|---\nD \\| E", Assert.Single(unordered.Items).Content.RenderMarkdown().Replace("\r\n", "\n"));
        Assert.Contains(@"A \| B \| C", written, StringComparison.Ordinal);
        Assert.Equal(NormalizeHtml(markdig), NormalizeHtml(office));
        Assert.Equal(NormalizeHtml(markdig), NormalizeHtml(reparsedOffice));
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);
    }

    [Fact]
    public void DefinitionList_TableDelimiterWithFewerCells_PadsLikeMarkdig_WhenTablesAreOn() {
        const string markdown = """
Term
:   First
| A | B |
|---|
| C |
""";

        var readerOptions = CreateMarkdigDefinitionListReaderOptions();
        readerOptions.Tables = true;
        var result = MarkdownReader.ParseWithSyntaxTree(markdown, readerOptions);
        var definition = Assert.Single(Assert.Single(Assert.IsType<DefinitionListBlock>(Assert.Single(result.Document.Blocks)).Groups).Definitions);
        var table = Assert.IsType<TableBlock>(definition.Blocks[1]);
        var written = NormalizeMarkdown(result.Document.ToMarkdown());
        var reparsedOffice = MarkdownReader.Parse(written, readerOptions).ToHtmlFragment(CreateMarkdigDefinitionListHtmlOptions());
        var office = result.Document.ToHtmlFragment(CreateMarkdigDefinitionListHtmlOptions());
        var markdig = MarkdigMarkdown.ToHtml(markdown, CreateMarkdigDefinitionListAndPipeTablesPipeline());

        Assert.Equal(new[] { "A", "B" }, table.Headers);
        Assert.Equal(new[] { "C" }, Assert.Single(table.Rows));
        Assert.Equal("Term\n:   First\n    | A | B |\n    | --- | --- |\n    | C |  |", written);
        Assert.Equal(NormalizeHtml(markdig), NormalizeHtml(office));
        Assert.Equal(NormalizeHtml(markdig), NormalizeHtml(reparsedOffice));
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);
    }

    private static Markdig.MarkdownPipeline CreateMarkdigDefinitionListPipeline() {
        var builder = new Markdig.MarkdownPipelineBuilder();
        Markdig.MarkdownExtensions.UseDefinitionLists(builder);
        return builder.Build();
    }

    private static Markdig.MarkdownPipeline CreateMarkdigDefinitionListAndPipeTablesPipeline() {
        var builder = new Markdig.MarkdownPipelineBuilder();
        Markdig.MarkdownExtensions.UseDefinitionLists(builder);
        Markdig.MarkdownExtensions.UsePipeTables(builder);
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

    private static string NormalizeMarkdown(string markdown) =>
        markdown
            .Replace("\r\n", "\n")
            .Replace('\r', '\n')
            .TrimEnd('\n');
}
