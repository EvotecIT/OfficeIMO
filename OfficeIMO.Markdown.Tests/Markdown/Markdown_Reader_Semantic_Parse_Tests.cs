using System.Text.Json;
using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite;

public sealed class Markdown_Reader_Semantic_Parse_Tests {
    [Fact]
    public void Parse_CommonMark_NestedLists_Matches_SourceBacked_Contract() {
        const string markdown =
            "# Checklist\r\n" +
            "\r\n" +
            "- Check **status** and a [runbook](https://example.com/runbook).\r\n" +
            "  - Evidence `01-01` is present.\r\n" +
            "  - Nested note:\r\n" +
            "    - [ ] capture result\r\n" +
            "    - [x] preserve text\r\n" +
            "- Continue with the next check.\r\n";
        var options = MarkdownReaderOptions.CreateCommonMarkProfile();

        var semanticDocument = MarkdownReader.ParseSemantic(markdown, options);
        var sourceBacked = MarkdownReader.ParseWithSyntaxTree(markdown, options);

        Assert.Equal(sourceBacked.Document.ToHtmlFragment(), semanticDocument.ToHtmlFragment());
        Assert.Equal(sourceBacked.Document.ToMarkdown(), semanticDocument.ToMarkdown());
        Assert.Contains("[ ] capture result", semanticDocument.ToHtmlFragment(), StringComparison.Ordinal);
        Assert.Contains("[x] preserve text", semanticDocument.ToHtmlFragment(), StringComparison.Ordinal);
        var link = Assert.Single(semanticDocument.DescendantObjectsOfType<LinkInline>());
        Assert.NotNull(link.LabelInlines);
        Assert.Same(link, link.LabelInlines!.Parent);
        Assert.Same(semanticDocument, link.LabelInlines.Document);
        MarkdownInvariantAssert.SemanticTreeIsWellFormed(semanticDocument);
        MarkdownInvariantAssert.SemanticTreeIsWellFormed(sourceBacked.Document);
        MarkdownInvariantAssert.SyntaxTreeIsWellFormed(sourceBacked.FinalSyntaxTree);
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(sourceBacked);
    }

    [Theory]
    [InlineData("\n")]
    [InlineData("\r\n")]
    [InlineData("\r")]
    public void Parse_CommonMark_LineEndings_Render_Equivalently(string lineEnding) {
        string markdown = string.Join(lineEnding, new[] {
            "# Header",
            string.Empty,
            "- first **strong** item",
            "  - nested [link](https://example.com)",
            "- second `code` item",
            string.Empty
        });
        var options = MarkdownReaderOptions.CreateCommonMarkProfile();

        var semanticDocument = MarkdownReader.ParseSemantic(markdown, options);
        var sourceBacked = MarkdownReader.ParseWithSyntaxTree(markdown, options);

        Assert.Equal(sourceBacked.Document.ToHtmlFragment(), semanticDocument.ToHtmlFragment());
        Assert.Equal(sourceBacked.Document.ToMarkdown(), semanticDocument.ToMarkdown());
        MarkdownInvariantAssert.SemanticTreeIsWellFormed(semanticDocument);
    }

    [Fact]
    public void Parse_CommonMark_Official_Corpus_Matches_SourceBacked_Rendering() {
        string path = Path.Combine(
            AppContext.BaseDirectory,
            "..", "..", "..",
            "Markdown",
            "Fixtures",
            "CommonMark",
            "commonmark-0.31.2-spec.json");
        var examples = JsonSerializer.Deserialize<List<CommonMarkSpecExample>>(
            File.ReadAllText(path),
            new JsonSerializerOptions { PropertyNameCaseInsensitive = true });

        Assert.NotNull(examples);
        Assert.NotEmpty(examples!);
        foreach (var example in examples!) {
            var options = MarkdownReaderOptions.CreateCommonMarkProfile();
            var semanticDocument = MarkdownReader.ParseSemantic(example.Markdown, options);
            var sourceBackedDocument = MarkdownReader.ParseWithSyntaxTree(example.Markdown, options).Document;

            Assert.True(
                string.Equals(
                    CommonMarkHtmlComparison.Normalize(sourceBackedDocument.ToHtmlFragment()),
                    CommonMarkHtmlComparison.Normalize(semanticDocument.ToHtmlFragment()),
                    StringComparison.Ordinal),
                $"CommonMark example {example.Example} ({example.Section}) produced different HTML. " +
                $"Source-backed: {Escape(sourceBackedDocument.ToHtmlFragment())}; semantic: {Escape(semanticDocument.ToHtmlFragment())}.");
            MarkdownInvariantAssert.SemanticTreeIsWellFormed(semanticDocument);
        }
    }

    [Fact]
    public void ParseSemantic_Does_Not_Capture_Public_Source_Spans() {
        const string markdown =
            "# Heading #\n\n" +
            "> quoted\n\n" +
            "---\n\n" +
            "[^note]: footnote body\n";

        var semanticDocument = MarkdownReader.ParseSemantic(markdown);
        var sourceBackedDocument = MarkdownReader.ParseWithSyntaxTree(markdown).Document;

        var semanticHeading = Assert.Single(semanticDocument.DescendantObjectsOfType<HeadingBlock>());
        Assert.Null(semanticHeading.LevelSourceSpan);
        Assert.Null(semanticHeading.TextSourceSpan);
        Assert.Null(semanticHeading.OpeningMarkerSourceSpan);
        Assert.Null(semanticHeading.ClosingMarkerSourceSpan);

        Assert.Null(Assert.Single(semanticDocument.DescendantObjectsOfType<HorizontalRuleBlock>()).MarkerSourceSpan);
        Assert.Empty(Assert.Single(semanticDocument.DescendantObjectsOfType<QuoteBlock>()).MarkerSourceSpans);

        var semanticFootnote = Assert.Single(semanticDocument.DescendantObjectsOfType<FootnoteDefinitionBlock>());
        Assert.Null(semanticFootnote.OpeningMarkerSourceSpan);
        Assert.Null(semanticFootnote.LabelSourceSpan);
        Assert.Null(semanticFootnote.SeparatorMarkerSourceSpan);

        Assert.NotNull(Assert.Single(sourceBackedDocument.DescendantObjectsOfType<HeadingBlock>()).OpeningMarkerSourceSpan);
        Assert.NotNull(Assert.Single(sourceBackedDocument.DescendantObjectsOfType<HorizontalRuleBlock>()).MarkerSourceSpan);
        Assert.NotEmpty(Assert.Single(sourceBackedDocument.DescendantObjectsOfType<QuoteBlock>()).MarkerSourceSpans);
        Assert.NotNull(Assert.Single(sourceBackedDocument.DescendantObjectsOfType<FootnoteDefinitionBlock>()).OpeningMarkerSourceSpan);
    }

    [Fact]
    public void ObjectTreeBinder_Binds_All_Child_Contracts_On_Custom_Inline() {
        var nestedInlines = new InlineSequence().Text("inline child");
        var blockChild = new ParagraphBlock(new InlineSequence().Text("block child"));
        var custom = new CompositeContainerInline(nestedInlines, blockChild);
        var document = MarkdownDoc.Create()
            .Add(new ParagraphBlock(new InlineSequence().AddRaw(custom)));

        Assert.Same(custom, nestedInlines.Parent);
        Assert.Same(custom, blockChild.Parent);
        Assert.Same(document, nestedInlines.Document);
        Assert.Same(document, blockChild.Document);
        Assert.Equal(0, nestedInlines.IndexInParent);
        Assert.Equal(1, blockChild.IndexInParent);
        Assert.Same(blockChild, nestedInlines.NextSibling);
        Assert.Same(nestedInlines, blockChild.PreviousSibling);
    }

    private static string Escape(string value) => value.Replace("\r", "\\r").Replace("\n", "\\n");

    private sealed class CompositeContainerInline : MarkdownInline, IRenderableMarkdownInline, IPlainTextMarkdownInline,
        IInlineContainerMarkdownInline, IChildMarkdownBlockContainer {
        public CompositeContainerInline(InlineSequence nestedInlines, IMarkdownBlock blockChild) {
            NestedInlines = nestedInlines;
            ChildBlocks = new[] { blockChild };
        }

        public InlineSequence NestedInlines { get; }
        public IReadOnlyList<IMarkdownBlock> ChildBlocks { get; }

        InlineSequence? IInlineContainerMarkdownInline.NestedInlines => NestedInlines;
        string IRenderableMarkdownInline.RenderMarkdown() => NestedInlines.RenderMarkdown();
        string IRenderableMarkdownInline.RenderHtml() => NestedInlines.RenderHtml();
        void IPlainTextMarkdownInline.AppendPlainText(StringBuilder builder) =>
            InlinePlainText.AppendPlainText(builder, NestedInlines);
    }
}
