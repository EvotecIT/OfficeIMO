using System.Text;
using OfficeIMO.Markdown;
using OfficeIMO.Word.Markdown;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class MarkdownAllSeverityBatch12SecurityTests {
    [Fact]
    public void MarkdownReader_MixedEmphasisMarkersUseIndexedClosingRuns() {
        string markdown = "*outer "
            + string.Join(" ", Enumerable.Repeat("_inner", 4_000))
            + "*";

        MarkdownDoc document = MarkdownReader.Parse(markdown);

        Assert.Single(document.Blocks);
        Assert.Contains("outer", document.ToMarkdown(), StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownReader_RejectsDetailsBeyondConfiguredNestingDepth() {
        string markdown = BuildNestedDetails(8);
        var options = new MarkdownReaderOptions { MaxNestingDepth = 4 };

        InvalidDataException exception = Assert.Throws<InvalidDataException>(
            () => MarkdownReader.Parse(markdown, options));

        Assert.Contains("MaxNestingDepth (4)", exception.Message,
            StringComparison.Ordinal);
        Assert.Single(MarkdownReader.Parse(
            BuildNestedDetails(2), options).Blocks);
    }

    [Fact]
    public void MarkdownReader_HtmlBlockWithManyTagStartsAvoidsSuffixCopies() {
        string markdown = "<div>" + new string('<', 50_000) + "</div>";

        MarkdownDoc document = MarkdownReader.Parse(markdown);

        Assert.Single(document.Blocks);
    }

    [Fact]
    public void MarkdownToWord_RejectsDeepBlockAndInlineTrees() {
        var blockDocument = MarkdownDoc.Create();
        blockDocument.Add(CreateNestedDetailsBlock(8));
        var options = new MarkdownToWordOptions { MaxNestingDepth = 4 };

        InvalidDataException blockException = Assert.Throws<InvalidDataException>(
            () => blockDocument.ToWordDocument(options));
        Assert.Contains("block nesting", blockException.Message,
            StringComparison.Ordinal);

        var inlineDocument = MarkdownDoc.Create();
        inlineDocument.Add(new ParagraphBlock(CreateNestedInlineSequence(8)));
        InvalidDataException inlineException = Assert.Throws<InvalidDataException>(
            () => inlineDocument.ToWordDocument(options));
        Assert.Contains("inline nesting", inlineException.Message,
            StringComparison.Ordinal);

        var safeDocument = MarkdownDoc.Create();
        safeDocument.Add(new ParagraphBlock(CreateNestedInlineSequence(2)));
        using var converted = safeDocument.ToWordDocument(options);
        Assert.NotNull(converted);
    }

    [Fact]
    public void MarkdownRenderer_TableClipboardExportsNeutralizeSpreadsheetFormulas() {
        var options = new MarkdownRenderer.MarkdownRendererOptions {
            EnableTableCopyButtons = true
        };
        string shell = MarkdownRenderer.MarkdownRenderer.BuildShellHtml(
            "Security", options);

        Assert.Contains("function omdSpreadsheetSafeCell(value)", shell,
            StringComparison.Ordinal);
        Assert.Contains("replace(/[\\t\\r\\n]+/g, ' ')", shell,
            StringComparison.Ordinal);
        Assert.Contains("/^[=+\\-@]/.test(s)", shell,
            StringComparison.Ordinal);
        Assert.Equal(2, CountOccurrences(shell,
            "omdSpreadsheetSafeCell(omdCellText(c))"));
    }

    private static string BuildNestedDetails(int depth) {
        var markdown = new StringBuilder();
        for (int index = 0; index < depth; index++) {
            markdown.AppendLine("<details>");
            markdown.AppendLine("<summary>Level</summary>");
        }
        markdown.AppendLine("safe");
        for (int index = 0; index < depth; index++) {
            markdown.AppendLine("</details>");
        }
        return markdown.ToString();
    }

    private static DetailsBlock CreateNestedDetailsBlock(int depth) {
        IMarkdownBlock child = new ParagraphBlock(
            new InlineSequence().Text("safe"));
        for (int index = 0; index < depth; index++) {
            child = new DetailsBlock(new SummaryBlock("Level"),
                new[] { child });
        }
        return (DetailsBlock)child;
    }

    private static InlineSequence CreateNestedInlineSequence(int depth) {
        InlineSequence current = new InlineSequence().Text("safe");
        for (int index = 0; index < depth; index++) {
            var parent = new InlineSequence();
            parent.ReplaceItems(new IMarkdownInline[] {
                new BoldSequenceInline(current)
            });
            current = parent;
        }
        return current;
    }

    private static int CountOccurrences(string source, string value) {
        int count = 0;
        int offset = 0;
        while ((offset = source.IndexOf(value, offset,
                   StringComparison.Ordinal)) >= 0) {
            count++;
            offset += value.Length;
        }
        return count;
    }
}
