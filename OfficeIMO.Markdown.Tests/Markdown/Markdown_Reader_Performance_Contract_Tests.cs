using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite;

public sealed class Markdown_Reader_Performance_Contract_Tests {
#if NET8_0_OR_GREATER
    [Fact]
    public void TableHeavyParse_DoesNotRebindTheGrowingDocumentForEveryBlock() {
        string markdown = BuildTableHeavyMarkdown(sectionCount: 80);
        _ = OfficeIMO.Markdown.MarkdownReader.ParseWithSyntaxTree("# Warmup\n\n| A | B |\n| - | - |\n| 1 | 2 |");

        long allocatedBefore = GC.GetAllocatedBytesForCurrentThread();
        MarkdownParseResult result = OfficeIMO.Markdown.MarkdownReader.ParseWithSyntaxTree(markdown);
        long allocatedBytes = GC.GetAllocatedBytesForCurrentThread() - allocatedBefore;

        Assert.Equal(160, result.Document.Blocks.Count);
        MarkdownInvariantAssert.SemanticTreeIsWellFormed(result.Document);
        Assert.True(
            allocatedBytes < 128L * 1024 * 1024,
            $"Table-heavy parsing allocated {allocatedBytes / (1024d * 1024d):N1} MB; repeated whole-document binding has likely returned.");
    }

    [Fact]
    public void LargeTableCell_WithManyNonLinks_DoesNotCopyTheRemainingTextForEachAutolinkCandidate() {
        var markdown = new StringBuilder("| Data |\n| --- |\n| ");
        for (int i = 0; i < 500; i++) markdown.Append("h w ");
        markdown.Append(new string('a', 64 * 1024));
        markdown.Append(" http://example.com www.example.com |\n");
        _ = MarkdownReader.Parse("| Data |\n| --- |\n| https://example.com |\n");

        long allocatedBefore = GC.GetAllocatedBytesForCurrentThread();
        MarkdownDoc document = MarkdownReader.Parse(markdown.ToString());
        long allocatedBytes = GC.GetAllocatedBytesForCurrentThread() - allocatedBefore;
        string html = document.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.Contains("href=\"http://example.com\"", html, StringComparison.Ordinal);
        Assert.Contains("href=\"https://www.example.com\"", html, StringComparison.Ordinal);
        Assert.True(allocatedBytes < 30L * 1024 * 1024,
            $"Parsing a large table cell allocated {allocatedBytes / (1024d * 1024d):N1} MB; rejected autolink prefixes should not copy the remaining cell.");
    }
#endif

    private static string BuildTableHeavyMarkdown(int sectionCount) {
        var markdown = new StringBuilder();
        for (int section = 1; section <= sectionCount; section++) {
            markdown.Append("## Section ").Append(section).AppendLine();
            markdown.AppendLine();
            markdown.AppendLine("| Name | Value | Status |");
            markdown.AppendLine("| --- | ---: | --- |");
            for (int row = 1; row <= 5; row++) {
                markdown.Append("| Item ").Append(row).Append(" | ").Append(section * row).AppendLine(" | Ready |");
            }

            markdown.AppendLine();
        }

        return markdown.ToString();
    }
}
