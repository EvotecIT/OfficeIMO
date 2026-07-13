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
