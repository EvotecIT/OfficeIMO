using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class MarkdownBlankLinesTests {
    [Fact]
    public void CollapseDuplicateBlankLines_CollapsesRepeatedBlankLinesAndPreservesText() {
        const string markdown = """
            # Transcript


            Status: healthy



            Impact: none
            """;

        var normalized = MarkdownBlankLines.CollapseDuplicateBlankLines(markdown).Replace("\r\n", "\n");

        Assert.Contains("# Transcript\n\nStatus: healthy\n\nImpact: none", normalized, StringComparison.Ordinal);
    }

    [Fact]
    public void CollapseDuplicateBlankLines_PreservesOriginalCrlfLineEndings() {
        const string markdown = "# Transcript\r\n\r\n\r\nStatus: healthy\r\n\r\nImpact: none\r\n";

        var normalized = MarkdownBlankLines.CollapseDuplicateBlankLines(markdown);

        Assert.Contains("\r\n\r\nStatus: healthy\r\n\r\nImpact: none\r\n", normalized, StringComparison.Ordinal);
        Assert.DoesNotContain("\n\n\n", normalized, StringComparison.Ordinal);
    }
}
