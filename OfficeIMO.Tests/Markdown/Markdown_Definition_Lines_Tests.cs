using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class MarkdownDefinitionLinesTests {
    [Fact]
    public void SeparateAdjacentDefinitionLikeLinesOutsideFencedCodeBlocks_InsertsBlankLineBetweenAdjacentDefinitionLikeLines() {
        const string markdown = """
            # Transcript

            Status: healthy
            Impact: none
            """;

        var normalized = MarkdownDefinitionLines.SeparateAdjacentDefinitionLikeLinesOutsideFencedCodeBlocks(markdown)
            .Replace("\r\n", "\n");

        Assert.Contains("Status: healthy\n\nImpact: none", normalized, StringComparison.Ordinal);
    }

    [Fact]
    public void SeparateAdjacentDefinitionLikeLinesOutsideFencedCodeBlocks_DoesNotSplitAdjacentDefinitionLikeLinesInsideFence() {
        const string markdown = """
            # Transcript

            ```text
            Status: healthy
            Impact: none
            ```
            """;

        var normalized = MarkdownDefinitionLines.SeparateAdjacentDefinitionLikeLinesOutsideFencedCodeBlocks(markdown)
            .Replace("\r\n", "\n");

        Assert.Contains("```text\nStatus: healthy\nImpact: none\n```", normalized, StringComparison.Ordinal);
    }

    [Fact]
    public void SeparateAdjacentDefinitionLikeLinesOutsideFencedCodeBlocks_DoesNotTreatInlineCodeSeparatorAsDefinitionBoundary() {
        const string markdown = """
            Use `key: value` syntax when defining pairs.
            Keep this line unchanged.
            """;

        var normalized = MarkdownDefinitionLines.SeparateAdjacentDefinitionLikeLinesOutsideFencedCodeBlocks(markdown);

        Assert.Equal(markdown, normalized);
    }
}
