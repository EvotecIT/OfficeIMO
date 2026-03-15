using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class MarkdownTranscriptPreparationTests {
    [Fact]
    public void PrepareIntelligenceXTranscriptBody_ComposesNormalizationAndOrderedListRepair() {
        const string markdown = """
            1) First check
            2) Second check
            """;

        var prepared = MarkdownTranscriptPreparation.PrepareIntelligenceXTranscriptBody(markdown)
            .Replace("\r\n", "\n");

        Assert.Contains("1. First check\n\n2. Second check", prepared, StringComparison.Ordinal);
    }

    [Fact]
    public void PrepareIntelligenceXTranscriptForExport_CollapsesDuplicateBlankLines() {
        const string markdown = """
            # Transcript


            Status: healthy



            ### Result
            """;

        var prepared = MarkdownTranscriptPreparation.PrepareIntelligenceXTranscriptForExport(markdown)
            .Replace("\r\n", "\n");

        Assert.DoesNotContain("\n\n\n", prepared, StringComparison.Ordinal);
        Assert.Contains("# Transcript\n\nStatus: healthy\n\n### Result", prepared, StringComparison.Ordinal);
    }

    [Fact]
    public void PrepareIntelligenceXTranscriptForDocx_OptionallySeparatesGroupedDefinitionLikeParagraphs() {
        const string markdown = """
            Status: healthy
            Impact: none
            """;

        var preserved = MarkdownTranscriptPreparation.PrepareIntelligenceXTranscriptForDocx(markdown, preservesGroupedDefinitionLikeParagraphs: true)
            .Replace("\r\n", "\n");
        var repaired = MarkdownTranscriptPreparation.PrepareIntelligenceXTranscriptForDocx(markdown, preservesGroupedDefinitionLikeParagraphs: false)
            .Replace("\r\n", "\n");

        Assert.Contains("Status: healthy\nImpact: none", preserved, StringComparison.Ordinal);
        Assert.Contains("Status: healthy\n\nImpact: none", repaired, StringComparison.Ordinal);
    }
}
