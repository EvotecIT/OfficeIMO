using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class MarkdownTranscriptTransportMarkersTests {
    [Fact]
    public void StripIntelligenceXCachedEvidenceTransportMarkers_RemovesMarkerLinesCaseInsensitively() {
        const string markdown = """
            # Transcript

            [Cached evidence fallback]
            IX:CACHED-TOOL-EVIDENCE:V1

            ### Result
            """;

        var normalized = MarkdownTranscriptTransportMarkers.StripIntelligenceXCachedEvidenceTransportMarkers(markdown)
            .Replace("\r\n", "\n");

        Assert.DoesNotContain("cached-tool-evidence", normalized, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("[Cached evidence fallback]", normalized, StringComparison.Ordinal);
        Assert.Contains("### Result", normalized, StringComparison.Ordinal);
    }

    [Fact]
    public void StripIntelligenceXCachedEvidenceTransportMarkers_PreservesWindowsLineEndings() {
        const string markdown = "# Transcript\r\n\r\nix:cached-tool-evidence:v1\r\n\r\n### Result\r\n";

        var normalized = MarkdownTranscriptTransportMarkers.StripIntelligenceXCachedEvidenceTransportMarkers(markdown);

        Assert.Contains("\r\n", normalized, StringComparison.Ordinal);
        Assert.DoesNotContain("ix:cached-tool-evidence:v1", normalized, StringComparison.OrdinalIgnoreCase);
    }
}
