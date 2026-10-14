using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void TextDecorationSubtreeScanEnforcesThePublicWorkBudget() {
        string html = DecorationScanInput(64, 65536);
        var options = new HtmlRenderOptions { MaxLayoutDepth = 160, MaxLayoutOperations = 100 };
        HtmlDomLimitException error = Assert.Throws<HtmlDomLimitException>(() => HtmlRenderTestDriver.Render(html, options));
        Assert.Equal(HtmlRenderDiagnosticCodes.LayoutOperationLimitExceeded, error.Code);
        Assert.Equal(nameof(HtmlRenderOptions.MaxLayoutOperations), error.LimitSource);
        Assert.Contains("text decoration", error.Message);
    }

    [Fact]
    public void TextDecorationDescendantResolutionsReuseTheBudgetedSubtreeSummary() {
        // One 64-level subtree scan, 64 character chunks and the ordinary inline
        // traversal fit. Repeated ancestor rescans would exceed this shared
        // public work budget despite emitting no text.
        var scene = HtmlRenderTestDriver.Render(DecorationScanInput(64, 16384),
            new HtmlRenderOptions { MaxLayoutDepth = 160, MaxLayoutOperations = 220 });
        Assert.Empty(scene.Pages.SelectMany(p => EnumerateCorpusVisuals(p.Scene)).OfType<HtmlRenderText>());
    }

    [Fact]
    public void TextDecorationScanPreservesSupplementaryBidiAcrossChunkBoundaries() {
        var scene = HtmlRenderTestDriver.Render("<span style='text-decoration:overline 3px'>" + new string('x', 255) + "\U0001E900</span>");
        Assert.Contains(scene.Diagnostics, d => d.Code == "HtmlRenderTextDecorationThicknessApproximated"
            && d.Detail?.Contains("bidi", StringComparison.Ordinal) == true);
        Assert.Empty(DecorationLines(scene));
    }

    private static string DecorationScanInput(int depth, int characters) =>
        "<div style='text-decoration:overline 3px;text-decoration-skip-ink:none'>"
        + string.Concat(Enumerable.Repeat("<span>", depth))
        + "<span style='display:none'>" + new string('x', characters) + "</span>"
        + string.Concat(Enumerable.Repeat("</span>", depth)) + "</div>";
}
