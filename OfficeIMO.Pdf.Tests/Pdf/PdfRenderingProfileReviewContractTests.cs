using OfficeIMO.Drawing;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Pdf.Tests;

public sealed class PdfRenderingProfileReviewContractTests {
    [Fact]
    public void FallbackPlannerAdvancesPastCompleteUncoveredWhitespaceGrapheme() {
        var combiningMarkRange = new OfficeFontUnicodeRangeSet(new[] {
            new OfficeFontUnicodeRange(0x0301, 0x0301)
        });
        var fallbackSet = new PdfEmbeddedFontFallbackSet(new[] {
            new PdfEmbeddedFontFallbackCandidate(
                "Combining mark",
                ManagedTextShapingTestAssets.CreateFont(0x0301),
                combiningMarkRange)
        });

        PdfTextFallbackPlan plan = fallbackSet.PlanText(" \u0301");

        Assert.Empty(plan.Segments);
        Assert.Empty(plan.Diagnostics);
    }
}
