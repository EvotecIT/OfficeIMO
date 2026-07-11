using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Drawing.Tests;

public class DrawingArabicTextShaperTests {
    [Fact]
    public void ArabicTextShaper_AppliesContextualFormsAndRoundTripsLogicalText() {
        const string logical = "سلام";

        string shaped = OfficeArabicTextShaper.Shape(logical);

        Assert.Equal("\uFEB3\uFEE0\uFE8E\uFEE1", shaped);
        Assert.Equal(logical, OfficeArabicTextShaper.ToLogicalText(shaped));
        Assert.Equal(logical.Length, shaped.Length);
        Assert.True(OfficeArabicTextShaper.CanShapeAllJoiningCharacters(logical));
    }

    [Fact]
    public void ArabicTextShaper_RespectsNonJoiningLettersMarksAndJoinControls() {
        Assert.Equal("\uFE8D\uFE8F", OfficeArabicTextShaper.Shape("اب"));
        Assert.Equal("\uFE91\u064E\uFE90", OfficeArabicTextShaper.Shape("بَب"));
        Assert.Equal("\uFE8F\u200C\uFE8F", OfficeArabicTextShaper.Shape("ب\u200Cب"));
        Assert.Equal("\uFE91\u200D\uFE90", OfficeArabicTextShaper.Shape("ب\u200Dب"));
    }

    [Fact]
    public void ArabicTextShaper_LeavesUnsupportedJoiningScriptsVisibleToDiagnostics() {
        Assert.False(OfficeArabicTextShaper.CanShapeAllJoiningCharacters("ܫܠܡ"));
        Assert.Equal("ܫܠܡ", OfficeArabicTextShaper.Shape("ܫܠܡ"));
        Assert.True(OfficeArabicTextShaper.CanShapeAllJoiningCharacters("123، سلام"));
    }
}
