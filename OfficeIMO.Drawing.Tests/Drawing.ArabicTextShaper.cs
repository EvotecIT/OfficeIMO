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
    public void ArabicTextShaper_ShapesExtendedPersianAndUrduLetters() {
        // peh + farsi yeh + keheh; heh goal + yeh barree, then tteh + alef (Presentation Forms-A).
        Assert.Equal("\uFB58\uFBFF\uFB8F", OfficeArabicTextShaper.Shape("\u067E\u06CC\u06A9"));
        Assert.Equal("\uFBA8\uFBAF\u0020\uFB68\uFE8E", OfficeArabicTextShaper.Shape("\u06C1\u06D2\u0020\u0679\u0627"));
        Assert.Equal("\uFB8A", OfficeArabicTextShaper.Shape("\u0698"));
        Assert.Equal("\u06C1\u06D2\u0020\u0679\u0627", OfficeArabicTextShaper.ToLogicalText("\uFBA8\uFBAF\u0020\uFB68\uFE8E"));
        Assert.True(OfficeArabicTextShaper.CanShapeAllJoiningCharacters("\u067E\u06CC\u06A9"));
        Assert.True(OfficeArabicTextShaper.CanShapeAllJoiningCharacters("\u06C1\u06D2\u0020\u0679\u0627"));
    }

    [Fact]
    public void ArabicTextShaper_LeavesUnsupportedJoiningScriptsVisibleToDiagnostics() {
        Assert.False(OfficeArabicTextShaper.CanShapeAllJoiningCharacters("ܫܠܡ"));
        Assert.Equal("ܫܠܡ", OfficeArabicTextShaper.Shape("ܫܠܡ"));
        Assert.True(OfficeArabicTextShaper.CanShapeAllJoiningCharacters("123، سلام"));
    }
}
