using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfTextDirectionAnalysisTests {
    [Fact]
    public void RestoresPureRightToLeftGlyphPaintSequenceToLogicalOrder() {
        string visualGlyphOrder = "ةيبر";

        string logical = PdfTextDirectionAnalysis.RestoreLogicalOrderFromGlyphPaintSequence(
            visualGlyphOrder,
            glyphSequenceProgressesLeftToRight: true);

        Assert.Equal("ربية", logical);
    }

    [Fact]
    public void PreservesRightToLeftTextWithoutGlyphOrderProvenance() {
        const string logicalText = "العربية";

        string result = PdfTextDirectionAnalysis.RestoreLogicalOrderFromGlyphPaintSequence(
            logicalText,
            glyphSequenceProgressesLeftToRight: false);

        Assert.Equal(logicalText, result);
    }

    [Theory]
    [InlineData("جاهز:")]
    [InlineData("جاهز123")]
    [InlineData("جاهزReady")]
    public void PreservesMixedDirectionalGlyphRunsWhenLogicalOrderIsAmbiguous(string text) {
        string result = PdfTextDirectionAnalysis.RestoreLogicalOrderFromGlyphPaintSequence(
            text,
            glyphSequenceProgressesLeftToRight: true);

        Assert.Equal(text, result);
    }
}
