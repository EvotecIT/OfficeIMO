using System.IO;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfDocumentRasterVisualBaselineTests {
    [Theory]
    [InlineData("markdown-theme-gallery-report")]
    [InlineData("markdown-theme-gallery-github-like")]
    [InlineData("markdown-theme-gallery-plain")]
    [InlineData("markdown-technical-document")]
    public void RasterComparisonRejectsRecordedMissingText(string scenario) {
        string root = GetPdfTestsProjectRoot();
        byte[] expected = File.ReadAllBytes(Path.Combine(root, "Pdf", "VisualBaselines", "officeimo-pdf-" + scenario + ".page1.poppler.png"));
        byte[] actual = File.ReadAllBytes(Path.Combine(root, "Pdf", "Fixtures", "RasterRegressions", scenario + "-missing-text.png"));

        VisualRasterComparison comparison = CompareRasterImages(expected, actual);

        // These real renderer failures fit every existing whole-page error budget.
        Assert.True(comparison.DifferentPixels <= comparison.AllowedDifferentPixels);
        Assert.True(comparison.MeanAbsoluteError <= comparison.MaximumMeanAbsoluteError);
        Assert.True(comparison.RootMeanSquareError <= comparison.MaximumRootMeanSquareError);
        Assert.True(comparison.MeanLuminanceError <= comparison.MaximumMeanLuminanceError);
        Assert.False(comparison.Passed);
        Assert.True(comparison.DarkPixelRetentionRatio < comparison.MinimumDarkPixelRetentionRatio);
    }

    [Theory]
    [InlineData(0, 255, false, false)]
    [InlineData(50, 255, false, true)]
    [InlineData(100, 0, false, false)]
    [InlineData(0, 255, true, false)]
    public void RasterComparisonChecksVisibleInkWithinPerceptualBounds(int retainedPixels, int alpha, bool roundPageSize, bool shouldPass) {
        OfficeRasterImage expected = new OfficeRasterImage(115, 100, OfficeColor.White);
        OfficeRasterImage actual = new OfficeRasterImage(roundPageSize ? 114 : 115, 100, OfficeColor.White);
        for (int x = 0; x < 100; x++) expected.SetPixel(x, 0, OfficeColor.Black);
        for (int x = 0; x < retainedPixels; x++) actual.SetPixel(x, 0, OfficeColor.FromRgba(0, 0, 0, (byte)alpha));

        VisualRasterComparison comparison = CompareRasterImages(OfficePngWriter.Encode(expected), OfficePngWriter.Encode(actual));

        Assert.True(comparison.DifferentPixels <= comparison.AllowedDifferentPixels);
        Assert.True(comparison.MeanAbsoluteError <= comparison.MaximumMeanAbsoluteError);
        Assert.True(comparison.RootMeanSquareError <= comparison.MaximumRootMeanSquareError);
        Assert.True(comparison.MeanLuminanceError <= comparison.MaximumMeanLuminanceError);
        Assert.Equal(shouldPass, comparison.Passed);
    }
}
