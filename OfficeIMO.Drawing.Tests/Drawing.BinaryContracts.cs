using OfficeIMO.Drawing;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingBinaryContractTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PublishedDescriptorLoaderPreservesDecodedBudgetAndFaceDescriptor(bool fitsBudget) {
        byte[] font = ManagedTextShapingTestAssets.CreateFont('A');
        var fonts = new OfficeFontFaceCollection();
        var descriptor = new OfficeFontFaceDescriptor(600, 100, OfficeFontSlant.Normal);
        int retainedFontBytes = font.Length * 2;

        bool added = fonts.TryAddBounded("Ink", font, descriptor, OfficeFontUnicodeRangeSet.All,
            fitsBudget ? retainedFontBytes : retainedFontBytes - 1, out int decodedBytes, out string? error);

        Assert.Equal(fitsBudget, added);
        if (fitsBudget) {
            Assert.Null(error);
            Assert.Equal(retainedFontBytes, decodedBytes);
            Assert.Equal(descriptor, Assert.Single(fonts.Faces).Descriptor);
        } else {
            Assert.Equal(0, decodedBytes);
            Assert.Empty(fonts.Faces);
            Assert.Contains("limit", error, StringComparison.OrdinalIgnoreCase);
        }
    }

    [Fact]
    public void PublishedTextMeasurementSignatureRetainsLayoutFrameAndInk() {
        var fonts = new OfficeFontFaceCollection().Add("Ink", ManagedTextShapingTestAssets.CreateFont('A'));
        var canvas = new OfficeRasterCanvas(new OfficeRasterImage(80, 80), fonts: fonts);

        // The explicit five-value type protects the signature embedded in released Pdf assemblies.
        (double Left, double Top, double Right, double Bottom, bool HasInk) bounds =
            canvas.MeasurePositionedTextBounds("A", 20, 20, 40, 40, 20, new OfficeFontInfo("Ink", 20), 10,
                OfficeTextAlignment.Left, OfficeTextFeatureSettings.Default, "normal", 20,
                OfficeTextDecorationStyle.None, OfficeTextDecorationStyle.None, OfficeTextDirection.Auto);

        Assert.True(bounds.HasInk);
        Assert.Equal(20D, bounds.Left);
        Assert.Equal(20D, bounds.Top);
        Assert.Equal(60D, bounds.Right);
        Assert.Equal(60D, bounds.Bottom);
    }
}
