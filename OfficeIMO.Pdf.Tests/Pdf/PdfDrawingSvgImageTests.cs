using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfDrawingSvgImageTests {
    [Fact]
    public void DrawingSvgImageRetainsVectorPaintInPdf() {
        byte[] svg = Encoding.UTF8.GetBytes(
            "<svg xmlns='http://www.w3.org/2000/svg' width='20' height='10'><rect width='10' height='10' fill='red'/><rect x='10' width='10' height='10' fill='blue'/></svg>");
        var projection = new OfficeImageProjection(new OfficeImagePlacement(0D, 0D, 40D, 20D));
        OfficeDrawing drawing = new OfficeDrawing(40D, 20D)
            .AddImage(svg, "image/svg+xml", projection, "Red and blue status image");

        byte[] pdf = PdfDocument.Create(Options())
            .Drawing(drawing)
            .ToBytes();

        Assert.Equal("%PDF", Encoding.ASCII.GetString(pdf, 0, 4));
        Assert.Empty(PdfDocument.Load(pdf).Reader.Images());
        OfficeDrawing imported = PdfReadDocument.Open(pdf).Pages[0].ToDrawing();
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(imported, background: OfficeColor.White);
        OfficeColor left = raster.GetPixel(10, 10);
        OfficeColor right = raster.GetPixel(30, 10);
        Assert.True(left.R > left.B);
        Assert.True(right.B > right.R);
    }
    [Fact]
    public void DrawingSvgImageTextRemainsSearchableWithoutRasterization() {
        byte[] svg = Encoding.UTF8.GetBytes(
            "<svg xmlns='http://www.w3.org/2000/svg' width='40' height='20'><text x='2' y='14' font-family='Arial' font-size='8'>Status</text></svg>");
        var drawing = new OfficeDrawing(40D, 20D).AddImage(svg, "image/svg+xml",
            new OfficeImageProjection(new OfficeImagePlacement(0D, 0D, 40D, 20D)));

        byte[] pdf = PdfDocument.Create(Options()).Drawing(drawing).ToBytes();

        Assert.Empty(PdfDocument.Load(pdf).Reader.Images());
        Assert.Contains("Status", PdfReadDocument.Open(pdf).ExtractText(), System.StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(false, 0D)]
    [InlineData(true, 0D)]
    [InlineData(false, 180D)]
    public void SvgProjectionPreservesCropClipAndFrameTransform(bool flip, double rotation) {
        byte[] svg = Encoding.UTF8.GetBytes("<svg xmlns='http://www.w3.org/2000/svg' width='40' height='20'><rect width='40' height='20' fill='red'/><rect x='20' width='20' height='20' fill='blue'/></svg>");
        var projection = new OfficeImageProjection(new OfficeImagePlacement(8D, 8D, 24D, 12D),
            new OfficeImageSourceCrop(0.25D, 0D, 0.25D, 0D), rotationDegrees: rotation, flipHorizontal: flip);
        var drawing = new OfficeDrawing(40D, 30D).AddImage(svg, "image/svg+xml", projection);
        byte[] pdf = PdfDocument.Create(Options()).Drawing(drawing).ToBytes();
        Assert.Empty(PdfDocument.Load(pdf).Reader.Images());
        OfficeRasterImage actual = OfficeDrawingRasterRenderer.Render(PdfReadDocument.Open(pdf).Pages[0].ToDrawing(), background: OfficeColor.White);
        bool reversed = flip ^ (rotation != 0D);
        Assert.Equal(reversed ? OfficeColor.Blue : OfficeColor.Red, actual.GetPixel(12, 14));
        Assert.Equal(reversed ? OfficeColor.Red : OfficeColor.Blue, actual.GetPixel(28, 14));
        Assert.Equal(OfficeColor.White, actual.GetPixel(5, 14));
        Assert.Equal(OfficeColor.White, actual.GetPixel(35, 14));
    }

    [Fact]
    public void SvgLuminanceMaskUsesDiagnosedRasterFallback() {
        byte[] svg = Encoding.UTF8.GetBytes("<svg xmlns='http://www.w3.org/2000/svg' width='40' height='20'><defs><mask id='m' maskUnits='userSpaceOnUse' x='0' y='0' width='40' height='20'><rect width='20' height='20' fill='white'/></mask></defs><rect width='40' height='20' fill='red' mask='url(#m)'/></svg>");
        var report = new PdfConversionReport();
        var drawing = new OfficeDrawing(40D, 20D).AddImage(svg, "image/svg+xml",
            new OfficeImageProjection(new OfficeImagePlacement(0D, 0D, 40D, 20D)));
        byte[] pdf = PdfDocument.Create(Options().ReportDiagnosticsTo(report)).Drawing(drawing).ToBytes();
        Assert.Contains(report.Warnings, warning => warning.Code == "SvgImageRasterFallback");
        Assert.Single(PdfDocument.Load(pdf).Reader.Images());
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(PdfReadDocument.Open(pdf).Pages[0].ToDrawing(), background: OfficeColor.White);
        Assert.Equal(OfficeColor.Red, raster.GetPixel(10, 10));
        Assert.Equal(OfficeColor.White, raster.GetPixel(30, 10));
    }

    private static PdfOptions Options() => new PdfOptions { PageWidth = 40D, PageHeight = 30D,
        MarginLeft = 0D, MarginRight = 0D, MarginTop = 0D, MarginBottom = 0D };
}
