using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfPageImageRendererTests {
    [Theory]
    [InlineData("", "W")]
    [InlineData("20 110 m", "W*")]
    [InlineData("20 110 0 30 re", "W")]
    [InlineData("20 110 80 0 re", "W*")]
    [InlineData("100 110 -80 0 re", "W")]
    [InlineData("20 110 m 100 110 l h", "W*")]
    [InlineData("20 110 m 20 140 l h", "W")]
    [InlineData("20 110 0 30 re W n 0 0 240 200 re", "W*")]
    [InlineData("0 0 5 5 re W n 20 110 80 30 re", "W")]
    public void EmptyGraphicsClipKeepsSearchableTextWithoutPaintingIt(string path, string clipOperator) {
        byte[] pdf = BuildSingleStreamPdf(
            "q\n" + path + " " + clipOperator + " n\n" +
            "1 0 0 rg 20 110 80 30 re f\n" +
            "BT /F1 20 Tf 20 130 Td (HIDDEN) Tj ET\nQ\n" +
            "0 0 1 rg BT /F1 20 Tf 20 70 Td (VISIBLE) Tj ET",
            EmptyClipFontResources, EmptyClipFontObject);

        PdfReadPage page = PdfReadDocument.Open(pdf).Pages[0];
        Assert.Equal(new[] { "HIDDEN", "VISIBLE" }, page.GetTextSpans().Select(span => span.Text));
        Assert.All(page.GetTextSpans(), span => Assert.True(span.IsVisible));
        Assert.Contains("HIDDEN", page.ExtractText());
        AssertEmptyClipPaint(page.ToDrawing());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void EmptyGraphicsClipIsInheritedByFormTextAndRestoredAfterTheForm(bool insideForm) {
        const string clip = "20 110 0 30 re W n\n";
        const string hidden = "1 0 0 rg BT /F1 20 Tf 20 130 Td (HIDDEN) Tj ET\n";
        string form = BuildStreamObject(6,
            "<< /Type /XObject /Subtype /Form /BBox [0 0 240 200] /Resources " + EmptyClipFontResources,
            (insideForm ? clip : "") + hidden);
        byte[] pdf = BuildSingleStreamPdf(
            "q\n" + (insideForm ? "" : clip) + "/Fm1 Do\nQ\n" +
            "0 0 1 rg BT /F1 20 Tf 20 70 Td (VISIBLE) Tj ET",
            "<< /Font << /F1 5 0 R >> /XObject << /Fm1 6 0 R >> >>", EmptyClipFontObject, form);

        PdfReadPage page = PdfReadDocument.Open(pdf).Pages[0];
        Assert.Equal(new[] { "HIDDEN", "VISIBLE" }, page.GetTextSpans().Select(span => span.Text).OrderBy(text => text, System.StringComparer.Ordinal));
        AssertEmptyClipPaint(page.ToDrawing());
    }

    [Theory]
    [InlineData("100 110 -80 40 re", "W")]
    [InlineData("20 150 80 -40 re", "W*")]
    [InlineData("100 150 -80 -40 re", "W")]
    public void ReversedRectangleClipStillPaintsItsNonemptyInterior(string path, string clipOperator) {
        byte[] pdf = BuildSingleStreamPdf(
            "q " + path + " " + clipOperator + " n\n" +
            "1 0 0 rg BT /F1 20 Tf 20 130 Td (INSIDE) Tj ET\nQ",
            EmptyClipFontResources, EmptyClipFontObject);
        PdfReadPage page = PdfReadDocument.Open(pdf).Pages[0];

        Assert.Equal("INSIDE", Assert.Single(page.GetTextSpans()).Text);
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(page.ToDrawing(), background: OfficeColor.White);
        Assert.True(CountEmptyClipColor(raster, red: true) > 0, "Negative rectangle dimensions reverse the path without emptying its interior.");
    }

    private const string EmptyClipFontResources = "<< /Font << /F1 5 0 R >> >>";
    private const string EmptyClipFontObject = "5 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>\nendobj";

    private static void AssertEmptyClipPaint(OfficeDrawing drawing) {
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(drawing, background: OfficeColor.White);
        Assert.Equal(0, CountEmptyClipColor(raster, red: true));
        Assert.True(CountEmptyClipColor(raster, red: false) > 0, "Q must restore visible blue text after the empty clipping scope.");
    }

    private static int CountEmptyClipColor(OfficeRasterImage raster, bool red) {
        int count = 0;
        for (int y = 0; y < raster.Height; y++) {
            for (int x = 0; x < raster.Width; x++) {
                OfficeColor pixel = raster.GetPixel(x, y);
                if (red ? pixel.R > pixel.G + 20 && pixel.R > pixel.B + 20
                        : pixel.B > pixel.R + 20 && pixel.B > pixel.G + 20) count++;
            }
        }
        return count;
    }
}
