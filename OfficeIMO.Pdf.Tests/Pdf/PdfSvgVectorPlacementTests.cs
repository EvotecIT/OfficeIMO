using System.Text;
using System.Text.RegularExpressions;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfSvgVectorPlacementTests {
    private const string Logo = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 40 20'>" +
        "<defs><linearGradient id='g'><stop offset='0' stop-color='#ff0000'/><stop offset='1' stop-color='#ff0000'/></linearGradient></defs>" +
        "<rect width='20' height='20' fill='url(#g)'/><rect x='20' width='20' height='20' fill='blue' fill-opacity='1'/></svg>";

    [Theory]
    [InlineData(0D, false, OfficeImageFit.Stretch)]
    [InlineData(90D, false, OfficeImageFit.Stretch)]
    [InlineData(0D, true, OfficeImageFit.Contain)]
    [InlineData(0D, false, OfficeImageFit.Cover)]
    public void FlowSvgImageUsesTheRasterImagePlacementContract(double rotation, bool flip, OfficeImageFit fit) {
        byte[] svg = Encoding.UTF8.GetBytes(Logo);
        var pixels = new OfficeRasterImage(40, 20, OfficeColor.Red);
        for (int y = 0; y < 20; y++)
            for (int x = 20; x < 40; x++) pixels.SetPixel(x, y, OfficeColor.Blue);
        byte[] raster = OfficePngWriter.Encode(pixels);
        byte[] Export(byte[] bytes) => PdfDocument.Create(PageOptions(120D, 120D)).Image(bytes, 60D, 30D, OfficeAlignment.Center,
            style: new PdfImageStyle { Fit = fit, RotationAngle = rotation, HorizontalFlip = flip,
                SourceCrop = new PdfImageSourceCrop(.25D, 0D, 0D, 0D) }).ToBytes();

        byte[] pdf = Export(svg);

        Assert.Empty(PdfDocument.Load(pdf).Reader.Images());
        OfficeRasterImage actual = OfficeDrawingRasterRenderer.Render(PdfReadDocument.Open(pdf).Pages[0].ToDrawing(), background: OfficeColor.White);
        OfficeRasterImage expected = OfficeDrawingRasterRenderer.Render(PdfReadDocument.Open(Export(raster)).Pages[0].ToDrawing(), background: OfficeColor.White);
        for (int y = 4; y < 116; y += 4) {
            for (int x = 4; x < 116; x += 4) {
                OfficeColor color = expected.GetPixel(x, y);
                // Compare interiors. Raster interpolation and vector edge coverage differ.
                if (color != expected.GetPixel(x - 2, y) || color != expected.GetPixel(x + 2, y) ||
                    color != expected.GetPixel(x, y - 2) || color != expected.GetPixel(x, y + 2)) continue;
                Assert.Equal(color, actual.GetPixel(x, y));
            }
        }
    }

    [Fact]
    public void RepeatedSvgImageIsWrittenOnceAsASharedForm() {
        const int pages = 4;
        // Separate arrays with equal SVG, as converters produce when they read the same picture again.
        PdfDocument document = PdfDocument.Create(new PdfCreateOptions { CompressContentStreams = false });
        document.Header(header => header.Image(Encoding.UTF8.GetBytes(Logo), OfficeLength.Points(40), OfficeLength.Points(20), "Header logo"));
        for (int page = 0; page < pages; page++) {
            if (page > 0) document.PageBreak();
            document.Paragraph(paragraph => paragraph.AddText("Page " + (page + 1)));
            document.Image(Encoding.UTF8.GetBytes(Logo), 80D, 40D, "Body logo");
        }

        byte[] pdf = document.ToBytes();
        string raw = Encoding.GetEncoding("ISO-8859-1").GetString(pdf);

        Assert.Equal(pages, PdfReadDocument.Open(pdf).Pages.Count);
        Assert.Empty(PdfDocument.Load(pdf).Reader.Images());
        Assert.Single(Regex.Matches(raw, @"/Subtype /Form /FormType 1 /BBox \[0 0 40 20\]"));
        Assert.Equal(2 * pages, Regex.Matches(raw, @"/Sv1 Do").Count);
        // Each placement keeps its own alternative text around the shared drawing.
        Assert.Equal(pages, Regex.Matches(raw, "/Figure << /Alt <" + Hex("Body logo") + ">").Count);
        Assert.Equal(pages, Regex.Matches(raw, "/Figure << /Alt <" + Hex("Header logo") + ">").Count);
    }

    [Fact]
    public void InlineSvgImageWithTextStaysVector() {
        byte[] badge = Encoding.UTF8.GetBytes("<svg xmlns='http://www.w3.org/2000/svg' width='60' height='20'>" +
            "<rect width='60' height='20' fill='#008000'/><text x='8' y='14' font-size='11' fill='#ffffff'>build</text></svg>");

        byte[] pdf = PdfDocument.Create(PageOptions(200D, 60D))
            .Paragraph(paragraph => {
                paragraph.AddText("Status ");
                paragraph.AddInlineImage(badge, OfficeLength.Points(60), OfficeLength.Points(20), "Build passing");
            })
            .ToBytes();

        Assert.Empty(PdfDocument.Load(pdf).Reader.Images());
        Assert.Contains("Status", PdfReadDocument.Open(pdf).ExtractText());
        OfficeRasterImage page = OfficeDrawingRasterRenderer.Render(PdfReadDocument.Open(pdf).Pages[0].ToDrawing(), background: OfficeColor.White);
        int green = 0;
        for (int y = 0; y < page.Height; y++)
            for (int x = 0; x < page.Width; x++)
                if (page.GetPixel(x, y) == OfficeColor.FromRgb(0, 128, 0)) green++;
        Assert.True(green > 400, "The badge background was not drawn: " + green);
    }

    private static string Hex(string text) {
        var hex = new StringBuilder();
        foreach (byte value in Encoding.ASCII.GetBytes(text)) hex.Append(value.ToString("X2", System.Globalization.CultureInfo.InvariantCulture));
        return hex.ToString();
    }

    private static PdfCreateOptions PageOptions(double width, double height) => new PdfCreateOptions { PageSetup = {
        Size = new PdfPageSize(OfficeLength.Points(width), OfficeLength.Points(height)),
        Margins = { Left = OfficeLength.Points(0D), Right = OfficeLength.Points(0D), Top = OfficeLength.Points(0D), Bottom = OfficeLength.Points(0D) }
    } };
}
