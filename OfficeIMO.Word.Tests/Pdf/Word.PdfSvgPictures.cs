using System.Text;
using DocumentFormat.OpenXml.Packaging;
using System.Text.RegularExpressions;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class WordPdfSvgPictureTests {
    private const string Logo = "<svg xmlns='http://www.w3.org/2000/svg' width='40' height='20'>" +
        "<rect width='20' height='20' fill='#ff0000'/><rect x='20' width='20' height='20' fill='#0000ff'/></svg>";

    [Fact]
    public void SvgPicturesInBodyAndHeaderAreWrittenOnceAsVectors() {
        using WordDocument document = WordDocument.Create();
        document.Sections[0].PageSetup.Header.AddParagraph()
            .AddImage(new MemoryStream(Encoding.UTF8.GetBytes(Logo)), "logo.svg", OfficeLength.Points(40), OfficeLength.Points(20));
        for (int index = 0; index < 3; index++) {
            if (index > 0) document.AddPageBreak();
            document.AddParagraph("Page " + (index + 1));
            document.AddParagraph().AddImage(new MemoryStream(Encoding.UTF8.GetBytes(Logo)), "logo.svg", OfficeLength.Points(160), OfficeLength.Points(80));
        }

        byte[] pdf = document.ToPdf().ToBytes();

        Assert.Empty(PdfDocument.Load(pdf).Reader.Images());
        Assert.Single(Regex.Matches(Encoding.GetEncoding("ISO-8859-1").GetString(pdf), @"/Subtype /Form /FormType 1 /BBox \[0 0 40 20\]"));
        PdfReadDocument read = PdfReadDocument.Open(pdf);
        Assert.Equal(3, read.Pages.Count);
        Assert.Contains("Page 3", read.ExtractText());
        OfficeRasterImage page = OfficeDrawingRasterRenderer.Render(read.Pages[0].ToDrawing(), background: OfficeColor.White);
        int red = 0, blue = 0;
        for (int y = 0; y < page.Height; y++) {
            for (int x = 0; x < page.Width; x++) {
                OfficeColor color = page.GetPixel(x, y);
                if (color == OfficeColor.FromRgb(255, 0, 0)) red++;
                else if (color == OfficeColor.FromRgb(0, 0, 255)) blue++;
            }
        }
        // The body picture is 80 by 80 points per half, the header 20 by 20.
        Assert.InRange(red, 6000, 7200);
        Assert.InRange(blue, 6000, 7200);
    }

    [Fact]
    public void LinkedCroppedSvgPictureStaysVectorInsideItsLinkArea() {
        using WordDocument document = WordDocument.Create();
        WordParagraph paragraph = document.AddParagraph();
        paragraph.AddImage(new MemoryStream(Encoding.UTF8.GetBytes(Logo)), "logo.svg", OfficeLength.Points(80), OfficeLength.Points(40));
        // The crop keeps only the blue half of the logo.
        paragraph.Image!.CropLeftRatio = .5D;
        paragraph.Image.Hyperlink = new Uri("https://example.com/logo");

        byte[] pdf = document.ToPdf().ToBytes();

        Assert.Empty(PdfDocument.Load(pdf).Reader.Images());
        PdfLinkAnnotation link = Assert.Single(PdfDocumentReadResult.Load(pdf).GetLinksByUri("https://example.com/logo")).SourceLink;
        OfficeRasterImage page = OfficeDrawingRasterRenderer.Render(PdfReadDocument.Open(pdf).Pages[0].ToDrawing(), background: OfficeColor.White);
        int red = 0, left = int.MaxValue, top = int.MaxValue, right = -1, bottom = -1;
        for (int y = 0; y < page.Height; y++) {
            for (int x = 0; x < page.Width; x++) {
                OfficeColor color = page.GetPixel(x, y);
                if (color == OfficeColor.FromRgb(255, 0, 0)) red++;
                if (color != OfficeColor.FromRgb(0, 0, 255)) continue;
                left = Math.Min(left, x);
                top = Math.Min(top, y);
                right = Math.Max(right, x + 1);
                bottom = Math.Max(bottom, y + 1);
            }
        }

        Assert.Equal(0, red);
        Assert.InRange(Math.Abs(left - link.X1), 0D, 1.5D);
        Assert.InRange(Math.Abs(right - link.X2), 0D, 1.5D);
        Assert.InRange(Math.Abs(top - (page.Height - link.Y2)), 0D, 1.5D);
        Assert.InRange(Math.Abs(bottom - (page.Height - link.Y1)), 0D, 1.5D);
        Assert.InRange(link.Width, 79D, 81D);
        Assert.InRange(link.Height, 39D, 41D);
    }

    [Fact]
    public void LinkedSvgPictureFlowedFromAnUnsupportedWrapKeepsItsLink() {
        using WordDocument document = WordDocument.Create();
        WordParagraph paragraph = document.AddParagraph();
        paragraph.AddImage(new MemoryStream(Encoding.UTF8.GetBytes(Logo)), "logo.svg", OfficeLength.Points(80), OfficeLength.Points(40));
        // Tight wrapping is outside the fixed-placement profile, so the picture is placed in document flow.
        paragraph.Image!.WrapText = WordImageTextWrapping.Tight;
        paragraph.Image.Hyperlink = new Uri("https://example.com/flowed");

        byte[] pdf = document.ToPdf(null, out OfficeConversionReport report).ToBytes();

        Assert.Contains(report.Diagnostics, diagnostic => diagnostic.Code == "native-anchored-image-wrap-approximated");
        Assert.Empty(PdfDocument.Load(pdf).Reader.Images());
        PdfLinkAnnotation link = Assert.Single(PdfDocumentReadResult.Load(pdf).GetLinksByUri("https://example.com/flowed")).SourceLink;
        Assert.InRange(link.Width, 79D, 81D);
        Assert.InRange(link.Height, 39D, 41D);
    }

    [Theory]
    [InlineData(WordImageTextWrapping.InFrontOfText)]
    [InlineData(WordImageTextWrapping.BehindText)]
    [InlineData(WordImageTextWrapping.Square)]
    public void FloatingSvgPictureKeepsItsLinkOverItsPlacedFrame(WordImageTextWrapping wrapping) {
        using WordDocument document = WordDocument.Create();
        WordParagraph paragraph = document.AddParagraph("Floating picture");
        paragraph.AddImage(new MemoryStream(Encoding.UTF8.GetBytes(Logo)), "logo.svg", OfficeLength.Points(80), OfficeLength.Points(40));
        WordImage image = paragraph.Image!;
        image.WrapText = wrapping;
        image.HorizontalPositionRelativeFrom = WordHorizontalRelativePosition.Page;
        image.HorizontalPositionOffset = OfficeLength.Points(100);
        image.VerticalPositionRelativeFrom = WordVerticalRelativePosition.Page;
        image.VerticalPositionOffset = OfficeLength.Points(200);
        image.Hyperlink = new Uri("https://example.com/floating");

        byte[] pdf = document.ToPdf(null, out OfficeConversionReport report).ToBytes();

        Assert.DoesNotContain(report.Diagnostics, diagnostic => diagnostic.Code == "native-anchored-image-flowed" ||
            diagnostic.Code == "native-anchored-image-wrap-approximated");
        Assert.Empty(PdfDocument.Load(pdf).Reader.Images());
        PdfLinkAnnotation link = Assert.Single(PdfDocumentReadResult.Load(pdf).GetLinksByUri("https://example.com/floating")).SourceLink;
        OfficeRasterImage page = OfficeDrawingRasterRenderer.Render(PdfReadDocument.Open(pdf).Pages[0].ToDrawing(), background: OfficeColor.White);
        Assert.InRange(link.X1, 99D, 101D);
        Assert.InRange(page.Height - link.Y2, 199D, 201D);
        Assert.InRange(link.Width, 79D, 81D);
        Assert.InRange(link.Height, 39D, 41D);
        // The picture paints inside the link area: red left half, blue right half.
        Assert.Equal(OfficeColor.FromRgb(255, 0, 0), page.GetPixel(120, 220));
        Assert.Equal(OfficeColor.FromRgb(0, 0, 255), page.GetPixel(160, 220));
    }

    [Fact]
    public void SvgPictureWhoseSecondSvgPartDrawsStaysVectorWithoutAFallbackWarning() {
        using var stream = new MemoryStream();
        using (WordDocument document = WordDocument.Create(stream)) {
            document.AddParagraph().AddImage(new MemoryStream(Encoding.UTF8.GetBytes(Logo)), "logo.svg", OfficeLength.Points(80), OfficeLength.Points(40));
            document.Save();
        }
        OfficeSvgPartDamage.Truncate(stream);
        AddSecondSvgPart(stream, Logo);

        using WordDocument reopened = WordDocument.Load(stream);
        byte[] pdf = reopened.ToPdf(null, out OfficeConversionReport report).ToBytes();

        Assert.Empty(PdfDocument.Load(pdf).Reader.Images());
        Assert.DoesNotContain(report.Diagnostics, diagnostic => diagnostic.Code == OfficeDiagnosticCodes.SvgImageRasterFallback);
    }

    // Appends a second asvg:svgBlip extension, pointing to a new SVG part, after the picture's damaged one.
    private static void AddSecondSvgPart(MemoryStream package, string svg) {
        using (var document = WordprocessingDocument.Open(package, true)) {
            var main = document.MainDocumentPart!;
            DocumentFormat.OpenXml.Drawing.BlipExtension extension = main.Document!.Descendants<DocumentFormat.OpenXml.Drawing.BlipExtension>()
                .Single(item => item.Uri?.Value == "{96DAC541-7B7A-43D3-8B79-37D633B846F1}");
            ImagePart part = main.AddImagePart("image/svg+xml");
            using (var content = new MemoryStream(Encoding.UTF8.GetBytes(svg))) part.FeedData(content);
            var second = (DocumentFormat.OpenXml.Drawing.BlipExtension)extension.CloneNode(true);
            second.Descendants<DocumentFormat.OpenXml.Office2019.Drawing.SVG.SVGBlip>().Single().Embed = main.GetIdOfPart(part);
            extension.InsertAfterSelf(second);
        }
        package.Position = 0;
    }

    [Fact]
    public void SvgPictureThatCannotBeDrawnExportsItsRasterFallbackAndReportsIt() {
        using var stream = new MemoryStream();
        using (WordDocument document = WordDocument.Create(stream)) {
            document.AddParagraph().AddImage(new MemoryStream(Encoding.UTF8.GetBytes(Logo)), "logo.svg", OfficeLength.Points(80), OfficeLength.Points(40));
            document.Save();
        }
        OfficeSvgPartDamage.Truncate(stream);

        using WordDocument reopened = WordDocument.Load(stream);
        byte[] pdf = reopened.ToPdf(null, out OfficeConversionReport report).ToBytes();

        Assert.Single(PdfDocument.Load(pdf).Reader.Images());
        Assert.Single(report.Diagnostics, diagnostic => diagnostic.Code == OfficeDiagnosticCodes.SvgImageRasterFallback);
    }
}
