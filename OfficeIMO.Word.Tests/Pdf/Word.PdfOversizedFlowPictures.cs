using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using UglyToad.PdfPig.Content;
using Xunit;
using PdfCore = OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData("picture-normal-explicit-defaults.docx", 1, 1, 1, 50D, 62D, 240D, 160D, 231.67D)]
    [InlineData("picture-wide-margin-explicit-defaults.docx", 1, 1, 1, 50D, 62D, 360D, 120D, 191.69D)]
    [InlineData("picture-wide-center-explicit-defaults.docx", 1, 1, 1, 50D, 62D, 360D, 120D, 191.69D)]
    [InlineData("picture-wide-right-explicit-defaults.docx", 1, 1, 1, 50D, 62D, 360D, 120D, 191.69D)]
    [InlineData("picture-wide-pageclip-explicit-defaults.docx", 1, 1, 1, 50D, 62D, 420D, 140D, 211.61D)]
    [InlineData("picture-wide-nextpage-explicit-defaults.docx", 2, 2, 2, 50D, 50D, 360D, 120D, 179.69D)]
    [InlineData("picture-tall-explicit-defaults.docx", 3, 2, 3, 50D, 50D, 160D, 440D, 59.66D)]
    [InlineData("picture-wide-mixed-explicit-defaults.docx", 1, 1, 1, 50D, 73.5D, 360D, 120D, 214.61D)]
    [InlineData("picture-tall-mixed-explicit-defaults.docx", 3, 2, 3, 128.91D, 50D, 160D, 440D, 71.06D)]
    public void SaveAsPdf_FlowPicturesRetainAuthoredSizeAndNativePagePlacement(
        string fixture, int pageCount, int picturePage, int followingPage,
        double left, double top, double width, double height, double followingBaseline) {
        // These valid synthetic controls were independently exported by native
        // Word. They distinguish overflow of the body frame from actual page
        // clipping, alignment, page admission and mixed text/image line flow.
        byte[] bytes = ConvertQualifiedFlowPicture(fixture);
        using var pdf = PdfPigDocument.Open(bytes);
        Page[] pages = pdf.GetPages().ToArray();
        Assert.Equal(pageCount, pages.Length);
        for (int page = 0; page < pages.Length; page++) {
            Assert.InRange(Math.Abs(pages[page].Width - 400D), 0D, .15D);
            Assert.InRange(Math.Abs(pages[page].Height - 500D), 0D, .15D);
            if (page + 1 != picturePage) Assert.Empty(pages[page].GetImages());
        }

        var picture = Assert.Single(pages[picturePage - 1].GetImages());
        Assert.InRange(Math.Abs(picture.BoundingBox.Left - left), 0D, .35D);
        Assert.InRange(Math.Abs(pages[picturePage - 1].Height - picture.BoundingBox.Top - top), 0D, .35D);
        Assert.Equal(width, picture.BoundingBox.Width, 3);
        Assert.Equal(height, picture.BoundingBox.Height, 3);
        Assert.Contains("BEFOREPICTURE", FlowPictureText(pages[0]));
        Assert.Contains("AFTERPICTURE", FlowPictureText(pages[followingPage - 1]));
        var following = pages[followingPage - 1].GetWords().Last(word => word.Text == "AFTER");
        Assert.InRange(Math.Abs(pages[followingPage - 1].Height -
            following.Letters[0].StartBaseLine.Y - followingBaseline), 0D, .5D);
        if (fixture.Contains("mixed")) {
            Assert.Contains("BEFOREINLINEAFTERINLINE", string.Concat(pages.Select(FlowPictureText)));
        }
        if (width > 400D - left) {
            var bitmap = RenderFlowPicture(bytes, picturePage);
            // Paint reaches the physical page edge while retaining the full
            // authored image transform beyond it; it is not shrunk to fit.
            AssertFlowPicturePixel(bitmap, 390, (int)top + 8, 0, 255, 0);
        }
    }

    [Fact]
    public void SaveAsPdf_ExactLineHeightClipsOnlyOversizedPicturePaint() {
        byte[] bytes = ConvertQualifiedFlowPicture("picture-wide-mixed-exact24-explicit-defaults.docx");
        using var pdf = PdfPigDocument.Open(bytes);
        Page page = Assert.Single(pdf.GetPages());
        var picture = Assert.Single(page.GetImages());
        Assert.InRange(Math.Abs(picture.BoundingBox.Left - 50D), 0D, .15D);
        Assert.InRange(Math.Abs(page.Height - picture.BoundingBox.Top + 14.8D), 0D, .35D);
        Assert.Equal(360D, picture.BoundingBox.Width, 3);
        Assert.Equal(120D, picture.BoundingBox.Height, 3);
        Assert.Contains("BEFOREPICTUREBEFOREINLINEAFTERINLINEAFTERPICTURE", FlowPictureText(page));

        var bitmap = RenderFlowPicture(bytes, 1);
        // Native Word retains the 360x120 image transform, but clips its paint
        // to y62..105.16. Only the lower blue/yellow strip remains visible.
        AssertFlowPicturePixel(bitmap, 100, 68, 0, 0, 255);
        AssertFlowPicturePixel(bitmap, 270, 70, 255, 255, 0);
        AssertFlowPicturePixel(bitmap, 300, 50, 255, 255, 255);
        AssertFlowPicturePixel(bitmap, 300, 110, 255, 255, 255);
        AssertFlowPicturePixel(bitmap, 49, 70, 255, 255, 255);
        AssertFlowPicturePixel(bitmap, 398, 70, 255, 255, 0);

        // The following text lies below the image's clipping rectangle. A
        // paragraph-wide clip would retain extracted text but hide this paint.
        int blackPixels = 0;
        for (int y = 121; y < 131; y++) {
            for (int x = 50; x < 140; x++) {
                var pixel = bitmap.GetPixel(x, y);
                if (pixel.R < 40 && pixel.G < 40 && pixel.B < 40) blackPixels++;
            }
        }
        Assert.True(blackPixels > 20, "Following inline text must paint outside the picture-only clip.");
    }

    [Theory]
    [InlineData("picture-exact24-explicit-line2.docx", 2, -14.8D, 0, 0, 255, 129.26D)]
    [InlineData("picture-exact24-explicit-line3.docx", 3, 9.2D, 255, 0, 0, 153.29D)]
    public void SaveAsPdf_ExactHeightPictureOnLaterLinesRetainsNativeParagraphClip(
        string fixture, int pictureLine, double pictureTop, byte red, byte green, byte blue, double followingBaseline) {
        // Native Word exports of both explicit-break controls clip at the
        // paragraph top, allowing the picture to cover earlier paragraph lines.
        // A clip at the current line top would discard this qualified paint.
        byte[] bytes = ConvertQualifiedFlowPicture(fixture);
        using var pdf = PdfPigDocument.Open(bytes);
        Page page = Assert.Single(pdf.GetPages());
        var picture = Assert.Single(page.GetImages());
        Assert.InRange(Math.Abs(page.Height - picture.BoundingBox.Top - pictureTop), 0D, .15D);
        Assert.Equal(360D, picture.BoundingBox.Width, 3);
        Assert.Equal(120D, picture.BoundingBox.Height, 3);
        string text = FlowPictureText(page);
        Assert.Contains("BEFOREPICTUREFIRSTLINE", text);
        if (pictureLine == 3) Assert.Contains("SECONDLINE", text);
        Assert.Contains("AFTERINLINEAFTERPICTURE", text);
        var following = Assert.Single(page.GetWords(), word => word.Text == "INLINE");
        Assert.InRange(Math.Abs(page.Height - following.Letters[0].StartBaseLine.Y - followingBaseline), 0D, .15D);

        var bitmap = RenderFlowPicture(bytes, 1);
        AssertFlowPicturePixel(bitmap, 100, 65, red, green, blue);
        AssertFlowPicturePixel(bitmap, 100, 61, 255, 255, 255);
        AssertFlowPicturePixel(bitmap, 300, 50, 255, 255, 255);
        AssertFlowPicturePixel(bitmap, 300, (int)Math.Ceiling(pictureTop + 120D) + 4, 255, 255, 255);
        int followingTextPixels = 0;
        for (int y = (int)followingBaseline - 9; y < (int)followingBaseline + 2; y++) {
            for (int x = 50; x < 140; x++) {
                var pixel = bitmap.GetPixel(x, y);
                if (pixel.R < 40 && pixel.G < 40 && pixel.B < 40) followingTextPixels++;
            }
        }
        Assert.True(followingTextPixels > 20, "Text after the image must remain visible outside its clipping rectangle.");
    }

    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    public void SaveAsPdf_DecoratedBodyPictureRetainsNaturalSizeAcrossColumnAdmission(int columns) {
        string path = Path.Combine(AppContext.BaseDirectory, "WordLayoutFixtures",
            "picture-tall-explicit-defaults.docx");
        using WordDocument source = WordDocument.Load(path);
        source.Sections[0].ColumnCount = columns;
        source.Paragraphs[1].ShadingFillColorHex = "EFEFEF";
        source.Paragraphs[1].KeepLinesTogether = true;
        using WordDocument document = WordDocument.Load(new MemoryStream(source.ToBytes()));
        Assert.Empty(document.ValidateDocument());
        using var pdf = PdfPigDocument.Open(document.ToPdf(new WordPdfExportOptions { IncludePageNumbers = false }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        Page[] pages = pdf.GetPages().ToArray();
        var picture = Assert.Single(pages.SelectMany(page => page.GetImages()));
        Assert.Equal(160D, picture.BoundingBox.Width, 3);
        Assert.Equal(440D, picture.BoundingBox.Height, 3);
        string text = string.Concat(pages.Select(FlowPictureText));
        Assert.Contains("BEFOREPICTURE", text);
        Assert.Contains("AFTERPICTURE", text);
    }

    private static byte[] ConvertQualifiedFlowPicture(string fixture) {
        string path = Path.Combine(AppContext.BaseDirectory, "WordLayoutFixtures", fixture);
        byte[] source = File.ReadAllBytes(path);
        using (var package = WordprocessingDocument.Open(new MemoryStream(source), false)) {
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2013).Validate(package));
        }
        using WordDocument document = WordDocument.Load(path);
        var main = document._wordprocessingDocument.MainDocumentPart!;
        string xml = main.Document.OuterXml + main.DocumentSettingsPart!.Settings!.OuterXml;
        byte[] result = document.ToPdf(new WordPdfExportOptions { IncludePageNumbers = false }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });
        Assert.Equal(xml, main.Document.OuterXml + main.DocumentSettingsPart.Settings.OuterXml);
        Assert.Equal(source, File.ReadAllBytes(path));
        return result;
    }

    private static string FlowPictureText(Page page) => string.Concat(page.Letters
        .Where(letter => !string.IsNullOrWhiteSpace(letter.Value)).Select(letter => letter.Value));

    private static OfficeRasterImage RenderFlowPicture(byte[] bytes, int page) {
        var pdf = PdfCore.PdfDocument.Load(bytes);
        var raster = pdf.Render.Pages(PdfCore.PdfPageSelection.From(page),
            new PdfCore.PdfPageRenderOptions { Dpi = 72 })[0];
        Assert.True(OfficePngReader.TryDecode(raster.Bytes!, out var bitmap));
        return bitmap!;
    }

    private static void AssertFlowPicturePixel(OfficeRasterImage bitmap, int x, int y,
        byte red, byte green, byte blue) {
        var actual = bitmap.GetPixel(x, y);
        Assert.InRange(Math.Abs(actual.R - red), 0, 8);
        Assert.InRange(Math.Abs(actual.G - green), 0, 8);
        Assert.InRange(Math.Abs(actual.B - blue), 0, 8);
    }
}
