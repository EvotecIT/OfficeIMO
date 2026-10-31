using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData("h01-inline-picture-before-text.docx", false)]
    [InlineData("h02-inline-picture-after-text.docx", false)]
    [InlineData("h03-right-tab-picture-title.docx", true)]
    [InlineData("legacy11.docx", false)]
    [InlineData("legacy14.docx", false)]
    public void SaveAsPdf_ExactRunningHeaderClipsOnlyInlinePicturePaint(string fixture, bool rightTab) {
        // Independent OOXML controls were admitted by the SDK and exported by
        // native Word. The compatibility11/14 controls change only settings;
        // all five native exports retain the full 24x12 picture
        // frame at y15.6..27.6, with visible paint clipped to y18..27.6.
        // The line's baseline remains 27.6pt from the page top.
        string path = Path.Combine(AppContext.BaseDirectory, "WordLayoutFixtures", fixture);
        byte[] source = File.ReadAllBytes(path);
        using WordDocument document = WordDocument.Load(path);
        Assert.Empty(document.ValidateDocument());
        var main = document._wordprocessingDocument.MainDocumentPart!;
        var header = Assert.Single(main.HeaderParts).Header;
        string originalBody = main.Document.OuterXml;
        string originalHeader = header.OuterXml;

        byte[] bytes = document.ToPdf(new WordPdfExportOptions { IncludePageNumbers = false }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });
        PdfImagePlacement image = Assert.Single(PdfImageExtractor.ExtractImagePlacements(bytes));
        Assert.Equal(24D, image.Width, 3);
        Assert.Equal(12D, image.Height, 3);
        using var pdf = PdfPigDocument.Open(bytes);
        var page = Assert.Single(pdf.GetPages());
        Assert.Equal(240D, page.Width, 3);
        Assert.Equal(320D, page.Height, 3);
        Assert.Equal(15.6D, page.Height - image.Y - image.Height, 2);

        var words = page.GetWords().ToArray();
        var title = Assert.Single(words, word => word.Text == "TITLE");
        var firstTitleWord = Assert.Single(words, word => word.Text == (rightTab ? "RIGHT" : "INLINE"));
        Assert.Equal(27.6D, page.Height - title.Letters[0].StartBaseLine.Y, 2);
        Assert.Equal(firstTitleWord.Letters[0].StartBaseLine.Y, title.Letters[0].StartBaseLine.Y, 3);
        Assert.Single(words, word => word.Text == "INTRO");
        if (fixture == "h02-inline-picture-after-text.docx") {
            Assert.True(title.Letters.Last().EndBaseLine.X <= image.X + .05D);
        } else {
            Assert.Equal(24D, image.X, 3);
            Assert.True(firstTitleWord.Letters[0].StartBaseLine.X >= image.X + image.Width - .05D);
        }
        if (rightTab) {
            TabStop stop = Assert.Single(header.Descendants<TabStop>(), tab => tab.Val?.Value == TabStopValues.Right);
            double tabTarget = checked((uint)(document.Sections[0].Margins.Left.GetValueOrDefault().ToTwips())) / 20D + stop.Position!.Value / 20D;
            Assert.Equal(216D, tabTarget, 3);
            // Right tabs align the advance endpoint, independently of the
            // actual font's ink overhang or the title's measured width.
            Assert.InRange(Math.Abs(title.Letters.Last().EndBaseLine.X - tabTarget), 0D, .3D);
        }

        PdfImageClipInfo clip = Assert.IsType<PdfImageClipInfo>(image.Clip);
        Assert.True(clip.IsRectangle && clip.IsExact);
        Assert.False(clip.ContainsTextClipping);
        // Content operators serialize coordinates to millipoints; compare
        // their distance instead of rounding opposite sides of a midpoint.
        Assert.InRange(Math.Abs(image.X - clip.X), 0D, .001D);
        Assert.Equal(image.Width, clip.Width, 2);
        Assert.Equal(18D, clip.Y, 2);
        Assert.Equal(9.6D, clip.Height, 2);
        Assert.Equal(27.6D, clip.Y + clip.Height, 2);
        Assert.Equal(page.Height - image.Y, clip.Y + clip.Height, 2);

        // The shared helper renders at 72dpi, so these source point positions
        // are pixel positions. Interior samples avoid antialiased clip edges.
        var bitmap = RenderFlowPicture(bytes, image.PageNumber);
        int blueX = (int)Math.Floor(image.X + image.Width / 2D);
        int goldX = (int)Math.Floor(image.X + 2D);
        AssertFlowPicturePixel(bitmap, blueX, 16, 255, 255, 255);
        AssertFlowPicturePixel(bitmap, blueX, 22, 35, 104, 163);
        AssertFlowPicturePixel(bitmap, goldX, 22, 232, 161, 45);

        // A clip around the remaining paragraph could preserve extracted text
        // while hiding TITLE after the image. Inspect its actual painted ink.
        int left = Math.Max(0, (int)Math.Floor(title.Letters.Min(letter => letter.BoundingBox.Left)));
        int right = Math.Min(bitmap.Width, (int)Math.Ceiling(title.Letters.Max(letter => letter.BoundingBox.Right)));
        int top = Math.Max(0, (int)Math.Floor(page.Height - title.Letters.Max(letter => letter.BoundingBox.Top)));
        int bottom = Math.Min(bitmap.Height, (int)Math.Ceiling(page.Height - title.Letters.Min(letter => letter.BoundingBox.Bottom)));
        int titleInk = 0;
        for (int y = top; y < bottom; y++) {
            for (int x = left; x < right; x++) {
                var pixel = bitmap.GetPixel(x, y);
                if (pixel.R < 60 && pixel.G < 60 && pixel.B < 60) titleInk++;
            }
        }
        Assert.True(titleInk > 0, "The title must paint outside the restored picture-only clip.");
        Assert.Equal(originalBody, main.Document.OuterXml);
        Assert.Equal(originalHeader, header.OuterXml);
        Assert.Equal(source, File.ReadAllBytes(path));
        Assert.Empty(document.ValidateDocument());
    }

    [Theory]
    [InlineData("h1-vml-mode15.docx", false, 15.6D, 27.6D)]
    [InlineData("f1-vml-mode15.docx", true, 287.6D, 299.6D)]
    public void SaveAsPdf_InlineVmlRunningPicturePreservesSourceRunAndFullPaint(
        string fixture, bool footer, double imageTop, double baseline) {
        // Each control changes only the drawing carrier. Native Word retains
        // the VML picture's full paint rather than the DrawingML exact-line clip.
        // The assertions below use the authored nominal frame. The native
        // footer export differs by .05pt in image placement and .136pt in the
        // title baseline; this test does not claim that print-metric precision.
        string path = Path.Combine(AppContext.BaseDirectory, "WordLayoutFixtures", fixture);
        byte[] source = File.ReadAllBytes(path);
        using WordDocument document = WordDocument.Load(path);
        Assert.Empty(document.ValidateDocument());
        string body = document._wordprocessingDocument.MainDocumentPart!.Document.OuterXml;
        string story = footer
            ? Assert.Single(document._wordprocessingDocument.MainDocumentPart.FooterParts).Footer.OuterXml
            : Assert.Single(document._wordprocessingDocument.MainDocumentPart.HeaderParts).Header.OuterXml;

        byte[] bytes = document.ToPdf(new WordPdfExportOptions { IncludePageNumbers = false }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });
        using var pdf = PdfPigDocument.Open(bytes);
        var page = Assert.Single(pdf.GetPages());
        var words = page.GetWords().ToArray();
        Assert.Equal(footer ? new[] { "INTRO", "INLINE", "TITLE" } : new[] { "INLINE", "TITLE", "INTRO" },
            words.Select(word => word.Text));
        PdfImagePlacement image = Assert.Single(PdfImageExtractor.ExtractImagePlacements(bytes));
        Assert.Equal(24D, image.X, 3);
        Assert.Equal(24D, image.Width, 3);
        Assert.Equal(12D, image.Height, 3);
        Assert.Equal(imageTop, page.Height - image.Y - image.Height, 2);
        var title = Assert.Single(words, word => word.Text == "TITLE");
        var firstTitleWord = Assert.Single(words, word => word.Text == "INLINE");
        Assert.Equal(baseline, page.Height - title.Letters[0].StartBaseLine.Y, 2);
        Assert.Equal(firstTitleWord.Letters[0].StartBaseLine.Y, title.Letters[0].StartBaseLine.Y, 3);
        Assert.True(firstTitleWord.Letters[0].StartBaseLine.X >= image.X + image.Width - .05D);

        var bitmap = RenderFlowPicture(bytes, image.PageNumber);
        int blueX = (int)Math.Floor(image.X + image.Width / 2D);
        AssertFlowPicturePixel(bitmap, blueX, (int)Math.Floor(imageTop + 1D), 35, 104, 163);
        AssertFlowPicturePixel(bitmap, blueX, (int)Math.Floor(imageTop + 7D), 35, 104, 163);
        Assert.Equal(body, document._wordprocessingDocument.MainDocumentPart.Document.OuterXml);
        Assert.Equal(story, footer
            ? Assert.Single(document._wordprocessingDocument.MainDocumentPart.FooterParts).Footer.OuterXml
            : Assert.Single(document._wordprocessingDocument.MainDocumentPart.HeaderParts).Header.OuterXml);
        Assert.Equal(source, File.ReadAllBytes(path));
        Assert.Empty(document.ValidateDocument());
    }
}
