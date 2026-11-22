using OfficeIMO.Drawing;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PdfCore = OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using PdfRectangle = UglyToad.PdfPig.Core.PdfRectangle;

namespace OfficeIMO.Tests;

public sealed class PdfRunningStoryDrawingLayersTests {
    // As in Word, a header or footer drawing in front of text covers all of that story's text, while the
    // whole story still paints under the body text. A drawing behind text stays under all of it.
    [Theory]
    [InlineData("picture", true, true)]
    [InlineData("picture", false, true)]
    [InlineData("group", true, true)]
    [InlineData("group", false, true)]
    [InlineData("picture", true, false)]
    [InlineData("group", false, false)]
    public void SaveAsPdf_RunningStoryDrawingStacksAsInWord(string kind, bool header, bool inFront) {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph("BODYTEXT BODYTEXT BODYTEXT");
        for (int index = 0; index < 40; index++) document.AddParagraph("filler line");
        _ = document.PageSetup.Header;
        _ = document.PageSetup.Footer;
        WordHeaderFooter story = header ? document.PageSetup.GetHeader(WordHeaderFooterType.Default)!
            : document.PageSetup.GetFooter(WordHeaderFooterType.Default)!;
        WordParagraph anchor = story.AddParagraph("ANCHORLINE");
        // The drawing spans the story's two lines and the nearest body line.
        double top = header ? 30D : 700D;
        if (kind == "picture") {
            using var stream = new MemoryStream(OfficeRasterImageEncoder.Encode(new OfficeRasterImage(20, 20, OfficeColor.Red), OfficeImageExportFormat.Png));
            WordImage image = anchor.AddImage(stream, "red.png", OfficeLength.Points(200), OfficeLength.Points(70),
                options: new WordImageOptions { Wrapping = WordImageTextWrapping.InFrontOfText });
            image.HorizontalPositionRelativeFrom = WordHorizontalRelativePosition.Page;
            image.VerticalPositionRelativeFrom = WordVerticalRelativePosition.Page;
            image.HorizontalPositionOffset = OfficeLength.Points(60);
            image.VerticalPositionOffset = OfficeLength.Points(top);
        } else {
            anchor.AddShapeGroup(new[] {
                new WordShapeGroupItem(WordShapeType.Rectangle, OfficeLength.Points(0), OfficeLength.Points(0), OfficeLength.Points(100), OfficeLength.Points(70)) { FillColor = OfficeColor.Parse("FF0000") },
                new WordShapeGroupItem(WordShapeType.Rectangle, OfficeLength.Points(100), OfficeLength.Points(0), OfficeLength.Points(100), OfficeLength.Points(70)) { FillColor = OfficeColor.Parse("FF0000") }
            }, OfficeLength.Points(60), OfficeLength.Points(top));
        }
        DW.Anchor drawingAnchor = anchor._paragraph.Descendants<DW.Anchor>().Single();
        drawingAnchor.BehindDoc = !inFront;
        drawingAnchor.RemoveAllChildren<DW.WrapSquare>();
        if (drawingAnchor.GetFirstChild<DW.WrapNone>() == null) drawingAnchor.AddChild(new DW.WrapNone(), true);
        story.AddParagraph("FOLLOWINGLINE");

        var evidence = document.ConvertToPdfWithEvidence(new WordPdfExportOptions { IncludePageNumbers = false });
        Assert.DoesNotContain(evidence.Warnings, warning => warning.Code.StartsWith("native-shape-group-", StringComparison.Ordinal));
        byte[] bytes = evidence.Value.ToBytes(new PdfCore.PdfSaveOptions { AllowLoss = true });
        PdfRectangle anchorText, followingText, bodyText;
        using (var pig = PdfPigDocument.Open(bytes)) {
            var words = pig.GetPage(1).GetWords().ToList();
            anchorText = words.First(word => word.Text == "ANCHORLINE").BoundingBox;
            followingText = words.First(word => word.Text == "FOLLOWINGLINE").BoundingBox;
            var body = words.Where(word => word.Text is "BODYTEXT" or "filler" && word.BoundingBox.Left < 250D);
            bodyText = (header ? body.OrderByDescending(word => word.BoundingBox.Top) : body.OrderBy(word => word.BoundingBox.Bottom)).First().BoundingBox;
        }
        var pdf = PdfCore.PdfDocument.Load(new MemoryStream(bytes));
        byte[] rendered = pdf.Render.Pages(PdfCore.PdfPageSelection.From(1), new PdfCore.PdfPageRenderOptions { Dpi = 72 })[0].Bytes!;
        Assert.True(OfficePngReader.TryDecode(rendered, out var bitmap));

        Assert.Equal(!inFront, TextShowsThrough(bitmap!, anchorText));
        Assert.Equal(!inFront, TextShowsThrough(bitmap!, followingText));
        Assert.True(TextShowsThrough(bitmap!, bodyText), "The body text paints over the header and footer layer.");
    }

    // Counts dark text pixels inside a word's box on a 792pt page rendered at 72 dpi.
    private static bool TextShowsThrough(OfficeRasterImage bitmap, PdfRectangle box) {
        byte[] pixels = bitmap.GetPixels();
        int dark = 0;
        for (int y = (int)Math.Floor(792D - box.Top); y < (int)Math.Ceiling(792D - box.Bottom); y++) {
            for (int x = (int)Math.Floor(box.Left); x < (int)Math.Ceiling(box.Right); x++) {
                if (x < 0 || y < 0 || x >= bitmap.Width || y >= bitmap.Height) continue;
                int offset = (y * bitmap.Width + x) * 4;
                if (pixels[offset] < 150) dark++;
            }
        }
        return dark > 3;
    }
}
