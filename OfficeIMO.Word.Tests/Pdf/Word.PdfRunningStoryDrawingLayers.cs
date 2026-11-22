using OfficeIMO.Drawing;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using V = DocumentFormat.OpenXml.Vml;
using W = DocumentFormat.OpenXml.Wordprocessing;
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
    // A paragraph that also holds a text box takes the mixed text-box route.
    [InlineData("groupbox", true, true)]
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
        if (kind == "groupbox") {
            WordTextBox box = anchor.AddTextBox("BOXTEXT", WordImageTextWrapping.InFrontOfText);
            box.HorizontalPositionRelativeFrom = WordHorizontalRelativePosition.Page;
            box.HorizontalPositionOffset = OfficeLength.Points(400);
            box.VerticalPositionRelativeFrom = WordVerticalRelativePosition.Page;
            box.VerticalPositionOffset = OfficeLength.Points(top);
            box.Width = OfficeLength.Points(120);
            box.Height = OfficeLength.Points(30);
            box.RelativeWidthPercentage = 0;
            box.RelativeHeightPercentage = 0;
        }
        story.AddParagraph("FOLLOWINGLINE");
        // A table keeps the story on the rich route, which sends a paragraph with a text box through the mixed text-box route.
        if (kind == "groupbox") story.AddTable(1, 1);

        var evidence = document.ConvertToPdfWithEvidence(new WordPdfExportOptions { IncludePageNumbers = false });
        Assert.DoesNotContain(evidence.Warnings, warning => warning.Code.StartsWith("native-shape-group-", StringComparison.Ordinal) ||
            warning.Code == "native-header-footer-shape-group-unsupported");
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
        // The mixed route flows its text box, which moves the following line below the group.
        if (kind != "groupbox") Assert.Equal(!inFront, TextShowsThrough(bitmap!, followingText));
        Assert.True(TextShowsThrough(bitmap!, bodyText), "The body text paints over the header and footer layer.");
    }

    // A header group that is not drawn, placed relative to the margin or stored only as VML, is reported as
    // omitted. A group in a hidden run is not painted in Word either, so nothing is lost.
    [Theory]
    [InlineData("margin", false)]
    [InlineData("vml", false)]
    [InlineData("margin", true)]
    [InlineData("vml", true)]
    // The VML fallback Word selects, because no choice applies, is the group itself.
    [InlineData("vml-fallback", false)]
    // A drawn group beside a selected VML fallback the rich route cannot place keeps the story on the simple route.
    [InlineData("mixed", false)]
    public void SaveAsPdf_UndrawnHeaderShapeGroupIsReported(string kind, bool hidden) {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph("BODYTEXT");
        _ = document.PageSetup.Header;
        WordParagraph anchor = document.PageSetup.GetHeader(WordHeaderFooterType.Default)!.AddParagraph("ANCHORLINE");
        if (kind == "mixed") {
            anchor.AddShapeGroup(new[] {
                new WordShapeGroupItem(WordShapeType.Rectangle, OfficeLength.Points(0), OfficeLength.Points(0), OfficeLength.Points(100), OfficeLength.Points(70)) { FillColor = OfficeColor.Parse("FF0000") },
                new WordShapeGroupItem(WordShapeType.Rectangle, OfficeLength.Points(100), OfficeLength.Points(0), OfficeLength.Points(100), OfficeLength.Points(70)) { FillColor = OfficeColor.Parse("FF0000") }
            }, OfficeLength.Points(60), OfficeLength.Points(30));
            DW.Anchor drawingAnchor = anchor._paragraph.Descendants<DW.Anchor>().Single();
            drawingAnchor.RemoveAllChildren<DW.WrapSquare>();
            if (drawingAnchor.GetFirstChild<DW.WrapNone>() == null) drawingAnchor.AddChild(new DW.WrapNone(), true);
        }
        if (kind is "vml" or "vml-fallback" or "mixed") {
            var group = new V.Group {
                Style = "position:absolute;margin-left:60pt;margin-top:30pt;width:200pt;height:70pt;z-index:1;mso-position-horizontal-relative:page;mso-position-vertical-relative:page",
                CoordinateSize = "200,70"
            };
            group.Append(new V.Rectangle { Style = "position:absolute;left:0;top:0;width:200;height:70", FillColor = "red", Stroked = false });
            if (kind == "vml") anchor._paragraph.Append(new W.Run(new W.Picture(group)));
            else {
                var choice = new DocumentFormat.OpenXml.AlternateContentChoice { Requires = "x9" };
                choice.AddNamespaceDeclaration("x9", "urn:officeimo:unsupported-choice");
                // As Word stores it: the DrawingML copy in the choice, which this reader does not select.
                if (kind == "mixed") choice.Append(anchor._paragraph.Descendants<W.Drawing>().Single().CloneNode(true));
                anchor._paragraph.Append(new W.Run(new DocumentFormat.OpenXml.AlternateContent(choice,
                    new DocumentFormat.OpenXml.AlternateContentFallback(new W.Picture(group)))));
            }
        } else {
            anchor.AddShapeGroup(new[] {
                new WordShapeGroupItem(WordShapeType.Rectangle, OfficeLength.Points(0), OfficeLength.Points(0), OfficeLength.Points(100), OfficeLength.Points(70)) { FillColor = OfficeColor.Parse("FF0000") },
                new WordShapeGroupItem(WordShapeType.Rectangle, OfficeLength.Points(100), OfficeLength.Points(0), OfficeLength.Points(100), OfficeLength.Points(70)) { FillColor = OfficeColor.Parse("FF0000") }
            }, OfficeLength.Points(60), OfficeLength.Points(30));
            anchor._paragraph.Descendants<DW.Anchor>().Single().HorizontalPosition!.RelativeFrom = DW.HorizontalRelativePositionValues.Margin;
        }
        if (hidden) {
            W.Run groupRun = anchor._paragraph.Elements<W.Run>().Last();
            groupRun.PrependChild(new W.RunProperties(new W.Vanish()));
        }

        var evidence = document.ConvertToPdfWithEvidence(new WordPdfExportOptions { IncludePageNumbers = false });
        if (hidden) {
            Assert.DoesNotContain(evidence.Warnings, item => item.Code == "native-header-footer-shape-group-unsupported");
            return;
        }
        var warning = Assert.Single(evidence.Warnings, item => item.Code == "native-header-footer-shape-group-unsupported");
        Assert.Equal("default header", warning.Source);
        Assert.Equal(OfficeIMO.OfficeConversionLossKind.Omission, warning.LossKind);
    }

    // A supported group in a hidden run is not painted, as in Word.
    [Fact]
    public void SaveAsPdf_HiddenHeaderShapeGroupIsNotPainted() {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph("BODYTEXT");
        _ = document.PageSetup.Header;
        WordParagraph anchor = document.PageSetup.GetHeader(WordHeaderFooterType.Default)!.AddParagraph("ANCHORLINE");
        anchor.AddShapeGroup(new[] {
            new WordShapeGroupItem(WordShapeType.Rectangle, OfficeLength.Points(0), OfficeLength.Points(0), OfficeLength.Points(100), OfficeLength.Points(70)) { FillColor = OfficeColor.Parse("FF0000") },
            new WordShapeGroupItem(WordShapeType.Rectangle, OfficeLength.Points(100), OfficeLength.Points(0), OfficeLength.Points(100), OfficeLength.Points(70)) { FillColor = OfficeColor.Parse("FF0000") }
        }, OfficeLength.Points(60), OfficeLength.Points(30));
        DW.Anchor drawingAnchor = anchor._paragraph.Descendants<DW.Anchor>().Single();
        drawingAnchor.BehindDoc = false;
        drawingAnchor.RemoveAllChildren<DW.WrapSquare>();
        if (drawingAnchor.GetFirstChild<DW.WrapNone>() == null) drawingAnchor.AddChild(new DW.WrapNone(), true);
        anchor._paragraph.Elements<W.Run>().Last().PrependChild(new W.RunProperties(new W.Vanish()));

        var evidence = document.ConvertToPdfWithEvidence(new WordPdfExportOptions { IncludePageNumbers = false });
        Assert.DoesNotContain(evidence.Warnings, item => item.Code == "native-header-footer-shape-group-unsupported");
        byte[] bytes = evidence.Value.ToBytes(new PdfCore.PdfSaveOptions { AllowLoss = true });
        var pdf = PdfCore.PdfDocument.Load(new MemoryStream(bytes));
        byte[] rendered = pdf.Render.Pages(PdfCore.PdfPageSelection.From(1), new PdfCore.PdfPageRenderOptions { Dpi = 72 })[0].Bytes!;
        Assert.True(OfficePngReader.TryDecode(rendered, out var bitmap));
        byte[] pixels = bitmap!.GetPixels();
        int offset = (90 * bitmap.Width + 240) * 4;
        Assert.Equal("255,255,255", pixels[offset] + "," + pixels[offset + 1] + "," + pixels[offset + 2]);
    }

    // A wrapped header picture and an in-front one share the story's in-front layer, so the higher z-order wins where they overlap.
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void SaveAsPdf_WrappedHeaderPictureKeepsItsZOrderOverInFrontDrawings(bool wrappedOnTop) {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph("BODYTEXT");
        _ = document.PageSetup.Header;
        WordParagraph anchor = document.PageSetup.GetHeader(WordHeaderFooterType.Default)!.AddParagraph("ANCHORLINE");
        AddPagePicture(anchor, OfficeColor.Red, WordImageTextWrapping.InFrontOfText, 60D, 30D, wrappedOnTop ? 1U : 2U);
        AddPagePicture(anchor, OfficeColor.Parse("00FF00"), WordImageTextWrapping.Square, 110D, 40D, wrappedOnTop ? 2U : 1U);

        byte[] bytes = document.ToPdf(new WordPdfExportOptions { IncludePageNumbers = false }).ToBytes(new PdfCore.PdfSaveOptions { AllowLoss = true });
        var pdf = PdfCore.PdfDocument.Load(new MemoryStream(bytes));
        byte[] rendered = pdf.Render.Pages(PdfCore.PdfPageSelection.From(1), new PdfCore.PdfPageRenderOptions { Dpi = 72 })[0].Bytes!;
        Assert.True(OfficePngReader.TryDecode(rendered, out var bitmap));
        byte[] pixels = bitmap!.GetPixels();
        int offset = (65 * bitmap.Width + 135) * 4;
        Assert.Equal(wrappedOnTop ? "0,255,0" : "255,0,0", pixels[offset] + "," + pixels[offset + 1] + "," + pixels[offset + 2]);
    }

    private static void AddPagePicture(WordParagraph paragraph, OfficeColor color, WordImageTextWrapping wrapping,
        double left, double top, uint zOrder) {
        using var stream = new MemoryStream(OfficeRasterImageEncoder.Encode(new OfficeRasterImage(20, 20, color), OfficeImageExportFormat.Png));
        WordImage image = paragraph.AddImage(stream, "fill.png", OfficeLength.Points(100), OfficeLength.Points(60),
            options: new WordImageOptions { Wrapping = wrapping });
        image.HorizontalPositionRelativeFrom = WordHorizontalRelativePosition.Page;
        image.VerticalPositionRelativeFrom = WordVerticalRelativePosition.Page;
        image.HorizontalPositionOffset = OfficeLength.Points(left);
        image.VerticalPositionOffset = OfficeLength.Points(top);
        image._Image!.Anchor!.RelativeHeight = zOrder;
        // Rectangular wrapping needs zero effect extents.
        image._Image.Anchor.GetFirstChild<DW.EffectExtent>()!.BottomEdge = 0L;
        image._Image.Anchor.BehindDoc = false;
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
