using DocumentFormat.OpenXml;
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(false, true, 1)]
    [InlineData(true, true, 1)]
    [InlineData(false, false, 1)]
    [InlineData(true, false, 1)]
    [InlineData(false, true, 2)]
    [InlineData(true, true, 2)]
    [InlineData(false, true, 3)]
    [InlineData(true, true, 3)]
    public void SaveAsPdf_NoteInlinePictureRetainsItsTransformAndClipsOnlyExactLinePaint(bool endnote, bool exact, int pictureLine) {
        using WordDocument source = CreateFootnoteDiagnosticDocument();
        OpenXmlCompositeElement note = AddRichNoteFixture(source, endnote);
        note.RemoveAllChildren();
        var paragraph = new W.Paragraph(new W.ParagraphProperties(new W.SpacingBetweenLines {
            Before = "0", After = "0", Line = "240",
            LineRule = exact ? W.LineSpacingRuleValues.Exact : W.LineSpacingRuleValues.Auto }),
            NoteBodyMarker(endnote));
        for (int line = 1; line < pictureLine; line++)
            paragraph.Append(RichNoteRun("EARLIERLINE" + line), new W.Run(new W.Break()));
        paragraph.Append(RichNoteRun("BEFOREIMAGE ", new W.Bold()));
        note.Append(paragraph);
        var authored = new WordParagraph(source, paragraph);
        using (var pixels = new MemoryStream(OfficeIMO.Tests.Pdf.PdfPngTestImages.CreateRgbPng((byte)0, (byte)180, (byte)80)))
            authored.AddText(string.Empty).AddImage(pixels, "owned-note-image.png", 64, 32);
        authored.AddText(" AFTERIMAGE").Italic = true;
        note.Append(new W.Paragraph(RichNoteRun("FOLLOWINGNOTEPARAGRAPH")));
        Assert.Empty(source.ValidateDocument());

        using WordDocument document = WordDocument.Load(new MemoryStream(source.ToBytes()));
        Assert.Empty(document.ValidateDocument());
        var main = document._wordprocessingDocument.MainDocumentPart!;
        OpenXmlCompositeElement loaded = endnote ? main.EndnotesPart!.Endnotes! : main.FootnotesPart!.Footnotes!;
        string originalBody = main.Document.OuterXml;
        string originalNotes = loaded.OuterXml;
        byte[] bytes = document.ToPdf(FootnoteDiagnosticOptions()).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });
        PdfImagePlacement image = Assert.Single(PdfImageExtractor.ExtractImagePlacements(bytes));
        Assert.Equal(48D, image.Width, 3);
        Assert.Equal(24D, image.Height, 3);
        using var pdf = PdfPigDocument.Open(bytes);
        var page = pdf.GetPage(image.PageNumber);
        var letters = page.Letters.ToArray();
        var before = FindRichNoteToken(letters, "BEFOREIMAGE");
        var after = FindRichNoteToken(letters, "AFTERIMAGE");
        FindRichNoteToken(letters, "FOLLOWINGNOTEPARAGRAPH");
        Assert.True(before[before.Length - 1].EndBaseLine.X <= image.X + .01D);
        Assert.True(after[0].StartBaseLine.X >= image.X + image.Width - .01D);
        Assert.Equal(before[0].StartBaseLine.Y, after[0].StartBaseLine.Y, 3);
        var bitmap = RenderFlowPicture(bytes, image.PageNumber);
        int imageCenterX = (int)Math.Floor(image.X + image.Width / 2D);
        double imageTop = page.Height - image.Y - image.Height;
        if (exact) {
            PdfImageClipInfo clip = Assert.IsType<PdfImageClipInfo>(image.Clip);
            Assert.True(clip.IsRectangle && clip.IsExact);
            Assert.Equal(image.Width, clip.Width, 3);
            // Later lines may paint over earlier lines in this same paragraph;
            // the current line top must not replace the paragraph-fragment top.
            Assert.InRange(clip.Height, 12D * (pictureLine - 1) + .01D, 12D * pictureLine + .01D);
            Assert.Equal(page.Height - image.Y, clip.Y + clip.Height, 3);
            if (clip.Y > imageTop + .01D) {
                var removed = bitmap.GetPixel(imageCenterX,
                    (int)Math.Floor(imageTop + Math.Min(2D, (clip.Y - imageTop) / 2D)));
                Assert.False(removed.G > 150 && removed.R < 30 && removed.B < 110,
                    "The image must not paint above its paragraph fragment into the preceding note/separator area.");
            } else AssertFlowPicturePixel(bitmap, imageCenterX, (int)Math.Floor(imageTop + 2D), 0, 180, 80);
            AssertFlowPicturePixel(bitmap, imageCenterX, (int)Math.Floor(clip.Y + clip.Height / 2D), 0, 180, 80);
        } else {
            Assert.Null(image.Clip);
            AssertFlowPicturePixel(bitmap, imageCenterX, (int)Math.Floor(imageTop + 2D), 0, 180, 80);
        }
        // Keep the clip local to the picture. Extracted text alone would not
        // detect an accidentally clipped following inline run.
        int left = Math.Max(0, (int)Math.Floor(after.Min(letter => letter.BoundingBox.Left)));
        int right = Math.Min(bitmap.Width, (int)Math.Ceiling(after.Max(letter => letter.BoundingBox.Right)));
        int top = Math.Max(0, (int)Math.Floor(page.Height - after.Max(letter => letter.BoundingBox.Top)));
        int bottom = Math.Min(bitmap.Height, (int)Math.Ceiling(page.Height - after.Min(letter => letter.BoundingBox.Bottom)));
        int followingInk = 0;
        for (int y = top; y < bottom; y++)
            for (int x = left; x < right; x++) {
                var pixel = bitmap.GetPixel(x, y);
                if (pixel.R < 60 && pixel.G < 60 && pixel.B < 60) followingInk++;
            }
        Assert.True(followingInk > 0, "Following inline text must still paint after the picture-only clip is restored.");
        Assert.Equal(originalBody, main.Document.OuterXml);
        Assert.Equal(originalNotes, loaded.OuterXml);
        Assert.Empty(document.ValidateDocument());
    }
}
