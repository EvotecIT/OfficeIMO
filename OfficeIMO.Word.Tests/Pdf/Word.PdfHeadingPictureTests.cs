using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Drawing;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PdfCore = OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData("Heading1", "inline")]
    [InlineData("Title", "inline")]
    [InlineData("Heading1", "whitespace")]
    [InlineData("Title", "whitespace")]
    [InlineData("Heading1", "anchor")]
    [InlineData("Title", "anchor")]
    [InlineData("Heading1", "picture-control")]
    [InlineData("Title", "picture-control")]
    public void SaveAsPdf_ImageOnlyHeadingPictureRetainsOrdinaryParagraphFrame(string style, string content) {
        using WordDocument source = CreateHeadingPictureDocument(style, content);
        using WordDocument reference = CreateHeadingPictureDocument("Normal", content);
        byte[] bytes = ConvertHeadingPictureDocument(source, $"image-{style}-{content}");
        byte[] expected = ConvertHeadingPictureDocument(reference, $"image-Normal-{content}");

        AssertHeadingPictureFramesMatch(bytes, expected, content == "anchor" ? 2 : 1);
        Assert.Empty(PdfCore.PdfDocument.Load(new System.IO.MemoryStream(bytes)).Outlines);
        Assert.Empty(PdfCore.PdfInspector.Inspect(bytes).FormFields);
    }

    [Theory]
    [InlineData("Heading1")]
    [InlineData("Title")]
    public void SaveAsPdf_InitialHeadingPictureRetainsAuthoredBeforeSpacing(string style) {
        using WordDocument source = CreateHeadingPictureDocument(style, "inline", atDocumentStart: true);
        using WordDocument reference = CreateHeadingPictureDocument("Normal", "inline", atDocumentStart: true);
        byte[] bytes = ConvertHeadingPictureDocument(source, $"initial-{style}");
        AssertHeadingPictureFramesMatch(bytes, ConvertHeadingPictureDocument(reference, "initial-Normal"), 1);
        Assert.Empty(PdfCore.PdfDocument.Load(new System.IO.MemoryStream(bytes)).Outlines);
    }

    [Theory]
    [InlineData("Heading1", false)]
    [InlineData("Title", false)]
    [InlineData("Heading1", true)]
    [InlineData("Title", true)]
    public void SaveAsPdf_MixedHeadingPictureRetainsTextAndControl(string style, bool cachedControl) {
        string content = cachedControl ? "date-control" : "text";
        using WordDocument source = CreateHeadingPictureDocument(style, content);
        using WordDocument reference = CreateHeadingPictureDocument("Normal", content);
        byte[] bytes = ConvertHeadingPictureDocument(source, $"mixed-{style}-{content}");
        byte[] expected = ConvertHeadingPictureDocument(reference, $"mixed-Normal-{content}");

        AssertHeadingPictureFramesMatch(bytes, expected, 1);
        using var visible = PdfPigDocument.Open(bytes);
        string text = Assert.Single(visible.GetPages()).Text;
        Assert.Contains("LEAD", text);
        Assert.Contains("TAIL", text);
        if (cachedControl) {
            PdfCore.PdfFormField field = Assert.Single(PdfCore.PdfInspector.Inspect(bytes).FormFields);
            Assert.Equal("HeadingPictureDate", field.Name);
            Assert.True(field.IsTextField);
            Assert.Equal("2026-10-20", field.Value);
            AssertHeadingPicturePaintMatches(bytes, expected);
        } else Assert.Empty(PdfCore.PdfInspector.Inspect(bytes).FormFields);

        // Widget appearance text is deliberately absent from raw page text, so
        // inspect the authored bookmark rather than infer headings from that text.
        var headings = PdfCore.PdfReadDocument.Open(bytes).Outlines;
        if (style == "Heading1") {
            var heading = Assert.Single(headings);
            Assert.Equal(1, heading.Level);
            Assert.Equal(1, heading.PageNumber);
            Assert.Contains("LEAD", heading.Title);
            Assert.Contains("TAIL", heading.Title);
            if (cachedControl) Assert.Contains("CACHED", heading.Title);
        } else Assert.Empty(headings);
    }

    private static WordDocument CreateHeadingPictureDocument(string style, string content, bool atDocumentStart = false) {
        WordDocument document = WordDocument.Create();
        SetObjectMarkNormalStyle(document);
        if (!atDocumentStart) document.AddParagraph("BEFORE");
        if (style == "Title")
            document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!.Append(
                new W.Style(new W.StyleName { Val = "Title" }, new W.BasedOn { Val = "Normal" }) {
                    StyleId = "Title", Type = W.StyleValues.Paragraph
                });
        WordParagraph paragraph = document.AddParagraph().SetStyleId(style);
        W.ParagraphStyleId styleId = (W.ParagraphStyleId)paragraph._paragraph.ParagraphProperties!.ParagraphStyleId!.CloneNode(true);
        paragraph._paragraph.ParagraphProperties = new W.ParagraphProperties(styleId,
            new W.SpacingBetweenLines { Before = atDocumentStart ? "400" : "0", After = "0", Line = "240", LineRule = W.LineSpacingRuleValues.Auto },
            new W.ParagraphMarkRunProperties(new W.RunFonts { Ascii = "Arial", HighAnsi = "Arial" },
                new W.Bold { Val = false }, new W.FontSize { Val = "24" }));
        if (content is "text" or "date-control") paragraph.AddText("LEAD ");
        if (content == "whitespace") paragraph.AddText(" \t ");
        byte[] picture = OfficeRasterImageEncoder.Encode(new OfficeRasterImage(64, 32, OfficeColor.Green), OfficeImageExportFormat.Png);
        using (var image = new MemoryStream(picture)) {
            WordImage inserted = paragraph.AddImage(image, "heading-picture.png", OfficeIMO.OfficeLength.Pixels(64), OfficeIMO.OfficeLength.Pixels(32));
            if (content == "picture-control") {
                // Exercise the separately owned picture-SDT path with the same
                // valid inline image, without a temporary filesystem asset.
                W.Run imageRun = (W.Run)inserted._Image!.Parent!;
                imageRun.Remove();
                paragraph._paragraph.Append(new W.SdtRun(
                    new W.SdtProperties(new W.SdtId { Val = 731 }, new W.SdtContentPicture()),
                    new W.SdtContentRun(imageRun)));
            }
        }
        if (content == "anchor") {
            byte[] pictureBytes = OfficeRasterImageEncoder.Encode(new OfficeRasterImage(16, 16, OfficeColor.Blue), OfficeImageExportFormat.Png);
            using var image = new MemoryStream(pictureBytes);
            WordImage anchored = paragraph.AddImage(image, "heading-anchor.png", OfficeIMO.OfficeLength.Pixels(16), OfficeIMO.OfficeLength.Pixels(16), options: new OfficeIMO.Word.WordImageOptions { Wrapping = WordImageTextWrapping.InFrontOfText });
            anchored.HorizontalPositionRelativeFrom = WordHorizontalRelativePosition.Page;
            anchored.VerticalPositionRelativeFrom = WordVerticalRelativePosition.Page;
            anchored.HorizontalPositionOffset = 240 * 12700;
            anchored.VerticalPositionOffset = 160 * 12700;
        }
        if (content == "date-control") {
            WordDatePicker date = paragraph.AddDatePicker(new DateTime(2026, 10, 20), tag: "HeadingPictureDate");
            date._sdtRun.SdtContentRun = new W.SdtContentRun(new W.Run(new W.Text("CACHED")));
        }
        if (content is "text" or "date-control") paragraph.AddText(" TAIL");
        foreach (W.Run run in paragraph._paragraph.Descendants<W.Run>())
            run.RunProperties = new W.RunProperties(new W.RunFonts { Ascii = "Arial", HighAnsi = "Arial" },
                new W.Bold { Val = false }, new W.Color { Val = "000000" }, new W.FontSize { Val = "24" });
        document.AddParagraph("FOLLOW");
        return document;
    }

    private static byte[] ConvertHeadingPictureDocument(WordDocument source, string artifactName) {
        byte[] bytes = source.ToBytes();
        using (WordprocessingDocument package = WordprocessingDocument.Open(new MemoryStream(bytes), false)) {
            ValidationErrorInfo[] errors = new OpenXmlValidator(FileFormatVersions.Office2016).Validate(package).ToArray();
            Assert.True(errors.Length == 0, string.Join("\n", errors.Select(error => error.Description + " @ " + error.Path?.XPath)));
        }
        using WordDocument loaded = WordDocument.Load(new MemoryStream(bytes));
        byte[] pdf = loaded.ToPdf(new WordPdfExportOptions {
            IncludePageNumbers = false, ResourcePolicy = PdfCore.PdfResourcePolicy.CreatePortableDeterministic()
        }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });
        string? output = Environment.GetEnvironmentVariable("OFFICEIMO_PDF_HEADING_PICTURE_OUTPUT");
        if (!string.IsNullOrWhiteSpace(output)) {
            Directory.CreateDirectory(output);
            File.WriteAllBytes(Path.Combine(output, artifactName + ".docx"), bytes);
            File.WriteAllBytes(Path.Combine(output, artifactName + ".pdf"), pdf);
        }
        return pdf;
    }

    private static void AssertHeadingPictureFramesMatch(byte[] actualBytes, byte[] expectedBytes, int pictureCount) {
        using var actual = PdfPigDocument.Open(actualBytes);
        using var expected = PdfPigDocument.Open(expectedBytes);
        var actualPage = Assert.Single(actual.GetPages());
        var expectedPage = Assert.Single(expected.GetPages());
        var actualImages = actualPage.GetImages().OrderBy(image => image.BoundingBox.Left).ToArray();
        var expectedImages = expectedPage.GetImages().OrderBy(image => image.BoundingBox.Left).ToArray();
        Assert.Equal(pictureCount, expectedImages.Length);
        Assert.Equal(pictureCount, actualImages.Length);
        for (int index = 0; index < pictureCount; index++) {
            var image = actualImages[index].BoundingBox;
            var reference = expectedImages[index].BoundingBox;
            Assert.Equal(reference.Width, image.Width, 3);
            Assert.Equal(reference.Height, image.Height, 3);
            Assert.InRange(Math.Abs(reference.Left - image.Left), 0D, .05D);
            Assert.InRange(Math.Abs(reference.Bottom - image.Bottom), 0D, .05D);
        }
        var following = Assert.Single(actualPage.GetWords(), word => word.Text == "FOLLOW");
        var referenceFollowing = Assert.Single(expectedPage.GetWords(), word => word.Text == "FOLLOW");
        Assert.InRange(Math.Abs(following.Letters[0].StartBaseLine.X - referenceFollowing.Letters[0].StartBaseLine.X), 0D, .05D);
        Assert.InRange(Math.Abs(following.Letters[0].StartBaseLine.Y - referenceFollowing.Letters[0].StartBaseLine.Y), 0D, .05D);
    }

    private static void AssertHeadingPicturePaintMatches(byte[] actualBytes, byte[] expectedBytes) {
        PdfCore.PdfPageRenderResult actual = Assert.Single(PdfCore.PdfDocument.Load(actualBytes).Render.Pages(
            options: new PdfCore.PdfPageRenderOptions { Dpi = 144 }));
        PdfCore.PdfPageRenderResult expected = Assert.Single(PdfCore.PdfDocument.Load(expectedBytes).Render.Pages(
            options: new PdfCore.PdfPageRenderOptions { Dpi = 144 }));
        Assert.True(actual.Succeeded, string.Join("\n", actual.Diagnostics));
        Assert.True(expected.Succeeded, string.Join("\n", expected.Diagnostics));
        Assert.True(OfficeRasterImageDecoder.TryDecode(actual.Bytes!, out OfficeRasterImage? actualImage));
        Assert.True(OfficeRasterImageDecoder.TryDecode(expected.Bytes!, out OfficeRasterImage? expectedImage));
        Assert.Equal(expectedImage!.Width, actualImage!.Width);
        Assert.Equal(expectedImage.Height, actualImage.Height);
        for (int y = 0; y < expectedImage.Height; y++) for (int x = 0; x < expectedImage.Width; x++) {
            OfficeColor reference = expectedImage.GetPixel(x, y);
            OfficeColor actualPixel = actualImage.GetPixel(x, y);
            if (!reference.Equals(actualPixel))
                Assert.Fail($"Paint differs at ({x}, {y}): expected {reference}, actual {actualPixel}.");
        }
    }
}
