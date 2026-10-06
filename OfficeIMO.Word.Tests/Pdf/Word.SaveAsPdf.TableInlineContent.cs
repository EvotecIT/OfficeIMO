using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData("body")]
    [InlineData("header")]
    public void SaveAsPdf_MixedPictureRunPreservesItsVisibleText(string frame) {
        using WordDocument document = WordDocument.Create();
        WordParagraph paragraph;
        if (frame == "header") {
            document.AddHeadersAndFooters();
            paragraph = RequireSectionHeader(document, 0, DocumentFormat.OpenXml.Wordprocessing.HeaderFooterValues.Default)
                .AddParagraph("Before");
            document.AddParagraph("Body");
        } else paragraph = document.AddParagraph("Before");
        using var image = new MemoryStream(OfficeIMO.Tests.Pdf.PdfPngTestImages.CreateRgbPng(2, 1));
        paragraph.AddImage(image, "inline.png", 16, 32);
        paragraph.AddText("After");
        using PdfPigDocument pdf = PdfPigDocument.Open(document.ToPdfBytes(new WordToPdfOptions { IncludePageNumbers = false }));
        string text = string.Concat(pdf.GetPages().Select(page => page.Text));
        Assert.Contains("Before", text);
        Assert.Contains("After", text);
        Assert.Single(pdf.GetPages().SelectMany(page => page.GetImages()));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SaveAsPdf_TableInlinePicturePreservesTextAndImageOrder(bool mixedRun) {
        string sourcePath = Path.Combine(_directoryWithFiles, $"PdfTableInlineOrder-{mixedRun}.docx");
        using WordDocument document = WordDocument.Create(sourcePath);
        WordParagraph paragraph = document.AddTable(1, 1, WordTableStyle.TableGrid).Rows[0].Cells[0].Paragraphs[0];
        paragraph.Text = "Before";
        paragraph.FontFamily = "Arial"; paragraph.FontSize = 12;
        paragraph.LineSpacingBeforePoints = 0; paragraph.LineSpacingAfterPoints = 0;
        paragraph.LineSpacingRule = WordLineSpacingRule.Auto; paragraph.LineSpacing = 240;
        using var image = new MemoryStream(OfficeIMO.Tests.Pdf.PdfPngTestImages.CreateRgbPng(2, 1));
        (mixedRun ? paragraph : paragraph.AddText(string.Empty)).AddImage(image, "inline.png", 16, 32, description: "Inline marker");
        paragraph.AddText("After");
        document.Save();
        string originalXml = document._wordprocessingDocument.MainDocumentPart!.Document.OuterXml;
        using PdfPigDocument pdf = PdfPigDocument.Open(document.ToPdfBytes(new WordToPdfOptions {
            IncludePageNumbers = false, ResourcePolicy = OfficeIMO.Pdf.PdfResourcePolicy.CreatePortableDeterministic()
        }));
        Assert.Equal(originalXml, document._wordprocessingDocument.MainDocumentPart.Document.OuterXml);
        var page = pdf.GetPage(1);
        var before = Assert.Single(page.GetWords(), word => word.Text == "Before");
        var after = Assert.Single(page.GetWords(), word => word.Text == "After");
        var picture = Assert.Single(page.GetImages());
        Assert.Equal(12D, picture.BoundingBox.Width, 3);
        Assert.Equal(24D, picture.BoundingBox.Height, 3);
        Assert.True(before.BoundingBox.Right <= picture.BoundingBox.Left + .5D);
        Assert.True(picture.BoundingBox.Right <= after.BoundingBox.Left + .5D);
        Assert.Equal(before.BoundingBox.Bottom, after.BoundingBox.Bottom, 3);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SaveAsPdf_TableInlinePictureWithoutTextSurvivesNestedProjection(bool nested) {
        using WordDocument document = WordDocument.Create();
        WordTableCell cell = document.AddTable(1, 1).Rows[0].Cells[0];
        if (nested) cell = cell.AddTable(1, 1).Rows[0].Cells[0];
        using var image = new MemoryStream(OfficeIMO.Tests.Pdf.PdfPngTestImages.CreateRgbPng(2, 1));
        cell.Paragraphs[0].AddImage(image, "inline.png", 16, 32);
        using PdfPigDocument pdf = PdfPigDocument.Open(document.ToPdfBytes(new WordToPdfOptions { IncludePageNumbers = false }));
        var picture = Assert.Single(pdf.GetPages().SelectMany(page => page.GetImages()));
        Assert.Equal(12D, picture.BoundingBox.Width, 3);
        Assert.Equal(24D, picture.BoundingBox.Height, 3);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SaveAsPdf_TableMixedPictureRunUsesVisibleFieldResult(bool hidden) {
        using WordDocument document = WordDocument.Create();
        WordParagraph paragraph = document.AddTable(1, 1).Rows[0].Cells[0].Paragraphs[0];
        paragraph.Text = "Before";
        using var image = new MemoryStream(OfficeIMO.Tests.Pdf.PdfPngTestImages.CreateRgbPng(2, 1));
        paragraph.AddImage(image, "inline.png", 16, 32);
        W.Run resultRun = paragraph._run!;
        W.Drawing drawing = resultRun.GetFirstChild<W.Drawing>()!;
        var instructionDrawing = (W.Drawing)drawing.CloneNode(true);
        resultRun.Append(new W.Text("After"));
        if (hidden) {
            resultRun.RunProperties ??= new W.RunProperties();
            resultRun.RunProperties.Vanish = new W.Vanish();
        }
        resultRun.Remove();
        paragraph._paragraph!.Append(
            new W.Run(new W.FieldChar { FieldCharType = W.FieldCharValues.Begin }),
            new W.Run(new W.FieldCode(" PRIVATE instruction "), instructionDrawing),
            new W.Run(new W.FieldChar { FieldCharType = W.FieldCharValues.Separate }),
            resultRun,
            new W.Run(new W.FieldChar { FieldCharType = W.FieldCharValues.End }));
        using PdfPigDocument pdf = PdfPigDocument.Open(document.ToPdfBytes(new WordToPdfOptions {
            IncludePageNumbers = false, MaxImagesPerParagraph = 1
        }));
        var pictures = pdf.GetPages().SelectMany(page => page.GetImages()).ToArray();
        string text = string.Concat(pdf.GetPages().Select(page => page.Text));
        Assert.DoesNotContain("instruction", text);
        if (hidden) {
            Assert.Empty(pictures);
            Assert.DoesNotContain("Before", text);
            Assert.DoesNotContain("After", text);
        } else {
            Assert.Single(pictures);
            Assert.Contains("Before", text);
            Assert.Contains("After", text);
        }
    }

    [Fact]
    public void SaveAsPdf_TableMixedRunPreservesMultiplePicturesAndTheirLimit() {
        using WordDocument document = WordDocument.Create();
        WordParagraph paragraph = document.AddTable(1, 1).Rows[0].Cells[0].Paragraphs[0];
        paragraph.Text = "Before";
        paragraph.FontFamily = "Arial";
        paragraph.FontSize = 12;
        paragraph.Bold = true;
        using var image = new MemoryStream(OfficeIMO.Tests.Pdf.PdfPngTestImages.CreateRgbPng(2, 1));
        paragraph.AddImage(image, "inline.png", 16, 32);
        W.Drawing second = (W.Drawing)paragraph._run!.GetFirstChild<W.Drawing>()!.CloneNode(true);
        second.Inline!.DocProperties!.Id = 2;
        paragraph._run.Append(new W.Text("Middle"), second, new W.Text("After"));
        Assert.Throws<InvalidDataException>(() => document.ToPdfDocumentResult(new WordToPdfOptions { MaxImagesPerParagraph = 1 }));
        using PdfPigDocument pdf = PdfPigDocument.Open(document.ToPdfBytes(new WordToPdfOptions {
            IncludePageNumbers = false, MaxImagesPerParagraph = 2
        }));
        var page = pdf.GetPage(1);
        var pictures = page.GetImages().OrderBy(picture => picture.BoundingBox.Left).ToArray();
        Assert.Equal(2, pictures.Length);
        var before = Assert.Single(page.GetWords(), word => word.Text == "Before");
        var middle = Assert.Single(page.GetWords(), word => word.Text == "Middle");
        var after = Assert.Single(page.GetWords(), word => word.Text == "After");
        Assert.True(before.BoundingBox.Right <= pictures[0].BoundingBox.Left + .5D);
        Assert.True(pictures[0].BoundingBox.Right <= middle.BoundingBox.Left + .5D);
        Assert.True(middle.BoundingBox.Right <= pictures[1].BoundingBox.Left + .5D);
        // Word can wrap at an inline image boundary without enlarging the
        // authored automatic grid to the whole text-and-picture sequence.
        Assert.True(after.BoundingBox.Top < middle.BoundingBox.Bottom);
        Assert.Contains("Bold", before.Letters[0].FontName);
        Assert.Contains("Bold", middle.Letters[0].FontName);
        Assert.Contains("Bold", after.Letters[0].FontName);
    }

    [Theory]
    [InlineData("body", false)]
    [InlineData("header", false)]
    [InlineData("table", false)]
    [InlineData("body", true)]
    [InlineData("header", true)]
    [InlineData("table", true)]
    public void SaveAsPdf_HiddenMixedPictureRunDoesNotExposeItsContents(string frame, bool contentControl) {
        using WordDocument document = WordDocument.Create();
        WordParagraph paragraph;
        if (frame == "header") {
            document.AddHeadersAndFooters();
            paragraph = RequireSectionHeader(document, 0, W.HeaderFooterValues.Default).AddParagraph("HiddenBefore");
        } else if (frame == "table") {
            paragraph = document.AddTable(1, 1).Rows[0].Cells[0].Paragraphs[0];
            paragraph.Text = "HiddenBefore";
        } else paragraph = document.AddParagraph("HiddenBefore");
        byte[] bytes = OfficeIMO.Tests.Pdf.PdfPngTestImages.CreateRgbPng(2, 1);
        if (contentControl) {
            string imagePath = Path.Combine(_directoryWithFiles, "hidden.png");
            File.WriteAllBytes(imagePath, bytes);
            paragraph.AddPictureControl(imagePath, 16, 32);
            var pictureRun = paragraph._paragraph!.Descendants<W.SdtRun>().Single().SdtContentRun!.Elements<W.Run>().Single();
            pictureRun.RunProperties ??= new W.RunProperties();
            pictureRun.RunProperties.Vanish = new W.Vanish();
        } else {
            using var image = new MemoryStream(bytes);
            paragraph.AddImage(image, "hidden.png", 16, 32);
            paragraph._run!.Append(new W.Text("HiddenAfter"));
        }
        paragraph._run!.RunProperties ??= new W.RunProperties();
        paragraph._run.RunProperties.Vanish = new W.Vanish();
        document.AddParagraph("VisibleBody");
        using PdfPigDocument pdf = PdfPigDocument.Open(document.ToPdfBytes(new WordToPdfOptions { IncludePageNumbers = false }));
        Assert.Empty(pdf.GetPages().SelectMany(page => page.GetImages()));
        string text = string.Concat(pdf.GetPages().Select(page => page.Text));
        Assert.DoesNotContain("HiddenBefore", text);
        Assert.DoesNotContain("HiddenAfter", text);
        Assert.Contains("VisibleBody", text);
    }
}
