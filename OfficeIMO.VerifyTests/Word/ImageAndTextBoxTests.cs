using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Word;
using System;
using System.IO;
using System.Threading.Tasks;
using VerifyXunit;
using Xunit;

namespace OfficeIMO.VerifyTests.Word;

public class ImageAndTextBoxTests : VerifyTestBase {
    private static async Task DoTest(WordDocument document) {
        document.Save();

        var result = await ToVerifyResult(document._wordprocessingDocument);
        await Verifier.Verify(result, GetSettings());
    }

    private static string GetSampleImagePath() {
        return Path.GetFullPath(Path.Combine(
            AppContext.BaseDirectory,
            "..",
            "..",
            "..",
            "..",
            "OfficeIMO.Tests",
            "Images",
            "Kulek.jpg"));
    }

    [Fact]
    public async Task ImageDocument() {
        using var document = WordDocument.Create();
        document.AddParagraph("Image");
        document.AddParagraph().AddImage(GetSampleImagePath(), 50, 50);

        await DoTest(document);
    }

    [Fact]
    public async Task WrappedImageDocument() {
        using var document = WordDocument.Create();
        var paragraph = document.AddParagraph("Wrapped image");
        var image = paragraph.InsertImage(GetSampleImagePath(), 90, 45, WrapTextImage.Square, "Wrapped sample image");
        image.Title = "Wrapped title";
        image.horizontalPosition = new DocumentFormat.OpenXml.Drawing.Wordprocessing.HorizontalPosition() {
            RelativeFrom = DocumentFormat.OpenXml.Drawing.Wordprocessing.HorizontalRelativePositionValues.Page,
            PositionOffset = new DocumentFormat.OpenXml.Drawing.Wordprocessing.PositionOffset() { Text = "457200" }
        };
        image.verticalPosition = new DocumentFormat.OpenXml.Drawing.Wordprocessing.VerticalPosition() {
            RelativeFrom = DocumentFormat.OpenXml.Drawing.Wordprocessing.VerticalRelativePositionValues.Paragraph,
            PositionOffset = new DocumentFormat.OpenXml.Drawing.Wordprocessing.PositionOffset() { Text = "228600" }
        };

        await DoTest(document);
    }

    [Fact]
    public async Task HeaderImageDocument() {
        using var document = WordDocument.Create();
        document.AddHeadersAndFooters();

        var defaultHeader = document.Sections[0].GetOrCreateHeader(DocumentFormat.OpenXml.Wordprocessing.HeaderFooterValues.Default);
        var headerParagraph = defaultHeader.AddParagraph("Header image");
        var image = headerParagraph.InsertImage(GetSampleImagePath(), 64, 64, WrapTextImage.Square, "Header image");
        image.Title = "Header logo";

        await DoTest(document);
    }

    [Fact]
    public async Task TextBoxDocument() {
        using var document = WordDocument.Create();
        document.AddParagraph("Text box");
        var textBox = document.AddTextBox("Hello from textbox", WrapTextImage.Through);
        textBox.AutoFit = WordTextBoxAutoFitType.ShrinkTextOnOverflow;
        textBox.HorizontalPositionRelativeFrom = DocumentFormat.OpenXml.Drawing.Wordprocessing.HorizontalRelativePositionValues.Page;
        textBox.HorizontalPositionOffsetCentimeters = 2;
        textBox.VerticalPositionOffsetCentimeters = 3;

        await DoTest(document);
    }

    [Fact]
    public async Task HeaderTextBoxDocument() {
        using var document = WordDocument.Create();
        document.AddHeadersAndFooters();

        var defaultHeader = document.Sections[0].GetOrCreateHeader(DocumentFormat.OpenXml.Wordprocessing.HeaderFooterValues.Default);
        var textBox = defaultHeader.AddTextBox("Header textbox", WrapTextImage.Square);
        textBox.AutoFit = WordTextBoxAutoFitType.ResizeShapeToFitText;
        textBox.HorizontalAlignment = WordHorizontalAlignmentValues.Right;
        textBox.VerticalPositionOffsetCentimeters = 1.5;
        textBox.Paragraphs[0].AddHyperLink(" link", new Uri("https://officeimo.example/header"), addStyle: true);

        await DoTest(document);
    }
}
