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
        _ = document.ToBytes();

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
            "OfficeIMO.TestAssets",
            "Images",
            "Kulek.jpg"));
    }

    [Fact]
    public async Task ImageDocument() {
        using var document = WordDocument.Create();
        document.AddParagraph("Image");
        document.AddParagraph().AddImage(GetSampleImagePath(), OfficeIMO.OfficeLength.Pixels(50), OfficeIMO.OfficeLength.Pixels(50));

        await DoTest(document);
    }

    [Fact]
    public async Task WrappedImageDocument() {
        using var document = WordDocument.Create();
        var paragraph = document.AddParagraph("Wrapped image");
        var image = paragraph.AddImage(GetSampleImagePath(), OfficeIMO.OfficeLength.Pixels(90), OfficeIMO.OfficeLength.Pixels(45), options: new OfficeIMO.Word.WordImageOptions { Wrapping = WordImageTextWrapping.Square, Description = "Wrapped sample image" });
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

        var defaultHeader = document.Sections[0].GetOrCreateHeader(WordHeaderFooterType.Default);
        var headerParagraph = defaultHeader.AddParagraph("Header image");
        var image = headerParagraph.AddImage(GetSampleImagePath(), OfficeIMO.OfficeLength.Pixels(64), OfficeIMO.OfficeLength.Pixels(64), options: new OfficeIMO.Word.WordImageOptions { Wrapping = WordImageTextWrapping.Square, Description = "Header image" });
        image.Title = "Header logo";

        await DoTest(document);
    }

    [Fact]
    public async Task TextBoxDocument() {
        using var document = WordDocument.Create();
        document.AddParagraph("Text box");
        var textBox = document.AddTextBox("Hello from textbox", WordImageTextWrapping.Through);
        textBox.AutoFit = WordTextBoxAutoFitType.ShrinkTextOnOverflow;
        textBox.HorizontalPositionRelativeFrom = WordHorizontalRelativePosition.Page;
        textBox.HorizontalPositionOffsetCentimeters = 2;
        textBox.VerticalPositionOffsetCentimeters = 3;

        await DoTest(document);
    }

    [Fact]
    public async Task HeaderTextBoxDocument() {
        using var document = WordDocument.Create();
        document.AddHeadersAndFooters();

        var defaultHeader = document.Sections[0].GetOrCreateHeader(WordHeaderFooterType.Default);
        var textBox = defaultHeader.AddTextBox("Header textbox", WordImageTextWrapping.Square);
        textBox.AutoFit = WordTextBoxAutoFitType.ResizeShapeToFitText;
        textBox.HorizontalAlignment = WordTextBoxHorizontalAlignment.Right;
        textBox.VerticalPositionOffsetCentimeters = 1.5;
        textBox.Paragraphs[0].AddHyperLink(" link", new Uri("https://officeimo.example/header"), addStyle: true);

        await DoTest(document);
    }
}
