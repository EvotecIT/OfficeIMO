using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Word;
using System;
using System.IO;
using System.Threading.Tasks;
using VerifyXunit;
using Xunit;

namespace OfficeIMO.VerifyTests.Word;

public class ImageAndTextBoxTests : VerifyTestBase {
    private static async Task DoTest(WordprocessingDocument document) {
        var result = await ToVerifyResult(document);
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
        document.Save();

        await DoTest(document._wordprocessingDocument);
    }

    [Fact]
    public async Task TextBoxDocument() {
        using var document = WordDocument.Create();
        document.AddParagraph("Text box");
        document.AddTextBox("Hello from textbox");
        document.Save();

        await DoTest(document._wordprocessingDocument);
    }
}
