using System;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using VerifyXunit;
using Xunit;

using Color = SixLabors.ImageSharp.Color;

namespace OfficeIMO.VerifyTests.Word;

public class CustomAndBuiltinPropertiesTests : VerifyTestBase {

    private static async Task DoTest(WordprocessingDocument document) {
        var result = await ToVerifyResult(document);
        await Verifier.Verify(result, GetSettings());
    }

    [Fact]
    public async Task ValidateDocument() {
        using var document = WordDocument.Create();
        var paragraph = document.AddParagraph("Basic paragraph - Page 4");
        paragraph.ParagraphAlignment = JustificationValues.Center;

        document.CustomDocumentProperties.Add("TestProperty", new WordCustomProperty { Value = DateTime.Today });
        document.CustomDocumentProperties.Add("MyName", new WordCustomProperty("Some text"));
        document.CustomDocumentProperties.Add("IsTodayGreatDay", new WordCustomProperty(true));

        document.Save();

        await DoTest(document._wordprocessingDocument);
    }

    [Fact]
    public async Task BasicDocumentProperties() {
        using var document = WordDocument.Create();
        document.BuiltinDocumentProperties.Title = "This is my title";
        document.BuiltinDocumentProperties.Creator = "Przemysław Kłys";
        document.BuiltinDocumentProperties.Keywords = "word, docx, test";

        var paragraph = document.AddParagraph("Basic paragraph");
        paragraph.ParagraphAlignment = JustificationValues.Center;
        paragraph.Color = Color.Red;

        document.Save();

        await DoTest(document._wordprocessingDocument);
    }
}
