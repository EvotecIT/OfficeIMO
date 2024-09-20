using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using VerifyXunit;
using Xunit;

using Color = SixLabors.ImageSharp.Color;

namespace OfficeIMO.VerifyTests.Word;

public class BasicDocumentTests : VerifyTestBase {

    private static async Task DoTest(WordprocessingDocument document) {
        var result = await ToVerifyResult(document);
        await Verifier.Verify(result, GetSettings());
    }

    [Fact]
    public async Task BasicEmptyWord() {
        using var document = WordDocument.Create();
        document.BuiltinDocumentProperties.Title = "This is my title";
        document.BuiltinDocumentProperties.Creator = "Przemysław Kłys";
        document.BuiltinDocumentProperties.Keywords = "word, docx, test";
        document.Save();

        await DoTest(document._wordprocessingDocument);
    }

    [Fact]
    public async Task BasicWord() {
        using var document = WordDocument.Create();
        var paragraph = document.AddParagraph("Adding paragraph with some text");
        paragraph.ParagraphAlignment = JustificationValues.Center;

        paragraph.Color = Color.Red;

        paragraph = document.AddParagraph("Adding another paragraph with some more text");
        paragraph.Bold = true;
        paragraph = paragraph.AddText(" , but now we also decided to add more text to this paragraph using different style");
        paragraph.Underline = UnderlineValues.DashLong;
        paragraph = paragraph.AddText(" , and we still continue adding more text to existing paragraph.");
        paragraph.Color = Color.CornflowerBlue;

        document.Save();

        await DoTest(document._wordprocessingDocument);
    }

    [Fact]
    public async Task BasicWordWithBreaks() {
        using var document = WordDocument.Create();
        _ = document.AddParagraph("Adding paragraph1 with some text and pressing ENTER");

        var paragraph = document.AddParagraph("Adding paragraph2 with some text and pressing SHIFT+ENTER");
        paragraph.AddBreak();
        paragraph.AddText("Continue1");
        paragraph.AddBreak();
        paragraph.AddText("Continue2");
        paragraph.AddText(" Continue3");

        document.Breaks[0].Remove(); // removes break before continue1

        _ = document.AddParagraph("Adding paragraph3 with some text and pressing ENTER");
        _ = document.AddParagraph("Adding paragraph4 with some text and pressing SHIFT+ENTER");

        document.Save();

        await DoTest(document._wordprocessingDocument);
    }

    [Fact]
    public async Task BasicWordWithDefaultStyleChange() {
        using var document = WordDocument.Create();
        document.Settings.FontSize = 30;
        document.Settings.FontFamily = "Calibri Light";
        document.Settings.Language = "pl-PL";
        document.Settings.Language = "pt-Br";
        _ = document.AddParagraph("To jest po polsku");

        var paragraph = document.AddParagraph("Adding paragraph1 with some text and pressing ENTER");
        paragraph.FontSize = 15;
        paragraph.FontFamily = "Courier New";

        document.Save();

        await DoTest(document._wordprocessingDocument);
    }

    [Fact]
    public async Task BasicWordWithDefaultFontChange() {
        using var document = WordDocument.Create();
        document.Settings.FontSize = 30;
        document.Settings.FontFamily = "Calibri Light";
        document.Settings.FontFamilyHighAnsi = "Calibri Light";
        document.Settings.Language = "pt-Br";
        document.Settings.ZoomPreset = PresetZoomValues.BestFit;

        document.CompatibilitySettings.CompatibilityMode = CompatibilityMode.Word2013;
        document.CompatibilitySettings.CompatibilityMode = CompatibilityMode.None;

        const string title = "INSTRUMENTO PARTICULAR DE CONSTITUIÇÃO DE GARANTIA DE ALIENAÇÃO FIDUCIÁRIA DE IMÓVEL";

        document.AddParagraph(title).SetBold().ParagraphAlignment = JustificationValues.Center;

        document.Save();

        await DoTest(document._wordprocessingDocument);
    }
}
