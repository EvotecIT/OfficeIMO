using OfficeIMO.Rtf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using OfficeIMO.Word.Rtf;
using W = DocumentFormat.OpenXml.Wordprocessing;
using PdfCore = OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public sealed class RtfWordDefaultFontTests {
    [Theory]
    [InlineData(false, "Courier New", "Courier", 1)]
    [InlineData(true, "Courier New", "Courier", 1)]
    [InlineData(false, "Calibri", "Helvetica", 0)]
    public void SelectedDefaultFontSurvivesWordSaveReloadWithoutDirectFontIds(bool parsed,
        string family, string paintedFont, int expectedFontId) {
        RtfDocument authored = CreateSource(out int selectedFont, family);
        authored.AddParagraph().AddText("DEFAULTFONT").SetFontSize(12);
        RtfDocument source = parsed ? RtfDocument.Read(authored.ToRtf()).Document : authored;
        Assert.Equal(expectedFontId, selectedFont);
        Assert.Equal(selectedFont, source.Settings.DefaultFontId);
        Assert.All(source.Paragraphs.SelectMany(paragraph => paragraph.Runs), run => Assert.Null(run.FontId));
        string original = source.ToRtf();

        using WordDocument converted = source.ToWordDocument();
        using WordDocument reopened = WordDocument.Load(new MemoryStream(converted.ToBytes()));
        using PdfPigDocument pdf = PdfPigDocument.Open((reopened.ToPdf(PortableOptions())).ToBytes());

        AssertPaintedFont(pdf, "DEFAULTFONT", paintedFont);
        AssertDefaultFonts(reopened, family);
        Assert.All(reopened.OpenXmlDocument.MainDocumentPart!.Document!.Descendants<W.Run>(),
            run => Assert.Null(run.RunProperties?.RunFonts));
        Assert.Equal(original, source.ToRtf());
    }

    [Fact]
    public void DocumentDefaultFontReachesUnformattedBodyCellAndNoteRuns() {
        RtfDocument source = CreateSource(out _);
        source.AddParagraph().AddText("BODYFONT").SetFontSize(12);
        source.AddTable(1, 1).Rows[0].Cells[0].AddParagraph().AddText("CELLFONT").SetFontSize(12);
        RtfNote note = source.AddNote(RtfNoteKind.Footnote);
        note.AddParagraph().AddText("NOTEFONT").SetFontSize(12);
        source.AddParagraph("Anchor").AddNoteReference(note);
        string original = source.ToRtf();

        using WordDocument converted = source.ToWordDocument();
        using WordDocument reopened = WordDocument.Load(new MemoryStream(converted.ToBytes()));
        using PdfPigDocument pdf = PdfPigDocument.Open((reopened.ToPdf(PortableOptions())).ToBytes());

        AssertPaintedFont(pdf, "BODYFONT", "Courier");
        AssertPaintedFont(pdf, "CELLFONT", "Courier");
        AssertPaintedFont(pdf, "NOTEFONT", "Courier");
        AssertDefaultFonts(reopened, "Courier New");
        var main = reopened.OpenXmlDocument.MainDocumentPart!;
        Assert.All(main.Document!.Descendants<W.Run>(), run => Assert.Null(run.RunProperties?.RunFonts));
        Assert.All(main.FootnotesPart!.Footnotes!.Descendants<W.Run>(), run => Assert.Null(run.RunProperties?.RunFonts));
        Assert.Equal(original, source.ToRtf());
    }

    [Fact]
    public void ExplicitRunAndStyleFontsOverrideTheSelectedDocumentDefault() {
        RtfDocument source = CreateSource(out _);
        int styleFont = source.AddFont("Times New Roman");
        int directFont = source.AddFont("Arial");
        source.AddStyle(7, "Explicit serif").FontId = styleFont;
        source.AddParagraph().SetStyle(7).AddText("STYLEFONT").SetFontSize(12);
        RtfRun run = source.AddParagraph().SetStyle(7).AddText("DIRECTFONT").SetFontSize(12);
        run.FontId = directFont;
        string original = source.ToRtf();

        using WordDocument converted = source.ToWordDocument();
        using WordDocument reopened = WordDocument.Load(new MemoryStream(converted.ToBytes()));
        using PdfPigDocument pdf = PdfPigDocument.Open((reopened.ToPdf(PortableOptions())).ToBytes());

        AssertPaintedFont(pdf, "STYLEFONT", "Times");
        AssertPaintedFont(pdf, "DIRECTFONT", "Helvetica");
        Assert.Equal(original, source.ToRtf());
    }

    private static RtfDocument CreateSource(out int selectedFont, string family = "Courier New") {
        RtfDocument source = RtfDocument.Create();
        selectedFont = source.AddFont(family);
        source.Settings.SetDefaultFont(selectedFont);
        return source;
    }

    private static void AssertPaintedFont(PdfPigDocument pdf, string marker, string fontName) {
        var word = Assert.Single(pdf.GetPages().SelectMany(page => page.GetWords()), word => word.Text == marker);
        Assert.All(word.Letters, letter => Assert.Contains(fontName, letter.FontName, StringComparison.OrdinalIgnoreCase));
    }

    private static void AssertDefaultFonts(WordDocument document, string expected) {
        W.RunFonts fonts = Assert.IsType<W.RunFonts>(document.OpenXmlDocument.MainDocumentPart!
            .StyleDefinitionsPart!.Styles!.DocDefaults!.RunPropertiesDefault!.RunPropertiesBaseStyle!.RunFonts);
        Assert.Equal(expected, fonts.Ascii?.Value);
        Assert.Equal(expected, fonts.HighAnsi?.Value);
        Assert.Equal(expected, fonts.EastAsia?.Value);
        Assert.Equal(expected, fonts.ComplexScript?.Value);
        Assert.Null(fonts.AsciiTheme);
        Assert.Null(fonts.HighAnsiTheme);
        Assert.Null(fonts.EastAsiaTheme);
        Assert.Null(fonts.ComplexScriptTheme);
    }

    private static WordPdfExportOptions PortableOptions() => new() {
        ResourcePolicy = PdfCore.PdfResourcePolicy.CreatePortableDeterministic(),
        TextFallbacks = PdfCore.PdfTextFallbackFeatures.None
    };
}
