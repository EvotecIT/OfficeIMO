using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class PdfShortNumberingTabTests {
    [Theory]
    [InlineData(WordCompatibilityMode.Word2003, false, false, false, true, 18)]
    [InlineData(WordCompatibilityMode.Word2003, true, false, false, true, 18)]
    [InlineData(WordCompatibilityMode.Word2013, true, false, false, true, 18)]
    [InlineData(WordCompatibilityMode.Word2003, false, false, true, true, 18)]
    [InlineData(WordCompatibilityMode.Word2003, true, false, true, true, 18)]
    [InlineData(WordCompatibilityMode.Word2013, true, false, true, true, 18)]
    [InlineData(WordCompatibilityMode.Word2003, false, true, false, true, 18)]
    [InlineData(WordCompatibilityMode.Word2003, true, true, false, true, 18)]
    [InlineData(WordCompatibilityMode.Word2013, true, true, false, true, 18)]
    [InlineData(WordCompatibilityMode.Word2003, true, false, false, false, 18)]
    [InlineData(WordCompatibilityMode.Word2003, false, false, false, false, 54)]
    [InlineData(WordCompatibilityMode.Word2013, true, false, false, false, 54)]
    public void ShortNumberingTabsKeepFirstLineIndependentOfContinuationIndent(
        WordCompatibilityMode mode, bool ignoreIndent, bool table, bool inline, bool authored, double firstLineAdvance) {
        using WordDocument document = WordDocument.Create();
        document.CompatibilitySettings.CompatibilityMode = mode;
        document.CompatibilitySettings.DoNotUseIndentAsNumberingTabStop = ignoreIndent;
        WordList list = document.AddCustomList();
        ConfigureShortNumbering(list, authored);
        WordParagraph paragraph = table ? document.AddTable(1, 1).Rows[0].Cells[0].Paragraphs[0] : list.AddItem("BODY\nNEXT");
        paragraph.Text = "BODY\nNEXT";
        paragraph.FontFamily = "Courier New"; paragraph.FontSize = 12;
        if (inline) paragraph.CharacterScale = 95;
        if (table) paragraph._paragraph.ParagraphProperties = new ParagraphProperties(new NumberingProperties(
            new NumberingLevelReference { Val = 0 }, new NumberingId { Val = list.NumberId }));
        using var pdf = OpenPdf(document);
        var letters = pdf.GetPage(1).Letters.Where(letter => !string.IsNullOrWhiteSpace(letter.Value)).ToArray();
        Assert.Equal("1.BODYNEXT", string.Concat(letters.Select(letter => letter.Value)));
        Assert.InRange(Math.Abs(letters[2].StartBaseLine.X - letters[0].StartBaseLine.X - firstLineAdvance), 0D, 0.03D);
        Assert.InRange(Math.Abs(letters[6].StartBaseLine.X - letters[0].StartBaseLine.X - 54D), 0D, 0.03D);
    }

    [Theory]
    [InlineData(false, WordCompatibilityMode.Word2003, true, 18)]
    [InlineData(true, WordCompatibilityMode.Word2003, true, 18)]
    [InlineData(false, WordCompatibilityMode.Word2003, false, 18)]
    [InlineData(true, WordCompatibilityMode.Word2003, false, 18)]
    [InlineData(false, WordCompatibilityMode.Word2013, true, 18)]
    [InlineData(true, WordCompatibilityMode.Word2013, true, 18)]
    [InlineData(false, WordCompatibilityMode.Word2013, false, 54)]
    [InlineData(true, WordCompatibilityMode.Word2013, false, 54)]
    public void PlainRunningListsUseParagraphNumberingTabsAndMarkerTypography(
        bool footer, WordCompatibilityMode mode, bool authored, double firstLineAdvance) {
        using WordDocument document = WordDocument.Create();
        document.CompatibilitySettings.CompatibilityMode = mode;
        document.CompatibilitySettings.DoNotUseIndentAsNumberingTabStop = true;
        WordHeaderFooter story = footer ? document.FooterDefaultOrCreate : document.HeaderDefaultOrCreate;
        WordList list = story.AddList(WordListStyle.Custom);
        ConfigureShortNumbering(list, authored);
        WordParagraph paragraph = list.AddItem("RUNNING");
        paragraph.FontFamily = "Courier New"; paragraph.FontSize = 12;
        document.AddParagraph("BODY");
        Assert.Empty(document.ValidateDocument());
        using var pdf = OpenPdf(document);
        var letters = pdf.GetPage(1).Letters.Where(letter => !string.IsNullOrWhiteSpace(letter.Value)).ToArray();
        var marker = letters.First(letter => letter.Value == "1");
        var text = letters.First(letter => letter.Value == "R");
        Assert.Contains("RUNNING", pdf.GetPage(1).Text);
        Assert.InRange(Math.Abs(text.StartBaseLine.X - marker.StartBaseLine.X - firstLineAdvance), 0D, 0.03D);
    }

    private static void ConfigureShortNumbering(WordList list, bool authored) {
        list.Numbering.AddLevel(new WordListLevel(WordListLevelKind.DecimalDot));
        WordListLevel level = list.Numbering.Levels[0];
        level.IndentationLeft = 1440; level.IndentationHanging = 1080;
        if (authored) level.OpenXmlElement.PreviousParagraphProperties!.InsertAt(new Tabs(
            new TabStop { Val = TabStopValues.Number, Position = 720 }), 0);
        level.OpenXmlElement.NumberingSymbolRunProperties = new NumberingSymbolRunProperties(
            new RunFonts { Ascii = "Courier New", HighAnsi = "Courier New" }, new FontSize { Val = "24" });
    }

    private static PdfPigDocument OpenPdf(WordDocument document) => PdfPigDocument.Open(document.ToPdfBytes(new WordToPdfOptions {
        IncludePageNumbers = false, FontFamily = "Courier"
    }));
}
