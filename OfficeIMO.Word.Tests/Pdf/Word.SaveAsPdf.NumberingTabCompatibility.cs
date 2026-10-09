using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class PdfNumberingTabCompatibilityTests {
    [Theory]
    [InlineData(WordCompatibilityMode.Word2003, false, false, false, 18)]
    [InlineData(WordCompatibilityMode.Word2003, true, false, false, 132)]
    [InlineData(WordCompatibilityMode.Word2007, true, false, false, 132)]
    [InlineData(WordCompatibilityMode.Word2010, true, false, false, 132)]
    [InlineData(WordCompatibilityMode.Word2013, true, false, false, 18)]
    [InlineData(WordCompatibilityMode.Word2003, true, false, true, 132)]
    [InlineData(WordCompatibilityMode.Word2003, false, true, false, 18)]
    [InlineData(WordCompatibilityMode.Word2003, true, true, false, 132)]
    [InlineData(WordCompatibilityMode.Word2013, true, true, false, 18)]
    public void NumberingTabsRespectLegacyCompatibilityWithoutMovingContinuationIndent(
        WordCompatibilityMode mode, bool ignoreIndent, bool table, bool inline, double bodyOffset) {
        using WordDocument document = WordDocument.Create();
        document.CompatibilitySettings.CompatibilityMode = mode;
        document.CompatibilitySettings.DoNotUseIndentAsNumberingTabStop = ignoreIndent;
        WordList list = document.AddCustomList();
        list.Numbering.AddLevel(new WordListLevel(WordListLevelKind.DecimalDot));
        WordListLevel level = list.Numbering.Levels[0];
        level.IndentationLeft = 720;
        level.IndentationHanging = 360;
        level.OpenXmlElement.PreviousParagraphProperties!.InsertAt(new Tabs(
            new TabStop { Val = TabStopValues.Number, Position = 3000 }), 0);
        level.OpenXmlElement.NumberingSymbolRunProperties = new NumberingSymbolRunProperties(
            new RunFonts { Ascii = "Courier New", HighAnsi = "Courier New" }, new FontSize { Val = "24" });
        WordParagraph paragraph = table ? document.AddTable(1, 1).Rows[0].Cells[0].Paragraphs[0] : list.AddItem("BODY\nNEXT");
        paragraph.Text = "BODY\nNEXT";
        paragraph.FontFamily = "Courier New";
        paragraph.FontSize = 12;
        if (inline) paragraph.CharacterScale = 95;
        if (table) paragraph._paragraph.ParagraphProperties = new ParagraphProperties(new NumberingProperties(
            new NumberingLevelReference { Val = 0 }, new NumberingId { Val = list.NumberId }));
        using var pdf = PdfPigDocument.Open(document.ToPdfBytes(new WordToPdfOptions {
            IncludePageNumbers = false, FontFamily = "Courier"
        }));
        var letters = pdf.GetPage(1).Letters.Where(letter => !string.IsNullOrWhiteSpace(letter.Value)).ToArray();
        Assert.Equal("1.BODYNEXT", string.Concat(letters.Select(letter => letter.Value)));
        Assert.InRange(Math.Abs(letters[2].StartBaseLine.X - letters[0].StartBaseLine.X - bodyOffset), 0D, 0.03D);
        Assert.InRange(Math.Abs(letters[6].StartBaseLine.X - letters[0].StartBaseLine.X - 18D), 0D, 0.03D);
    }
}
