using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class PdfListMarkerAnchorTests {
    [Theory]
    [InlineData(WordListLevelAlignment.Left, 720)]
    [InlineData(WordListLevelAlignment.Right, 720)]
    [InlineData(WordListLevelAlignment.Center, 720)]
    [InlineData(WordListLevelAlignment.Right, 0)]
    [InlineData(WordListLevelAlignment.Center, 0)]
    public void TableListMarkersKeepTheirAnchorSeparateFromTheBodyIndent(WordListLevelAlignment alignment, int hangingTwips) {
        using WordDocument document = WordDocument.Create();
        WordList list = document.AddCustomList();
        list.Numbering.AddLevel(new WordListLevel(WordListLevelKind.DecimalDot).SetStartNumberingValue(12));
        WordListLevel level = list.Numbering.Levels[0];
        level.IndentationLeft = 1440;
        level.IndentationHanging = hangingTwips;
        level.LevelJustification = alignment;
        level.OpenXmlElement.NumberingSymbolRunProperties = new NumberingSymbolRunProperties(
            new RunFonts { Ascii = "Courier New", HighAnsi = "Courier New" }, new FontSize { Val = "24" });
        WordParagraph paragraph = document.AddTable(1, 1).Rows[0].Cells[0].Paragraphs[0];
        paragraph.Text = "MARKER BODY";
        paragraph.FontFamily = "Courier New";
        paragraph.FontSize = 12;
        paragraph._paragraph.ParagraphProperties = new ParagraphProperties(new NumberingProperties(
            new NumberingLevelReference { Val = 0 }, new NumberingId { Val = list.NumberId }));
        using var pdf = PdfPigDocument.Open(document.ToPdfBytes(new WordToPdfOptions {
            IncludePageNumbers = false, FontFamily = "Courier"
        }));
        var letters = pdf.GetPage(1).Letters.Where(letter => !string.IsNullOrWhiteSpace(letter.Value)).ToArray();
        Assert.Equal("12.MARKERBODY", string.Concat(letters.Select(letter => letter.Value)));
        double anchor = alignment switch {
            WordListLevelAlignment.Right => letters[2].EndBaseLine.X,
            WordListLevelAlignment.Center => (letters[0].StartBaseLine.X + letters[2].EndBaseLine.X) / 2D,
            _ => letters[0].StartBaseLine.X
        };
        if (alignment == WordListLevelAlignment.Right || hangingTwips > 0)
            Assert.InRange(Math.Abs(letters[3].StartBaseLine.X - anchor - hangingTwips / 20D), 0D, 0.03D);
    }

    [Theory]
    [InlineData(WordListLevelAlignment.Left, 720, false)]
    [InlineData(WordListLevelAlignment.Right, 720, false)]
    [InlineData(WordListLevelAlignment.Center, 720, false)]
    [InlineData(WordListLevelAlignment.Right, 0, false)]
    [InlineData(WordListLevelAlignment.Center, 0, false)]
    [InlineData(WordListLevelAlignment.Left, 720, true)]
    [InlineData(WordListLevelAlignment.Right, 720, true)]
    [InlineData(WordListLevelAlignment.Center, 720, true)]
    [InlineData(WordListLevelAlignment.Right, 0, true)]
    [InlineData(WordListLevelAlignment.Center, 0, true)]
    public void BodyListMarkersAlignAtTheAuthoredHangingAnchor(WordListLevelAlignment alignment, int hangingTwips, bool inline) {
        using WordDocument document = WordDocument.Create();
        WordList list = document.AddCustomList();
        list.Numbering.AddLevel(new WordListLevel(WordListLevelKind.DecimalDot).SetStartNumberingValue(12));
        WordListLevel level = list.Numbering.Levels[0];
        level.IndentationLeft = 1440;
        level.IndentationHanging = hangingTwips;
        level.LevelJustification = alignment;
        level.OpenXmlElement.NumberingSymbolRunProperties = new NumberingSymbolRunProperties(
            new RunFonts { Ascii = "Courier New", HighAnsi = "Courier New" }, new FontSize { Val = "24" }, new CharacterScale { Val = 100 });
        WordParagraph paragraph = list.AddItem("MARKER BODY");
        paragraph.FontFamily = "Courier New";
        paragraph.FontSize = 12;
        if (inline) paragraph.CharacterScale = 95;
        using var pdf = PdfPigDocument.Open(document.ToPdfBytes(new WordToPdfOptions {
            IncludePageNumbers = false, FontFamily = "Courier"
        }));
        var letters = pdf.GetPage(1).Letters.Where(letter => !string.IsNullOrWhiteSpace(letter.Value)).ToArray();
        Assert.Equal("12.MARKERBODY", string.Concat(letters.Select(letter => letter.Value)));
        double start = letters[0].StartBaseLine.X;
        double end = letters[2].EndBaseLine.X;
        double actualAnchor = alignment switch {
            WordListLevelAlignment.Right => end,
            WordListLevelAlignment.Center => (start + end) / 2D,
            _ => start
        };
        double expectedAnchor = 72D + 72D - hangingTwips / 20D;
        Assert.InRange(Math.Abs(actualAnchor - expectedAnchor), 0D, 0.03D);
        // A tab suffix retains the body indent when the marker ends at its anchor.
        if (alignment == WordListLevelAlignment.Right || hangingTwips > 0)
            Assert.InRange(Math.Abs(letters[3].StartBaseLine.X - 144D), 0D, 0.03D);
    }
}
