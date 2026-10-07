using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class PdfListMarkerAnchorTests {
    [Theory]
    [InlineData(WordListLevelAlignment.Left, WordListLevelSuffix.Nothing)]
    [InlineData(WordListLevelAlignment.Center, WordListLevelSuffix.Nothing)]
    [InlineData(WordListLevelAlignment.Right, WordListLevelSuffix.Nothing)]
    [InlineData(WordListLevelAlignment.Left, WordListLevelSuffix.Space)]
    [InlineData(WordListLevelAlignment.Center, WordListLevelSuffix.Space)]
    [InlineData(WordListLevelAlignment.Right, WordListLevelSuffix.Space)]
    public void SpaceAndNothingSuffixesFollowEachItemsActualMarkerWidth(WordListLevelAlignment alignment, WordListLevelSuffix suffix) {
        using WordDocument document = WordDocument.Create();
        WordList list = document.AddCustomList();
        list.Numbering.AddLevel(new WordListLevel(WordListLevelKind.DecimalDot).SetStartNumberingValue(9));
        WordListLevel level = list.Numbering.Levels[0];
        level.IndentationLeft = 1440;
        level.IndentationHanging = 720;
        level.LevelJustification = alignment;
        level.LevelSuffix = suffix;
        level.OpenXmlElement.NumberingSymbolRunProperties = new NumberingSymbolRunProperties(
            new RunFonts { Ascii = "Courier New", HighAnsi = "Courier New" }, new FontSize { Val = "24" });
        foreach (string text in new[] { "FIRST", "SECOND" }) {
            WordParagraph paragraph = list.AddItem(text);
            paragraph.FontFamily = "Courier New";
            paragraph.FontSize = 12;
            paragraph._paragraph.ParagraphProperties ??= new ParagraphProperties();
            paragraph._paragraph.ParagraphProperties.ParagraphMarkRunProperties = new ParagraphMarkRunProperties(
                new RunFonts { Ascii = "Arial", HighAnsi = "Arial" }, new FontSize { Val = "24" });
        }
        using var pdf = PdfPigDocument.Open(document.ToPdfBytes(new WordToPdfOptions {
            IncludePageNumbers = false, FontFamily = "Courier"
        }));
        var lines = pdf.GetPage(1).Letters.Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
            .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 2)).ToArray();
        Assert.Equal(2, lines.Length);
        for (int i = 0; i < lines.Length; i++) {
            var letters = lines[i].ToArray();
            int markerLength = i == 0 ? 2 : 3;
            Assert.Equal(i == 0 ? "9.FIRST" : "10.SECOND", string.Concat(letters.Select(letter => letter.Value)));
            double gap = letters[markerLength].StartBaseLine.X - letters[markerLength - 1].EndBaseLine.X;
            // Word's space suffix follows paragraph-mark typography, independently
            // of the Courier marker and body run.
            Assert.InRange(Math.Abs(gap - (suffix == WordListLevelSuffix.Space ? 3.336D : 0D)), 0D, 0.03D);
        }
    }

    [Theory]
    [InlineData(false, 0)]
    [InlineData(true, 0)]
    [InlineData(false, 40)]
    [InlineData(true, 40)]
    public void ListMarkerTrackingUsesTheParagraphMarkInsteadOfDirectBodyRunFormatting(bool table, int markerSpacingTwips) {
        using WordDocument document = WordDocument.Create();
        WordList list = document.AddCustomList();
        list.Numbering.AddLevel(new WordListLevel(WordListLevelKind.DecimalDot).SetStartNumberingValue(12));
        WordListLevel level = list.Numbering.Levels[0];
        level.IndentationLeft = 1440;
        level.IndentationHanging = 720;
        level.LevelJustification = WordListLevelAlignment.Right;
        level.OpenXmlElement.NumberingSymbolRunProperties = new NumberingSymbolRunProperties(
            new RunFonts { Ascii = "Courier New", HighAnsi = "Courier New" }, new FontSize { Val = "24" });
        WordParagraph paragraph = table ? document.AddTable(1, 1).Rows[0].Cells[0].Paragraphs[0] : list.AddItem("MARKER BODY");
        paragraph.Text = "MARKER BODY";
        paragraph.FontFamily = "Courier New";
        paragraph.FontSize = 12;
        paragraph.Spacing = 20;
        paragraph._paragraph.ParagraphProperties ??= new ParagraphProperties();
        paragraph._paragraph.ParagraphProperties.NumberingProperties = new NumberingProperties(
            new NumberingLevelReference { Val = 0 }, new NumberingId { Val = list.NumberId });
        paragraph._paragraph.ParagraphProperties.ParagraphMarkRunProperties = new ParagraphMarkRunProperties(
            new Spacing { Val = markerSpacingTwips });
        Assert.Empty(document.ValidateDocument());
        using var pdf = PdfPigDocument.Open(document.ToPdfBytes(new WordToPdfOptions {
            IncludePageNumbers = false, FontFamily = "Courier"
        }));
        var letters = pdf.GetPage(1).Letters.Where(letter => !string.IsNullOrWhiteSpace(letter.Value)).ToArray();
        Assert.Equal("12.MARKERBODY", string.Concat(letters.Select(letter => letter.Value)));
        double markerAdvance = letters[1].StartBaseLine.X - letters[0].StartBaseLine.X;
        double bodyAdvance = letters[4].StartBaseLine.X - letters[3].StartBaseLine.X;
        Assert.InRange(Math.Abs(markerAdvance - 7.2D - markerSpacingTwips / 20D), 0D, 0.03D);
        Assert.InRange(Math.Abs(bodyAdvance - 8.2D), 0D, 0.03D);
    }

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
        Assert.InRange(Math.Abs(letters[3].StartBaseLine.Y - letters[0].StartBaseLine.Y), 0D, 0.03D);
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
