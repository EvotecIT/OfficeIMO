using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData("body", WordListLevelSuffix.Nothing)]
    [InlineData("body", WordListLevelSuffix.Space)]
    [InlineData("body", WordListLevelSuffix.Tab)]
    [InlineData("table", WordListLevelSuffix.Nothing)]
    public void SaveAsPdf_NumberingLabelUsesBodyDescentWhenSnappingTheLineGrid(string route, WordListLevelSuffix suffix) {
        using WordDocument source = CreateNumberingMetricDocument(12.5D, suffix, grid: true, wrapped: false, table: route == "table");
        var letters = ReadNumberingMetricLetters(source);
        Assert.Equal(18D, letters[0].StartBaseLine.Y - letters[1].StartBaseLine.Y, 3);
        Assert.Equal(18D, letters[1].StartBaseLine.Y - letters[2].StartBaseLine.Y, 3);
    }

    [Theory]
    [InlineData(WordListLevelSuffix.Nothing)]
    [InlineData(WordListLevelSuffix.Space)]
    [InlineData(WordListLevelSuffix.Tab)]
    public void SaveAsPdf_LargeNumberingLabelAffectsOnlyItsFirstLine(WordListLevelSuffix suffix) {
        using WordDocument source = CreateNumberingMetricDocument(20D, suffix, grid: true, wrapped: true);
        var letters = ReadNumberingMetricLetters(source);
        Assert.Equal(22.5D, letters[0].StartBaseLine.Y - letters[1].StartBaseLine.Y, 3);
        Assert.Equal(18D, letters[1].StartBaseLine.Y - letters[2].StartBaseLine.Y, 3);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SaveAsPdf_NumberingLabelUsesItsRegisteredFontSlotWhenFamiliesShareAnAlias(bool namedBody) {
        using WordDocument source = CreateNumberingMetricDocument(12.5D, WordListLevelSuffix.Nothing, grid: false, wrapped: false);
        foreach (WordParagraph paragraph in source.Paragraphs) paragraph.FontFamily = "Calibri";
        var options = NumberingMetricPdfOptions();
        options.RegisterFontFamily(PdfStandardFont.Helvetica, new PdfEmbeddedFontFamily("Calibri",
            NumberingMetricFont(300)));
        options.RegisterFontFamily(PdfStandardFont.TimesRoman, new PdfEmbeddedFontFamily("Arial",
            NumberingMetricFont(600)));
        if (namedBody) options.RegisterNamedFontFamily(new PdfEmbeddedFontFamily("Calibri", NumberingMetricFont(300)));
        using var pdf = PdfPigDocument.Open(source.ToPdfBytes(new WordToPdfOptions {
            IncludePageNumbers = false, ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic(), PdfOptions = options
        }));
        var marker = pdf.GetPage(1).Letters.Where(letter => letter.Value is "g" or "j").ToArray();
        Assert.Equal("gjgjgj", string.Concat(marker.Select(letter => letter.Value)));
        Assert.Equal(7.5D, marker[1].StartBaseLine.X - marker[0].StartBaseLine.X, 3);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SaveAsPdf_NumberingLabelPreservesInheritedStrikeDecoration(bool table) {
        using WordDocument source = CreateNumberingMetricDocument(20D, WordListLevelSuffix.Nothing, grid: false, wrapped: false, table);
        foreach (Paragraph paragraph in source._wordprocessingDocument.MainDocumentPart!.Document.Descendants<Paragraph>()) {
            (paragraph.ParagraphProperties!.ParagraphMarkRunProperties ??= new ParagraphMarkRunProperties()).Append(new Strike());
            foreach (Run run in paragraph.Elements<Run>()) (run.RunProperties ??= new RunProperties()).Append(new Strike());
        }
        using var pdf = PdfPigDocument.Open(source.ToPdfBytes(new WordToPdfOptions {
            IncludePageNumbers = false, ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic(), PdfOptions = NumberingMetricPdfOptions()
        }));
        var page = pdf.GetPage(1);
        var marker = page.Letters.First(letter => letter.Value == "g");
        var strokes = page.Paths.Where(path => path.IsStroked).Select(path => path.GetBoundingRectangle())
            .Where(rectangle => rectangle.HasValue).Select(rectangle => rectangle!.Value);
        Assert.Contains(strokes, rectangle => rectangle.Height < .1D &&
            rectangle.Left <= marker.StartBaseLine.X && rectangle.Right > marker.StartBaseLine.X + .1D &&
            rectangle.Bottom > marker.StartBaseLine.Y && rectangle.Bottom < marker.StartBaseLine.Y + 16D);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SaveAsPdf_HalfSpacingRaisesTheBaselineByHalfTheBodyFontAscent(bool numbered) {
        using WordDocument full = CreateNumberingMetricDocument(20D, WordListLevelSuffix.Tab, grid: false, wrapped: true);
        using WordDocument half = CreateNumberingMetricDocument(20D, WordListLevelSuffix.Tab, grid: false, wrapped: true);
        half.Paragraphs[0].LineSpacing = 120;
        if (!numbered) {
            full.Paragraphs[0]._paragraph.ParagraphProperties!.RemoveAllChildren<NumberingProperties>();
            half.Paragraphs[0]._paragraph.ParagraphProperties!.RemoveAllChildren<NumberingProperties>();
        }
        var fullLetters = ReadNumberingMetricLetters(full);
        var halfLetters = ReadNumberingMetricLetters(half);
        Assert.Equal(5.5D, halfLetters[0].StartBaseLine.Y - fullLetters[0].StartBaseLine.Y, 3);
    }

    [Theory]
    [InlineData(WordListLevelSuffix.Nothing)]
    [InlineData(WordListLevelSuffix.Space)]
    [InlineData(WordListLevelSuffix.Tab)]
    public void SaveAsPdf_MultipleSpacingScalesBodyMetricsBelowALargeNumberingLabel(WordListLevelSuffix suffix) {
        using WordDocument source = CreateNumberingMetricDocument(20D, suffix, grid: false, wrapped: true);
        source.Paragraphs[0].LineSpacing = 360;
        var letters = ReadNumberingMetricLetters(source);
        Assert.Equal(24.75D, letters[0].StartBaseLine.Y - letters[1].StartBaseLine.Y, 3);
        Assert.Equal(24.75D, letters[1].StartBaseLine.Y - letters[2].StartBaseLine.Y, 3);
    }

    [Theory]
    [InlineData(false, false, "", 25.5D)]
    [InlineData(true, false, "", 36D)]
    [InlineData(false, true, "", 25.5D)]
    [InlineData(true, true, "", 36D)]
    [InlineData(false, false, "   ", 25.5D)]
    [InlineData(true, false, "   ", 36D)]
    [InlineData(false, true, "   ", 25.5D)]
    [InlineData(true, true, "   ", 36D)]
    [InlineData(false, false, "\t", 25.5D)]
    [InlineData(true, false, "\t", 36D)]
    [InlineData(false, true, "\t", 25.5D)]
    [InlineData(true, true, "\t", 36D)]
    public void SaveAsPdf_EmptyNumberedItemRetainsItsParagraphMarkMetrics(bool grid, bool table, string whitespace, double expected) {
        using WordDocument source = CreateNumberingMetricDocument(20D, WordListLevelSuffix.Tab, grid, wrapped: false, table);
        var main = source._wordprocessingDocument.MainDocumentPart!;
        Text firstText = main.Document.Descendants<Text>().First();
        if (whitespace == "\t") {
            firstText.Parent!.Append(new TabChar());
            firstText.Remove();
        } else {
            firstText.Text = whitespace;
            firstText.Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve;
        }
        string original = main.Document.OuterXml;
        using var pdf = PdfPigDocument.Open(source.ToPdfBytes(new WordToPdfOptions {
            IncludePageNumbers = false, ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic(), PdfOptions = NumberingMetricPdfOptions()
        }));
        Assert.Equal(original, main.Document.OuterXml);
        var markers = pdf.GetPages().SelectMany(page => page.Letters).Where(letter => letter.Value == "g").ToArray();
        Assert.Equal(3, markers.Length);
        Assert.Equal(expected, markers[0].StartBaseLine.Y - markers[1].StartBaseLine.Y, 3);
        Assert.Equal(expected, markers[1].StartBaseLine.Y - markers[2].StartBaseLine.Y, 3);
    }

    [Theory]
    [InlineData(false, false, 4D, 10.5D)]
    [InlineData(false, true, 4D, 10.5D)]
    [InlineData(true, false, -1.25D, 0D)]
    [InlineData(true, true, -1.25D, 0D)]
    public void SaveAsPdf_EmptyNumberedItemUsesExplicitMarkSize(bool grid, bool table, double firstShift, double nextShift) {
        using WordDocument normal = CreateNumberingMetricDocument(20D, WordListLevelSuffix.Tab, grid, false, table);
        using WordDocument marked = CreateNumberingMetricDocument(20D, WordListLevelSuffix.Tab, grid, false, table);
        foreach (WordDocument source in new[] { normal, marked })
            source._wordprocessingDocument.MainDocumentPart!.Document.Descendants<Text>().First().Text = string.Empty;
        Paragraph paragraph = marked._wordprocessingDocument.MainDocumentPart!.Document.Descendants<Paragraph>().First();
        (paragraph.ParagraphProperties!.ParagraphMarkRunProperties ??= new ParagraphMarkRunProperties()).Append(new FontSize { Val = "48" });
        var normalMarkers = ReadEmptyNumberingMarkers(normal, NumberingMetricPdfOptions());
        var markedMarkers = ReadEmptyNumberingMarkers(marked, NumberingMetricPdfOptions());
        Assert.Equal(firstShift, normalMarkers[0].StartBaseLine.Y - markedMarkers[0].StartBaseLine.Y, 3);
        Assert.Equal(nextShift, normalMarkers[1].StartBaseLine.Y - markedMarkers[1].StartBaseLine.Y, 3);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SaveAsPdf_EmptyNumberedItemUsesExplicitMarkFont(bool table) {
        using WordDocument source = CreateNumberingMetricDocument(20D, WordListLevelSuffix.Tab, false, false, table);
        Text firstText = source._wordprocessingDocument.MainDocumentPart!.Document.Descendants<Text>().First();
        firstText.Text = string.Empty;
        Paragraph paragraph = firstText.Ancestors<Paragraph>().First();
        (paragraph.ParagraphProperties!.ParagraphMarkRunProperties ??= new ParagraphMarkRunProperties()).Append(
            new RunFonts { Ascii = "Times New Roman", HighAnsi = "Times New Roman" });
        var options = NumberingMetricPdfOptions();
        options.EmbedStandardFont(PdfStandardFont.TimesRoman, ManagedTextShapingTestAssets.CreateFontWithLineBoxMetrics(
            1000, -1000, 0, 1000, Enumerable.Range(32, 95).ToArray()), "OfficeIMO-Mark-DeepDescent");
        var markers = ReadEmptyNumberingMarkers(source, options);
        Assert.Equal(31D, markers[0].StartBaseLine.Y - markers[1].StartBaseLine.Y, 3);
        Assert.Equal(25.5D, markers[1].StartBaseLine.Y - markers[2].StartBaseLine.Y, 3);
    }

    private static UglyToad.PdfPig.Content.Letter[] ReadEmptyNumberingMarkers(WordDocument source, PdfOptions options, int count = 3) {
        string original = source._wordprocessingDocument.MainDocumentPart!.Document.OuterXml;
        using var pdf = PdfPigDocument.Open(source.ToPdfBytes(new WordToPdfOptions {
            IncludePageNumbers = false, ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic(), PdfOptions = options
        }));
        Assert.Equal(original, source._wordprocessingDocument.MainDocumentPart.Document.OuterXml);
        var markers = pdf.GetPages().SelectMany(page => page.Letters).Where(letter => letter.Value == "g").ToArray();
        Assert.Equal(count, markers.Length);
        return markers;
    }

    [Theory]
    [InlineData(false, false, 51.5D)]
    [InlineData(false, true, 51.5D)]
    [InlineData(true, false, 62.75D)]
    [InlineData(true, true, 62.75D)]
    public void SaveAsPdf_TrailingNumberedLineBreakRetainsTheMarkLine(bool grid, bool table, double expected) {
        using WordDocument source = CreateNumberingMetricDocument(20D, WordListLevelSuffix.Tab, grid, false, table);
        Text firstText = source._wordprocessingDocument.MainDocumentPart!.Document.Descendants<Text>().First();
        Run firstRun = (Run)firstText.Parent!;
        firstRun.RunProperties!.FontSize = new FontSize { Val = "60" };
        firstRun.Append(new Break());
        firstText.Remove();
        var markers = ReadEmptyNumberingMarkers(source, NumberingMetricPdfOptions());
        Assert.Equal(expected, markers[0].StartBaseLine.Y - markers[1].StartBaseLine.Y, 3);
        Assert.Equal(grid ? 36D : 25.5D, markers[1].StartBaseLine.Y - markers[2].StartBaseLine.Y, 3);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void SaveAsPdf_TrailingOrdinaryLineBreakRetainsTheMarkLine(bool grid, bool table) {
        using WordDocument empty = CreateNumberingMetricDocument(20D, WordListLevelSuffix.Tab, grid, false, table);
        using WordDocument broken = CreateNumberingMetricDocument(20D, WordListLevelSuffix.Tab, grid, false, table);
        foreach (WordDocument source in new[] { empty, broken }) {
            Text firstText = source._wordprocessingDocument.MainDocumentPart!.Document.Descendants<Text>().First();
            firstText.Ancestors<Paragraph>().First().ParagraphProperties!.RemoveAllChildren<NumberingProperties>();
            firstText.Text = string.Empty;
            if (source == broken) {
                Run run = (Run)firstText.Parent!;
                run.RunProperties!.FontSize = new FontSize { Val = "60" };
                run.Append(new Break());
                firstText.Remove();
            }
        }
        var emptyMarkers = ReadEmptyNumberingMarkers(empty, NumberingMetricPdfOptions(), 2);
        var brokenMarkers = ReadEmptyNumberingMarkers(broken, NumberingMetricPdfOptions(), 2);
        Assert.Equal(grid ? 54D : 45D, emptyMarkers[0].StartBaseLine.Y - brokenMarkers[0].StartBaseLine.Y, 3);
    }

    private static WordDocument CreateNumberingMetricDocument(double markerSize, WordListLevelSuffix suffix,
        bool grid, bool wrapped, bool table = false) {
        WordDocument source = WordDocument.Create();
        if (grid) source.Sections[0]._sectionProperties.Append(new DocGrid { Type = DocGridValues.Lines, LinePitch = 360 });
        source._wordprocessingDocument.MainDocumentPart!.DocumentSettingsPart!.Settings!
            .GetFirstChild<Compatibility>()!.Append(new AdjustLineHeightInTable());
        WordList list = source.AddCustomList();
        var level = new WordListLevel(WordListLevelKind.DecimalDot) { LevelText = "gj", LevelSuffix = suffix };
        level.OpenXmlElement.NumberingSymbolRunProperties = new NumberingSymbolRunProperties(
            new RunFonts { Ascii = "Arial", HighAnsi = "Arial" }, new FontSize { Val = ((int)(markerSize * 2)).ToString() });
        list.Numbering.AddLevel(level);
        foreach (string text in wrapped ? new[] { "A\nB\nC" } : new[] { "A", "B", "C" }) {
            WordParagraph paragraph = list.AddItem(text);
            paragraph.FontFamily = "Arial"; paragraph.FontSize = 11;
            paragraph.LineSpacing = 240; paragraph.LineSpacingRule = WordLineSpacingRule.Auto;
            paragraph.LineSpacingBeforePoints = 0; paragraph.LineSpacingAfterPoints = 0;
        }
        if (table) {
            WordParagraph[] paragraphs = source.Paragraphs.ToArray();
            TableCell cell = source.AddTable(1, 1).Rows[0].Cells[0]._tableCell;
            cell.RemoveAllChildren<Paragraph>();
            foreach (WordParagraph paragraph in paragraphs) { paragraph._paragraph.Remove(); cell.Append(paragraph._paragraph); }
        }
        return source;
    }

    private static byte[] NumberingMetricFont(int width) {
        byte[] font = ManagedTextShapingTestAssets.CreateFontWithLineBoxMetrics(
            1000, -500, 0, 500, Enumerable.Range(32, 95).ToArray());
        int count = font[4] * 256 + font[5];
        for (int index = 0; index < count; index++) {
            int record = 12 + index * 16;
            if (System.Text.Encoding.ASCII.GetString(font, record, 4) != "hmtx") continue;
            int offset = (font[record + 8] << 24) | (font[record + 9] << 16) | (font[record + 10] << 8) | font[record + 11];
            font[offset] = (byte)(width >> 8); font[offset + 1] = (byte)width;
            return font;
        }
        throw new InvalidOperationException("The numbering fixture requires horizontal font metrics.");
    }

    private static PdfOptions NumberingMetricPdfOptions() {
        var options = new PdfOptions { DefaultFont = PdfStandardFont.Helvetica };
        options.EmbedStandardFont(PdfStandardFont.Helvetica, NumberingMetricFont(300), "OfficeIMO-Numbering-Regular");
        return options;
    }

    private static UglyToad.PdfPig.Content.Letter[] ReadNumberingMetricLetters(WordDocument source) {
        string original = source._wordprocessingDocument.MainDocumentPart!.Document.OuterXml;
        using var pdf = PdfPigDocument.Open(source.ToPdfBytes(new WordToPdfOptions {
            IncludePageNumbers = false, ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic(), PdfOptions = NumberingMetricPdfOptions()
        }));
        Assert.Equal(original, source._wordprocessingDocument.MainDocumentPart.Document.OuterXml);
        var letters = pdf.GetPages().SelectMany(page => page.Letters).Where(letter => letter.Value is "A" or "B" or "C").ToArray();
        Assert.Equal("ABC", string.Concat(letters.Select(letter => letter.Value)));
        return letters;
    }
}
