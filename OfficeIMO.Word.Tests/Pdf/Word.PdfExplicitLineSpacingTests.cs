using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SaveAsPdf_EmptyParagraphUsesItsMarkFontMetrics(bool nativeDoc) {
        const string markFamily = "OfficeIMO Tall Paragraph Mark";
        using WordDocument source = WordDocument.Create();
        var first = source.AddParagraph("A");
        var blank = source.AddParagraph();
        var last = source.AddParagraph("B");
        foreach (var paragraph in new[] { first, blank, last }) {
            paragraph.LineSpacingPoints = 20;
            paragraph.LineSpacingRule = WordLineSpacingRule.AtLeast;
            paragraph.LineSpacingBeforePoints = 0; paragraph.LineSpacingAfterPoints = 0;
            foreach (Run run in paragraph._paragraph.Descendants<Run>())
                (run.RunProperties ??= new RunProperties()).FontSize = new FontSize { Val = "16" };
        }
        blank._paragraph.RemoveAllChildren<Run>();
        blank._paragraph.ParagraphProperties!.ParagraphMarkRunProperties = new ParagraphMarkRunProperties(
            new RunFonts { Ascii = markFamily, HighAnsi = markFamily }, new FontSize { Val = "64" });
        using WordDocument document = WordDocument.Load(new MemoryStream(nativeDoc ? source.ToBytes(WordFileFormat.Doc) : source.ToBytes()));
        var options = new PdfOptions { DefaultFont = PdfStandardFont.Helvetica };
        options.EmbedStandardFont(PdfStandardFont.Helvetica,
            ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(Enumerable.Range(32, 95).ToArray()), "OfficeIMO-Portable-Regular");
        options.RegisterNamedFontFamily(new PdfEmbeddedFontFamily(markFamily, CreateFontWithLineMetrics(1200, -300, ' ')));
        using var pdf = PdfPigDocument.Open(document.ToPdfBytes(new WordToPdfOptions {
            IncludePageNumbers = false, ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic(), PdfOptions = options
        }));
        var letters = pdf.GetPage(1).Letters;
        // The visible line advances by 20pt; the 32pt mark uses its own 1.5-em font box (48pt).
        Assert.Equal(68D, Assert.Single(letters, letter => letter.Value == "A").StartBaseLine.Y -
            Assert.Single(letters, letter => letter.Value == "B").StartBaseLine.Y, 3);
    }

    [Theory]
    [InlineData(false, false, false, 52D)]
    [InlineData(false, false, true, 40D)]
    [InlineData(false, true, false, 52D)]
    [InlineData(false, true, true, 40D)]
    [InlineData(true, false, false, 52D)]
    [InlineData(true, false, true, 40D)]
    [InlineData(true, true, false, 52D)]
    [InlineData(true, true, true, 40D)]
    public void SaveAsPdf_EmptyParagraphMarksUseExplicitSpacingUnits(bool nativeDoc, bool table, bool exact, double expected) {
        using WordDocument source = WordDocument.Create();
        WordTableCell? cell = table ? source.AddTable(1, 1).Rows[0].Cells[0] : null;
        var first = cell == null ? source.AddParagraph("A") : cell.AddParagraph("A", removeExistingParagraphs: true);
        var blank = cell == null ? source.AddParagraph() : cell.AddParagraph();
        var last = cell == null ? source.AddParagraph("B") : cell.AddParagraph("B");
        foreach (var paragraph in new[] { first, blank, last }) {
            paragraph.LineSpacingPoints = 20;
            paragraph.LineSpacingRule = exact ? WordLineSpacingRule.Exact : WordLineSpacingRule.AtLeast;
            paragraph.LineSpacingBeforePoints = 0; paragraph.LineSpacingAfterPoints = 0;
            foreach (Run run in paragraph._paragraph.Descendants<Run>())
                (run.RunProperties ??= new RunProperties()).FontSize = new FontSize { Val = "16" };
        }
        blank._paragraph.RemoveAllChildren<Run>();
        blank._paragraph.ParagraphProperties!.ParagraphMarkRunProperties = new ParagraphMarkRunProperties(new FontSize { Val = "64" });
        using WordDocument document = WordDocument.Load(new MemoryStream(nativeDoc ? source.ToBytes(WordFileFormat.Doc) : source.ToBytes()));
        var options = new PdfOptions { DefaultFont = PdfStandardFont.Helvetica };
        options.EmbedStandardFont(PdfStandardFont.Helvetica,
            ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(Enumerable.Range(32, 95).ToArray()), "OfficeIMO-Portable-Regular");
        using var pdf = PdfPigDocument.Open(document.ToPdfBytes(new WordToPdfOptions {
            IncludePageNumbers = false, ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic(), PdfOptions = options
        }));
        var letters = pdf.GetPage(1).Letters;
        Assert.Equal(expected, Assert.Single(letters, letter => letter.Value == "A").StartBaseLine.Y -
            Assert.Single(letters, letter => letter.Value == "B").StartBaseLine.Y, 3);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SaveAsPdf_ZeroMinimumLineSpacingUsesNaturalFontHeight(bool nativeDoc) {
        using WordDocument source = WordDocument.Create();
        var paragraph = source.AddParagraph("A\nB");
        foreach (Run run in paragraph._paragraph.Descendants<Run>())
            (run.RunProperties ??= new RunProperties()).FontSize = new FontSize { Val = "16" };
        paragraph.LineSpacingPoints = 0; paragraph.LineSpacingRule = WordLineSpacingRule.AtLeast;
        using WordDocument document = WordDocument.Load(new MemoryStream(nativeDoc ? source.ToBytes(WordFileFormat.Doc) : source.ToBytes()));
        Assert.Equal(0, document.Paragraphs[0].LineSpacing);
        Assert.Equal(WordLineSpacingRule.AtLeast, document.Paragraphs[0].LineSpacingRule);
        var options = new PdfOptions { DefaultFont = PdfStandardFont.Helvetica };
        options.EmbedStandardFont(PdfStandardFont.Helvetica,
            ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(Enumerable.Range(32, 95).ToArray()), "OfficeIMO-Portable-Regular");
        using var pdf = PdfPigDocument.Open(document.ToPdfBytes(new WordToPdfOptions {
            IncludePageNumbers = false, ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic(), PdfOptions = options
        }));
        var letters = pdf.GetPage(1).Letters;
        Assert.Equal(8D, Assert.Single(letters, letter => letter.Value == "A").StartBaseLine.Y -
            Assert.Single(letters, letter => letter.Value == "B").StartBaseLine.Y, 3);
    }

    [Theory]
    [InlineData(false, false, false, 52D)]
    [InlineData(false, false, true, 40D)]
    [InlineData(false, true, false, 52D)]
    [InlineData(false, true, true, 40D)]
    [InlineData(true, false, false, 52D)]
    [InlineData(true, false, true, 40D)]
    [InlineData(true, true, false, 52D)]
    [InlineData(true, true, true, 40D)]
    public void SaveAsPdf_PreservesFormattedBlankLines(bool nativeDoc, bool table, bool exact, double expected) {
        using WordDocument source = WordDocument.Create();
        WordParagraph paragraph = table ? source.AddTable(1, 1).Rows[0].Cells[0].Paragraphs[0] : source.AddParagraph();
        SetExplicitSpacingBlankLineRuns(paragraph, "A", "B", exact ? WordLineSpacingRule.Exact : WordLineSpacingRule.AtLeast);
        string original = source._wordprocessingDocument.MainDocumentPart!.Document.OuterXml;
        using WordDocument document = WordDocument.Load(new MemoryStream(nativeDoc ? source.ToBytes(WordFileFormat.Doc) : source.ToBytes()));
        var options = new PdfOptions { DefaultFont = PdfStandardFont.Helvetica };
        options.EmbedStandardFont(PdfStandardFont.Helvetica,
            ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(Enumerable.Range(32, 95).ToArray()), "OfficeIMO-Portable-Regular");
        using var pdf = PdfPigDocument.Open(document.ToPdfBytes(new WordToPdfOptions {
            IncludePageNumbers = false, ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic(), PdfOptions = options
        }));
        Assert.Equal(original, source._wordprocessingDocument.MainDocumentPart.Document.OuterXml);
        Assert.Equal(1, pdf.NumberOfPages);
        var letters = pdf.GetPage(1).Letters;
        Assert.Equal(expected, Assert.Single(letters, letter => letter.Value == "A").StartBaseLine.Y -
            Assert.Single(letters, letter => letter.Value == "B").StartBaseLine.Y, 3);
    }

    [Fact]
    public void SaveAsPdf_ListItemsPreserveDistinctExactAndMinimumSpacingRules() {
        using WordDocument source = WordDocument.Create();
        var list = source.AddList(WordListStyle.Bulleted);
        foreach (var (first, last, rule) in new[] { ("A", "B", WordLineSpacingRule.Exact), ("C", "D", WordLineSpacingRule.AtLeast) }) {
            var paragraph = list.AddItem(string.Empty);
            SetExplicitSpacingBlankLineRuns(paragraph, first, last, rule);
        }
        var options = new PdfOptions { DefaultFont = PdfStandardFont.Helvetica };
        options.EmbedStandardFont(PdfStandardFont.Helvetica,
            ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(Enumerable.Range(32, 95).Concat(new[] { 0x2022 }).ToArray()), "OfficeIMO-Portable-Regular");
        using var pdf = PdfPigDocument.Open(source.ToPdfBytes(new WordToPdfOptions {
            IncludePageNumbers = false, ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic(), PdfOptions = options
        }));
        var letters = pdf.GetPage(1).Letters;
        double Y(string marker) => Assert.Single(letters, letter => letter.Value == marker).StartBaseLine.Y;
        Assert.Equal(40D, Y("A") - Y("B"), 3);
        Assert.Equal(52D, Y("C") - Y("D"), 3);
    }

    [Theory]
    [InlineData(false, "body", WordLineSpacingRule.Auto, 8D)]
    [InlineData(true, "body", WordLineSpacingRule.Auto, 8D)]
    [InlineData(false, "table", WordLineSpacingRule.Auto, 8D)]
    [InlineData(true, "table", WordLineSpacingRule.Auto, 8D)]
    [InlineData(false, "columns", WordLineSpacingRule.Auto, 8D)]
    [InlineData(true, "columns", WordLineSpacingRule.Auto, 8D)]
    [InlineData(false, "body", WordLineSpacingRule.Exact, 14D)]
    [InlineData(true, "body", WordLineSpacingRule.Exact, 14D)]
    [InlineData(false, "table", WordLineSpacingRule.Exact, 14D)]
    [InlineData(true, "table", WordLineSpacingRule.Exact, 14D)]
    [InlineData(false, "columns", WordLineSpacingRule.Exact, 14D)]
    [InlineData(true, "columns", WordLineSpacingRule.Exact, 14D)]
    [InlineData(false, "body", WordLineSpacingRule.AtLeast, 20D)]
    [InlineData(true, "body", WordLineSpacingRule.AtLeast, 20D)]
    [InlineData(false, "table", WordLineSpacingRule.AtLeast, 20D)]
    [InlineData(true, "table", WordLineSpacingRule.AtLeast, 20D)]
    [InlineData(false, "columns", WordLineSpacingRule.AtLeast, 20D)]
    [InlineData(true, "columns", WordLineSpacingRule.AtLeast, 20D)]
    public void SaveAsPdf_SmallFontParagraphsKeepTheirLineSpacingUnits(bool nativeDoc, string route, WordLineSpacingRule rule, double expected) {
        using WordDocument source = WordDocument.Create();
        WordParagraph paragraph;
        if (route == "table") {
            paragraph = source.AddTable(1, 1).Rows[0].Cells[0].Paragraphs[0];
            paragraph.Text = "A\nB\nC";
        } else {
            paragraph = source.AddParagraph("A\nB\nC");
            if (route == "columns") source.Sections[0].ColumnCount = 2;
        }
        foreach (Run run in paragraph._paragraph.Descendants<Run>()) {
            (run.RunProperties ??= new RunProperties()).FontSize = new FontSize { Val = "16" };
            run.RunProperties.RunFonts = new RunFonts { Ascii = "Arial", HighAnsi = "Arial" };
        }
        paragraph.LineSpacing = rule == WordLineSpacingRule.Auto ? 240 : (int)(expected * 20);
        paragraph.LineSpacingRule = rule;
        paragraph.LineSpacingBeforePoints = 0; paragraph.LineSpacingAfterPoints = 0;
        string original = source._wordprocessingDocument.MainDocumentPart!.Document.OuterXml;
        using WordDocument document = WordDocument.Load(new MemoryStream(nativeDoc ? source.ToBytes(WordFileFormat.Doc) : source.ToBytes()));
        var options = new PdfOptions { DefaultFont = PdfStandardFont.Helvetica };
        options.EmbedStandardFont(PdfStandardFont.Helvetica,
            ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(Enumerable.Range(32, 95).ToArray()), "OfficeIMO-Portable-Regular");
        using var pdf = PdfPigDocument.Open(document.ToPdfBytes(new WordToPdfOptions {
            IncludePageNumbers = false, ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic(), PdfOptions = options
        }));
        Assert.Equal(original, source._wordprocessingDocument.MainDocumentPart.Document.OuterXml);
        Assert.Equal(1, pdf.NumberOfPages);
        var letters = pdf.GetPage(1).Letters.Where(letter => letter.Value is "A" or "B" or "C").ToArray();
        Assert.Equal("ABC", string.Concat(letters.Select(letter => letter.Value)));
        Assert.Equal(expected, letters[0].StartBaseLine.Y - letters[1].StartBaseLine.Y, 3);
        Assert.Equal(expected, letters[1].StartBaseLine.Y - letters[2].StartBaseLine.Y, 3);
    }

    private static void SetExplicitSpacingBlankLineRuns(WordParagraph paragraph, string first, string last, WordLineSpacingRule rule) {
        paragraph._paragraph.RemoveAllChildren<Run>();
        paragraph._paragraph.Append(
            new Run(new RunProperties(new FontSize { Val = "16" }), new Text(first), new Break()),
            new Run(new RunProperties(new FontSize { Val = "64" }), new Break()),
            new Run(new RunProperties(new FontSize { Val = "16" }), new Text(last)));
        paragraph.LineSpacingPoints = 20; paragraph.LineSpacingRule = rule;
        paragraph.LineSpacingBeforePoints = 0; paragraph.LineSpacingAfterPoints = 0;
    }
}
