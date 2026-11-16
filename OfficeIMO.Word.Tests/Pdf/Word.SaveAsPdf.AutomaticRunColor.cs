using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(WordFileFormat.Docx, "014E78", "auto", true)]
    [InlineData(WordFileFormat.Docx, "025E23", "auto", false)]
    [InlineData(WordFileFormat.Docx, "010203", "000000", false)]
    [InlineData(WordFileFormat.Docx, "E5EFF7", "FFFFFF", true)]
    [InlineData(WordFileFormat.Doc, "010203", "auto", true)]
    public void SaveAsPdf_AutomaticRunColorUsesSolidParagraphBackground(WordFileFormat format, string fill, string textColor, bool white) {
        using WordDocument source = WordDocument.Create();
        WordParagraph paragraph = source.AddParagraph("ColorMarker");
        paragraph.ShadingFillColor = OfficeColor.Parse(fill);
        paragraph.RunColorValue = textColor;
        AssertAutomaticRunColorPdf(source, format, "ColorMarker", white, fill);
    }

    [Theory]
    [InlineData("010203", "FFFFFF", false)]
    [InlineData("E5EFF7", "000080", true)]
    public void SaveAsPdf_AutomaticRunColorUsesPaintedRunShadingBeforeParagraphFill(string paragraphFill, string runFill, bool white) {
        using WordDocument source = WordDocument.Create();
        WordParagraph paragraph = source.AddParagraph("RunShadedMarker");
        paragraph.ShadingFillColor = OfficeColor.Parse(paragraphFill);
        paragraph.BackgroundColor = OfficeColor.Parse(runFill);
        paragraph.RunColorValue = "auto";
        AssertAutomaticRunColorPdf(source, WordFileFormat.Docx, "RunShadedMarker", white, runFill);
    }

    [Theory]
    [InlineData(WordFileFormat.Docx, false)]
    [InlineData(WordFileFormat.Docx, true)]
    [InlineData(WordFileFormat.Doc, false)]
    public void SaveAsPdf_AutomaticRunColorInDerivedStyleOverridesInheritedRgb(WordFileFormat format, bool characterStyle) {
        using WordDocument source = WordDocument.Create();
        Styles styles = source._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        Style normal = styles.Elements<Style>().Single(style => style.StyleId?.Value == "Normal");
        normal.StyleRunProperties ??= new StyleRunProperties();
        normal.StyleRunProperties.Color = new Color { Val = "FF0000" };
        styles.Append(new Style(new StyleName { Val = "Automatic foreground" }, new BasedOn { Val = characterStyle ? "DefaultParagraphFont" : "Normal" },
            new StyleRunProperties(new Color { Val = "auto" })) {
            StyleId = "AutomaticForeground", Type = characterStyle ? StyleValues.Character : StyleValues.Paragraph, CustomStyle = true
        });
        WordParagraph paragraph = source.AddParagraph("StyledColorMarker");
        paragraph.ShadingFillColor = OfficeColor.Parse("010203");
        if (characterStyle) {
            Run run = paragraph._paragraph.Elements<Run>().Single();
            run.RunProperties ??= new RunProperties();
            run.RunProperties.RunStyle = new RunStyle { Val = "AutomaticForeground" };
        } else paragraph.SetStyleId("AutomaticForeground");
        AssertAutomaticRunColorPdf(source, format, "StyledColorMarker", true, "010203");
    }

    [Theory]
    [InlineData(WordFileFormat.Docx, "cell", "auto", true)]
    [InlineData(WordFileFormat.Docx, "cell", "000000", false)]
    [InlineData(WordFileFormat.Doc, "cell", "auto", true)]
    [InlineData(WordFileFormat.Docx, "paragraph", "auto", false)]
    public void SaveAsPdf_AutomaticRunColorUsesPaintedCellBackground(WordFileFormat format, string shaded, string textColor, bool white) {
        using WordDocument source = WordDocument.Create();
        WordTableCell cell = source.AddTable(1, 1).Rows[0].Cells[0];
        WordParagraph paragraph = cell.Paragraphs[0];
        paragraph.Text = "CellColorMarker";
        paragraph.RunColorValue = textColor;
        // Table-cell paragraphs render without their own fill, so only the cell fill is a backdrop.
        if (shaded == "cell") cell.ShadingFillColor = OfficeColor.Parse("000080");
        else paragraph.ShadingFillColor = OfficeColor.Parse("010203");
        AssertAutomaticRunColorPdf(source, format, "CellColorMarker", white, shaded == "cell" ? "000080" : null);
    }

    [Theory]
    [InlineData(WordFileFormat.Docx, false, "body")]
    [InlineData(WordFileFormat.Docx, true, "body")]
    [InlineData(WordFileFormat.Doc, true, "body")]
    [InlineData(WordFileFormat.Docx, false, "cell")]
    [InlineData(WordFileFormat.Docx, true, "cell")]
    [InlineData(WordFileFormat.Doc, true, "cell")]
    public void SaveAsPdf_HeadingAutomaticRunColorResolvesAgainstPaintedBackground(WordFileFormat format, bool direct, string placement) {
        using WordDocument source = WordDocument.Create();
        WordParagraph paragraph;
        if (placement == "cell") {
            WordTableCell cell = source.AddTable(1, 1).Rows[0].Cells[0];
            cell.ShadingFillColor = OfficeColor.Parse("000080");
            paragraph = cell.Paragraphs[0];
            paragraph.Text = "HeadingColorMarker";
        } else {
            paragraph = source.AddParagraph("HeadingColorMarker");
            paragraph.ShadingFillColor = OfficeColor.Parse("010203");
        }
        paragraph.Style = WordParagraphStyles.Heading1;
        Style heading = source._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!
            .Elements<Style>().Single(style => style.StyleId?.Value == "Heading1");
        heading.StyleRunProperties ??= new StyleRunProperties();
        heading.StyleRunProperties.Color = new Color { Val = direct ? "FF0000" : "auto" };
        if (direct) paragraph.RunColorValue = "auto";
        // Both the cell fill and a body heading's own fill are painted behind the heading text.
        AssertAutomaticRunColorPdf(source, format, "HeadingColorMarker", true, placement == "cell" ? "000080" : "010203");
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SaveAsPdf_ConditionalTableAutomaticRunColorOverridesInheritedRgb(bool dark) {
        using WordDocument source = WordDocument.Create();
        Styles styles = source._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        foreach (string styleId in new[] { "BaseColorTable", "AutomaticColorTable" }) {
            Style style = new(new StyleName { Val = styleId }) { StyleId = styleId, Type = StyleValues.Table, CustomStyle = true };
            if (styleId == "AutomaticColorTable") style.Append(new BasedOn { Val = "BaseColorTable" });
            style.Append(new TableStyleProperties(new RunPropertiesBaseStyle(new Color { Val = styleId == "BaseColorTable" ? "FF0000" : "auto" })) {
                Type = TableStyleOverrideValues.FirstRow
            });
            styles.Append(style);
        }
        WordTable table = source.AddTable(1, 1);
        table._tableProperties!.TableStyle = new TableStyle { Val = "AutomaticColorTable" };
        table.ConditionalFormattingFirstRow = true;
        WordTableCell cell = table.Rows[0].Cells[0];
        if (dark) cell.ShadingFillColor = OfficeColor.Parse("000080");
        cell.Paragraphs[0].Text = "ConditionalColorMarker";
        AssertAutomaticRunColorPdf(source, WordFileFormat.Docx, "ConditionalColorMarker", dark, dark ? "000080" : null);
    }

    [Theory]
    [InlineData(WordListStyle.Numbered, "1.ListColorMarker")]
    [InlineData(WordListStyle.Bulleted, "•ListColorMarker")]
    public void SaveAsPdf_ShadedListItemPaintsFillBehindAutomaticText(WordListStyle listStyle, string expectedText) {
        using WordDocument source = WordDocument.Create();
        WordParagraph item = source.AddList(listStyle).AddItem("ListColorMarker");
        item.ShadingFillColor = OfficeColor.Parse("010203");
        item.RunColorValue = "auto";
        AssertAutomaticRunColorPdf(source, WordFileFormat.Docx, expectedText, true, "010203");
    }

    [Fact]
    public void SaveAsPdf_ListMarkerIgnoresDarkRunShadingItIsNotPaintedOn() {
        using WordDocument source = WordDocument.Create();
        WordParagraph item = source.AddList(WordListStyle.Numbered).AddItem("ListRunMarker");
        item.BackgroundColor = OfficeColor.Parse("000080");
        item.RunColorValue = "auto";
        using WordDocument loaded = WordDocument.Load(new MemoryStream(source.ToBytes()));
        byte[] bytes = loaded.ToPdf(new WordPdfExportOptions { IncludePageNumbers = false }).ToBytes();
        Assert.Contains(FormatPdfFillOperator("000080"), PdfOperatorSearchText.From(bytes));
        using PdfPigDocument document = PdfPigDocument.Open(bytes);
        var letters = document.GetPages().SelectMany(page => page.Letters).ToArray();
        Assert.Equal("1.ListRunMarker", string.Concat(letters.Select(letter => letter.Value)));
        // The marker sits on the page, the item text on its painted run shading.
        Assert.All(letters.Take(2), letter => Assert.Equal(0D, letter.Color.ToRGBValues().Item1, 6));
        Assert.All(letters.Skip(2), letter => Assert.Equal(1D, letter.Color.ToRGBValues().Item1, 6));
    }

    [Fact]
    public void SaveAsPdf_NumberedShadedHeadingPaintsFillBehindAutomaticText() {
        using WordDocument source = WordDocument.Create();
        WordParagraph item = source.AddList(WordListStyle.Numbered).AddItem("HeadingListMarker");
        item.Style = WordParagraphStyles.Heading1;
        item.ShadingFillColor = OfficeColor.Parse("010203");
        item.RunColorValue = "auto";
        AssertAutomaticRunColorPdf(source, WordFileFormat.Docx, "1.HeadingListMarker", true, "010203");
    }

    [Fact]
    public void SaveAsPdf_ShadedParagraphWithTextBoxKeepsAutomaticTextVisible() {
        using WordDocument source = WordDocument.Create();
        WordParagraph paragraph = source.AddParagraph("PREFIX");
        paragraph.ShadingFillColor = OfficeColor.Parse("010203");
        paragraph.RunColorValue = "auto";
        paragraph.AddTextBox("BOXTEXT", WordImageTextWrapping.Square);
        using WordDocument loaded = WordDocument.Load(new MemoryStream(source.ToBytes()));
        byte[] bytes = loaded.ToPdf(new WordPdfExportOptions { IncludePageNumbers = false }).ToBytes();
        string commands = PdfOperatorSearchText.From(bytes);
        using PdfPigDocument document = PdfPigDocument.Open(bytes);
        var prefix = document.GetPages().SelectMany(page => page.GetWords()).Single(word => word.Text.Contains("PREFIX"));
        // White automatic text is only correct where the dark paragraph fill is actually painted.
        bool white = prefix.Letters.All(letter => letter.Color.ToRGBValues().Item1 > 0.5D);
        Assert.Equal(white, commands.Contains(FormatPdfFillOperator("010203")));
    }

    [Fact]
    public void SaveAsPdf_AutomaticRunColorUsesConditionalTableStyleCellFill() {
        using WordDocument source = WordDocument.Create();
        Style style = new(new StyleName { Val = "Navy header" }) { StyleId = "NavyHeader", Type = StyleValues.Table, CustomStyle = true };
        style.Append(new TableStyleProperties(new TableStyleConditionalFormattingTableCellProperties(
            new Shading { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "000080" })) { Type = TableStyleOverrideValues.FirstRow });
        source._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!.Append(style);
        WordTable table = source.AddTable(1, 1);
        table._tableProperties!.TableStyle = new TableStyle { Val = "NavyHeader" };
        table.ConditionalFormattingFirstRow = true;
        WordParagraph paragraph = table.Rows[0].Cells[0].Paragraphs[0];
        paragraph.Text = "HeaderColorMarker";
        paragraph.RunColorValue = "auto";
        AssertAutomaticRunColorPdf(source, WordFileFormat.Docx, "HeaderColorMarker", true, "000080");
    }

    /// <summary>Checks the marker's letter color and that the named backdrop is painted (or the dark one is absent).</summary>
    private static void AssertAutomaticRunColorPdf(WordDocument source, WordFileFormat format, string marker, bool white, string? paintedFill) {
        using WordDocument loaded = WordDocument.Load(new MemoryStream(source.ToBytes(new WordSaveOptions { Format = format })));
        byte[] bytes = loaded.ToPdf(new WordPdfExportOptions { IncludePageNumbers = false }).ToBytes();
        string commands = PdfOperatorSearchText.From(bytes);
        if (paintedFill != null) Assert.Contains(FormatPdfFillOperator(paintedFill), commands);
        else Assert.DoesNotContain(FormatPdfFillOperator("010203"), commands);
        using PdfPigDocument document = PdfPigDocument.Open(bytes);
        var letters = document.GetPages().SelectMany(page => page.Letters).ToArray();
        Assert.Equal(marker, string.Concat(letters.Select(letter => letter.Value)));
        Assert.All(letters, letter => {
            var (red, green, blue) = letter.Color.ToRGBValues();
            double expected = white ? 1D : 0D;
            Assert.Equal(expected, red, 6);
            Assert.Equal(expected, green, 6);
            Assert.Equal(expected, blue, 6);
        });
    }

    private static string FormatPdfFillOperator(string hex) {
        OfficeColor color = OfficeColor.Parse(hex);
        static string Channel(byte value) => Math.Round(value / 255D, 3).ToString("0.###", System.Globalization.CultureInfo.InvariantCulture);
        return Channel(color.R) + " " + Channel(color.G) + " " + Channel(color.B) + " rg";
    }
}
