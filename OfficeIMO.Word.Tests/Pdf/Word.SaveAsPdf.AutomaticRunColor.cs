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
            // Body headings render without their paragraph fill, so automatic text stays black.
            paragraph.ShadingFillColor = OfficeColor.Parse("010203");
        }
        paragraph.Style = WordParagraphStyles.Heading1;
        Style heading = source._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!
            .Elements<Style>().Single(style => style.StyleId?.Value == "Heading1");
        heading.StyleRunProperties ??= new StyleRunProperties();
        heading.StyleRunProperties.Color = new Color { Val = direct ? "FF0000" : "auto" };
        if (direct) paragraph.RunColorValue = "auto";
        bool painted = placement == "cell";
        AssertAutomaticRunColorPdf(source, format, "HeadingColorMarker", painted, painted ? "000080" : null);
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
