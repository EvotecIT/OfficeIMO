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
        AssertAutomaticRunColorPdf(source, format, "ColorMarker", white);
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
        AssertAutomaticRunColorPdf(source, format, "StyledColorMarker", true);
    }

    [Theory]
    [InlineData(WordFileFormat.Docx, false)]
    [InlineData(WordFileFormat.Docx, true)]
    [InlineData(WordFileFormat.Doc, false)]
    public void SaveAsPdf_AutomaticRunColorUsesCellBackgroundAndRetainsExplicitBlack(WordFileFormat format, bool explicitBlack) {
        using WordDocument source = WordDocument.Create();
        WordTableCell cell = source.AddTable(1, 1).Rows[0].Cells[0];
        cell.ShadingFillColor = OfficeColor.Parse("000000");
        WordParagraph paragraph = cell.Paragraphs[0];
        paragraph.Text = "CellColorMarker";
        paragraph.RunColorValue = explicitBlack ? "000000" : "auto";
        AssertAutomaticRunColorPdf(source, format, "CellColorMarker", !explicitBlack);
    }

    [Theory]
    [InlineData(WordFileFormat.Docx, false, false)]
    [InlineData(WordFileFormat.Docx, false, true)]
    [InlineData(WordFileFormat.Docx, true, true)]
    [InlineData(WordFileFormat.Doc, true, true)]
    public void SaveAsPdf_HeadingAutomaticRunColorResolvesDirectAndStyleDeclarations(WordFileFormat format, bool direct, bool dark) {
        using WordDocument source = WordDocument.Create();
        WordParagraph paragraph = source.AddParagraph("HeadingColorMarker");
        paragraph.Style = WordParagraphStyles.Heading1;
        Style heading = source._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!
            .Elements<Style>().Single(style => style.StyleId?.Value == "Heading1");
        heading.StyleRunProperties ??= new StyleRunProperties();
        heading.StyleRunProperties.Color = new Color { Val = direct ? "FF0000" : "auto" };
        if (direct) paragraph.RunColorValue = "auto";
        if (dark) paragraph.ShadingFillColor = OfficeColor.Parse("010203");
        AssertAutomaticRunColorPdf(source, format, "HeadingColorMarker", dark);
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
        if (dark) cell.ShadingFillColor = OfficeColor.Parse("000000");
        cell.Paragraphs[0].Text = "ConditionalColorMarker";
        AssertAutomaticRunColorPdf(source, WordFileFormat.Docx, "ConditionalColorMarker", dark);
    }

    private static void AssertAutomaticRunColorPdf(WordDocument source, WordFileFormat format, string marker, bool white) {
        using WordDocument loaded = WordDocument.Load(new MemoryStream(source.ToBytes(new WordSaveOptions { Format = format })));
        byte[] bytes = loaded.ToPdf(new WordPdfExportOptions { IncludePageNumbers = false }).ToBytes();
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
}
