using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests;

public sealed class WordRunTableBackgroundTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void WrappedTableRowsAndCellsRetainInheritedRunPaint(bool highlight) {
        using WordDocument document = CreateTableBackground("first-row", highlight);
        var options = new WordToPdfOptions {
            IncludePageNumbers = false, ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic()
        };
        using (PdfPigDocument plain = PdfPigDocument.Open(document.ToPdfBytes(options)))
            Assert.Equal((0D, 0D, 1D), PaintAt(plain.GetPage(1), "BACKGROUND"));

        Table table = document.Tables.Single()._table;
        TableRow row = table.Elements<TableRow>().Single();
        TableCell cell = row.Elements<TableCell>().Single();
        row.ReplaceChild(new SdtCell(new SdtContentCell((TableCell)cell.CloneNode(true))), cell);
        table.ReplaceChild(new SdtRow(new SdtContentRow((TableRow)row.CloneNode(true))), row);
        using WordDocument loaded = WordDocument.Load(new MemoryStream(document.ToBytes()));
        using PdfPigDocument wrapped = PdfPigDocument.Open(loaded.ToPdfBytes(options));
        Assert.Equal((0D, 0D, 1D), PaintAt(wrapped.GetPage(1), "BACKGROUND"));
        Assert.Single(loaded.Tables.Single().Rows);
        Assert.Single(loaded.Tables.Single().Rows[0].Cells);
    }

    [Theory]
    [InlineData("whole", false)]
    [InlineData("first-row", false)]
    [InlineData("inherited", false)]
    [InlineData("default", false)]
    [InlineData("nil", false)]
    [InlineData("clear", false)]
    [InlineData("first-row-nil", false)]
    [InlineData("whole", true)]
    [InlineData("first-row", true)]
    [InlineData("none", true)]
    [InlineData("first-row-none", true)]
    public void TableRunBackgroundOverridesDocumentDefaultsAndPaintsAfterReopen(string source, bool highlight) {
        using WordDocument document = CreateTableBackground(source, highlight);
        using WordDocument loaded = WordDocument.Load(new MemoryStream(document.ToBytes()));
        Table table = loaded.MainDocumentPartRoot.Document.Body!.Elements<Table>().Single();
        Paragraph paragraph = table.Descendants<Paragraph>().Single();
        var run = new WordParagraph(loaded, paragraph, paragraph.Elements<Run>().Single());
        string tableBefore = table.OuterXml;
        string stylesBefore = loaded.MainDocumentPartRoot.StyleDefinitionsPart!.Styles!.OuterXml;
        bool reset = source is "nil" or "clear" or "none" or "first-row-nil" or "first-row-none";

        if (highlight) Assert.Equal(reset ? HighlightColorValues.None : HighlightColorValues.Blue,
            WordRunFormatting.ResolveRunHighlight(run));
        else Assert.Equal(reset ? null : "0000FF", WordRunFormatting.ResolveRunShadingFillColorHex(run));
        Assert.Equal(tableBefore, table.OuterXml);
        Assert.Equal(stylesBefore, loaded.MainDocumentPartRoot.StyleDefinitionsPart.Styles.OuterXml);

        using PdfPigDocument pdf = PdfPigDocument.Open(loaded.ToPdfBytes(new WordToPdfOptions {
            IncludePageNumbers = false, ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic()
        }));
        var page = pdf.GetPage(1);
        Assert.Contains("BACKGROUND", page.Text, StringComparison.Ordinal);
        var firstLetter = page.Letters.First(letter => letter.Value == "B");
        var backgrounds = page.Paths.Where(path => path.IsFilled && path.GetBoundingRectangle() is { } bounds &&
            bounds.Left <= firstLetter.StartBaseLine.X && bounds.Right >= firstLetter.StartBaseLine.X &&
            bounds.Bottom <= firstLetter.StartBaseLine.Y && bounds.Top >= firstLetter.StartBaseLine.Y).ToArray();
        Assert.Equal(!reset, backgrounds.Any(path => path.FillColor.ToRGBValues() == (0D, 0D, 1D)));
        Assert.DoesNotContain(backgrounds, path => path.FillColor.ToRGBValues() == (1D, 0D, 0D) ||
            path.FillColor.ToRGBValues() == (1D, 1D, 0D));
    }

    [Theory]
    [InlineData("direct")]
    [InlineData("character")]
    [InlineData("paragraph")]
    public void HigherPriorityRunSourcesOverrideTableBackground(string source) {
        using WordDocument document = CreateTableBackground("first-row", highlight: false);
        WordParagraph paragraph = document.Tables[0].Rows[0].Cells[0].Paragraphs[0];
        if (source == "direct") paragraph.RunShadingFillColorHex = "00FF00";
        else {
            document.MainDocumentPartRoot.StyleDefinitionsPart!.Styles!.Append(new Style(new StyleRunProperties(Shade("00FF00"))) {
                StyleId = "HigherBackground", Type = source == "character" ? StyleValues.Character : StyleValues.Paragraph
            });
            if (source == "character") paragraph.SetCharacterStyleId("HigherBackground");
            else paragraph.SetStyleId("HigherBackground");
        }

        using WordDocument loaded = WordDocument.Load(new MemoryStream(document.ToBytes()));
        Paragraph raw = loaded.MainDocumentPartRoot.Document.Body!.Descendants<TableCell>().Single().Elements<Paragraph>().Single();
        Assert.Equal("00FF00", WordRunFormatting.ResolveRunShadingFillColorHex(
            new WordParagraph(loaded, raw, raw.Elements<Run>().Single())));
    }

    [Theory]
    [InlineData("firstRow", "0620", 0, 1)]
    [InlineData("lastRow", "0640", 2, 1)]
    [InlineData("firstCol", "0680", 1, 0)]
    [InlineData("lastCol", "0700", 1, 2)]
    [InlineData("band1Horz", "0400", 0, 0)]
    [InlineData("band2Horz", "0400", 1, 0)]
    [InlineData("band1Vert", "0200", 0, 0)]
    [InlineData("band2Vert", "0200", 0, 1)]
    public void SupportedConditionsUseTheContainingCellAndLeaveOtherCellsAtDefaults(string condition, string look, int row, int column) {
        using WordDocument document = WordDocument.Create();
        Styles styles = document.MainDocumentPartRoot.StyleDefinitionsPart!.Styles!;
        styles.DocDefaults = new DocDefaults(new RunPropertiesDefault(new RunPropertiesBaseStyle(Shade("FF0000"))));
        styles.Append(new Style(new TableStyleProperties(new RunPropertiesBaseStyle(Shade("0000FF"))) {
            Type = new TableStyleOverrideValues(condition)
        }) { Type = StyleValues.Table, StyleId = "ConditionalBackground" });
        WordTable table = document.AddTable(3, 3);
        table._tableProperties!.TableStyle = new TableStyle { Val = "ConditionalBackground" };
        table._tableProperties.TableLook = new TableLook { Val = look };
        for (int r = 0; r < 3; r++)
            for (int c = 0; c < 3; c++) table.Rows[r].Cells[c].Paragraphs[0].Text = $"Cell{r}{c}";

        using WordDocument loaded = WordDocument.Load(new MemoryStream(document.ToBytes()));
        Table raw = loaded.MainDocumentPartRoot.Document.Body!.Elements<Table>().Single();
        Assert.Equal("0000FF", ReadShading(loaded, raw.Elements<TableRow>().ElementAt(row).Elements<TableCell>().ElementAt(column)));
        int otherRow = condition.EndsWith("Horz", StringComparison.Ordinal) ? 1 - row
            : condition.EndsWith("Vert", StringComparison.Ordinal) ? row : 2 - row;
        int otherColumn = condition.EndsWith("Vert", StringComparison.Ordinal) ? 1 - column
            : condition.EndsWith("Horz", StringComparison.Ordinal) ? column : 2 - column;
        Assert.Equal("FF0000", ReadShading(loaded, raw.Elements<TableRow>().ElementAt(otherRow).Elements<TableCell>().ElementAt(otherColumn)));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SecondBandRunBackgroundSurvivesSaveAndNativePdfConversion(bool vertical) {
        using WordDocument document = WordDocument.Create();
        Styles styles = document.MainDocumentPartRoot.StyleDefinitionsPart!.Styles!;
        styles.DocDefaults = new DocDefaults(new RunPropertiesDefault(new RunPropertiesBaseStyle(Shade("FFFF00"))));
        styles.Append(new Style(new TableStyleProperties(new RunPropertiesBaseStyle(Shade("0000FF"))) {
            Type = vertical ? TableStyleOverrideValues.Band2Vertical : TableStyleOverrideValues.Band2Horizontal
        }) { Type = StyleValues.Table, StyleId = "SecondBandBackground" });
        WordTable table = document.AddTable(vertical ? 1 : 2, vertical ? 2 : 1);
        table._tableProperties!.TableStyle = new TableStyle { Val = "SecondBandBackground" };
        table._tableProperties.TableLook = new TableLook { Val = vertical ? "0200" : "0400" };
        table.Rows[0].Cells[0].Paragraphs[0].Text = "FIRST";
        table.Rows[vertical ? 0 : 1].Cells[vertical ? 1 : 0].Paragraphs[0].Text = "SECOND";
        using WordDocument loaded = WordDocument.Load(new MemoryStream(document.ToBytes()));
        using var pdf = PdfPigDocument.Open(loaded.ToPdfBytes(new WordToPdfOptions {
            IncludePageNumbers = false, ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic()
        }));
        Assert.Equal((1D, 1D, 0D), PaintAt(pdf.GetPage(1), "FIRST"));
        Assert.Equal((0D, 0D, 1D), PaintAt(pdf.GetPage(1), "SECOND"));
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void BandSizesAndFirstEdgeRestartTheBodyPattern(bool vertical, bool firstRow) {
        using WordDocument document = WordDocument.Create();
        Styles styles = document.MainDocumentPartRoot.StyleDefinitionsPart!.Styles!;
        styles.Append(new Style(
            new StyleTableProperties(new TableStyleRowBandSize { Val = 2 }, new TableStyleColumnBandSize { Val = 2 }),
            new TableStyleProperties(new RunPropertiesBaseStyle(Shade("FF0000"))) {
                Type = vertical ? TableStyleOverrideValues.Band1Vertical : TableStyleOverrideValues.Band1Horizontal
            },
            new TableStyleProperties(new RunPropertiesBaseStyle(Shade("0000FF"))) {
                Type = vertical ? TableStyleOverrideValues.Band2Vertical : TableStyleOverrideValues.Band2Horizontal
            },
            new TableStyleProperties(new RunPropertiesBaseStyle(Shade("00FF00"))) {
                Type = TableStyleOverrideValues.FirstRow
            },
            new TableStyleProperties(new RunPropertiesBaseStyle(Shade("00FF00"))) {
                Type = TableStyleOverrideValues.FirstColumn
            }) { Type = StyleValues.Table, StyleId = "SizedBands" });
        WordTable table = document.AddTable(vertical ? 1 : 5, vertical ? 5 : 1);
        table._tableProperties!.TableStyle = new TableStyle { Val = "SizedBands" };
        table._tableProperties.TableLook = new TableLook { Val = firstRow ? vertical ? "0280" : "0420" : vertical ? "0200" : "0400" };
        foreach (WordTableRow row in table.Rows)
            foreach (WordTableCell cell in row.Cells) cell.Paragraphs[0].Text = "PAINT";
        using WordDocument loaded = WordDocument.Load(new MemoryStream(document.ToBytes()));
        Table raw = loaded.MainDocumentPartRoot.Document.Body!.Elements<Table>().Single();
        for (int index = 0; index < 5; index++) {
            int bodyIndex = index - (firstRow ? 1 : 0);
            string expected = bodyIndex < 0 ? "00FF00" : bodyIndex / 2 % 2 == 0 ? "FF0000" : "0000FF";
            Assert.Equal(expected, ReadShading(loaded, raw.Elements<TableRow>().ElementAt(vertical ? 0 : index)
                .Elements<TableCell>().ElementAt(vertical ? index : 0)));
        }
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void ExplicitZeroBandSizeStopsInheritedBanding(bool vertical, bool inherited) {
        using WordDocument document = WordDocument.Create();
        Styles styles = document.MainDocumentPartRoot.StyleDefinitionsPart!.Styles!;
        var band = new TableStyleProperties(new RunPropertiesBaseStyle(Shade("0000FF"))) {
            Type = vertical ? TableStyleOverrideValues.Band1Vertical : TableStyleOverrideValues.Band1Horizontal
        };
        var parent = new Style(new StyleTableProperties(new TableStyleRowBandSize { Val = 2 },
            new TableStyleColumnBandSize { Val = 2 }), band) { Type = StyleValues.Table, StyleId = "BandParent" };
        styles.Append(parent);
        var current = new Style(new StyleRunProperties(Shade("FFFF00")),
            new StyleTableProperties(vertical ? new TableStyleColumnBandSize { Val = 0 } : new TableStyleRowBandSize { Val = 0 })) {
            Type = StyleValues.Table, StyleId = "ZeroBand"
        };
        if (inherited) current.BasedOn = new BasedOn { Val = "BandParent" };
        else current.Append(band.CloneNode(true));
        styles.Append(current);
        WordTable table = document.AddTable(vertical ? 1 : 3, vertical ? 3 : 1);
        table._tableProperties!.TableStyle = new TableStyle { Val = "ZeroBand" };
        table._tableProperties.TableLook = new TableLook { Val = vertical ? "0200" : "0400" };
        foreach (WordTableRow row in table.Rows)
            foreach (WordTableCell cell in row.Cells) cell.Paragraphs[0].Text = "NO BAND";
        using WordDocument loaded = WordDocument.Load(new MemoryStream(document.ToBytes()));
        Table raw = loaded.MainDocumentPartRoot.Document.Body!.Elements<Table>().Single();
        Assert.All(raw.Descendants<TableCell>(), cell => Assert.Equal("FFFF00", ReadShading(loaded, cell)));
    }

    [Fact]
    public void ExpandedLookFalseOverridesHexAndDisablesInheritedCondition() {
        using WordDocument document = CreateTableBackground("inherited", highlight: false);
        Table raw = document.Tables[0]._table;
        raw.GetFirstChild<TableProperties>()!.TableLook = new TableLook { Val = "0620", FirstRow = false };
        using WordDocument loaded = WordDocument.Load(new MemoryStream(document.ToBytes()));
        Table reopened = loaded.MainDocumentPartRoot.Document.Body!.Elements<Table>().Single();
        // The derived whole-table green is visible when the inherited first-row blue is disabled.
        Assert.Equal("00FF00", ReadShading(loaded, reopened.Descendants<TableCell>().Single()));
    }

    [Fact]
    public void LogicalLastColumnUsesGridBeforeAndMergedSpan() {
        using WordDocument document = WordDocument.Create();
        Styles styles = document.MainDocumentPartRoot.StyleDefinitionsPart!.Styles!;
        styles.DocDefaults = new DocDefaults(new RunPropertiesDefault(new RunPropertiesBaseStyle(Shade("FF0000"))));
        styles.Append(new Style(new TableStyleProperties(new RunPropertiesBaseStyle(Shade("0000FF"))) {
            Type = TableStyleOverrideValues.LastColumn
        }) { Type = StyleValues.Table, StyleId = "MergedBackground" });
        document.MainDocumentPartRoot.Document.Body!.Append(new Table(
            new TableProperties(new TableStyle { Val = "MergedBackground" }, new TableLook { Val = "0700" }),
            new TableGrid(new GridColumn(), new GridColumn(), new GridColumn()),
            new TableRow(new TableRowProperties(new GridBefore { Val = 1 }),
                new TableCell(new TableCellProperties(new GridSpan { Val = 2 }), new Paragraph(new Run(new Text("BACKGROUND")))))));

        using WordDocument loaded = WordDocument.Load(new MemoryStream(document.ToBytes()));
        Assert.Equal("0000FF", ReadShading(loaded, loaded.MainDocumentPartRoot.Document.Body!.Descendants<TableCell>().Single()));
    }

    [Theory]
    [InlineData("TableNormal")]
    [InlineData("NormalTable")]
    public void BuiltInNormalTablePropertiesAreIgnoredEvenAsAnAncestor(string builtInId) {
        using WordDocument document = CreateTableBackground("whole", highlight: false);
        Styles styles = document.MainDocumentPartRoot.StyleDefinitionsPart!.Styles!;
        foreach (Style existing in styles.Elements<Style>().Where(style => style.StyleId?.Value == builtInId).ToArray()) existing.Remove();
        styles.Append(new Style(new StyleRunProperties(Shade("0000FF"))) { Type = StyleValues.Table, StyleId = builtInId });
        Style custom = styles.Elements<Style>().Single(style => style.StyleId == "TableBackground");
        custom.StyleRunProperties = null;
        custom.BasedOn = new BasedOn { Val = builtInId };

        using WordDocument loaded = WordDocument.Load(new MemoryStream(document.ToBytes()));
        Assert.Equal("FF0000", ReadShading(loaded, loaded.MainDocumentPartRoot.Document.Body!.Descendants<TableCell>().Single()));
    }


    [Fact]
    public void NativePdfRefreshesConditionalTablePositionsBetweenConversions() {
        using WordDocument document = CreateTableBackground("first-row", highlight: true);
        WordTable table = document.Tables[0];
        table.AddRow(1).Cells[0].Paragraphs[0].Text = "SECOND";
        var options = new WordToPdfOptions { IncludePageNumbers = false, ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic() };
        using (var before = PdfPigDocument.Open(document.ToPdfBytes(options))) {
            Assert.Equal((0D, 0D, 1D), PaintAt(before.GetPage(1), "BACKGROUND"));
            Assert.Equal((1D, 1D, 0D), PaintAt(before.GetPage(1), "SECOND"));
        }
        TableRow first = table._table.Elements<TableRow>().First();
        first.Remove(); table._table.Append(first);
        using var after = PdfPigDocument.Open(document.ToPdfBytes(options));
        Assert.Equal((1D, 1D, 0D), PaintAt(after.GetPage(1), "BACKGROUND"));
        Assert.Equal((0D, 0D, 1D), PaintAt(after.GetPage(1), "SECOND"));
    }

    private static (double, double, double) PaintAt(UglyToad.PdfPig.Content.Page page, string text) {
        var letter = page.GetWords().Single(word => word.Text == text).Letters[0];
        var path = page.Paths.Last(path => path.IsFilled &&
            (path.FillColor.ToRGBValues() == (0D, 0D, 1D) || path.FillColor.ToRGBValues() == (1D, 1D, 0D)) &&
            path.GetBoundingRectangle() is { } bounds &&
            bounds.Left <= letter.StartBaseLine.X && bounds.Right >= letter.StartBaseLine.X &&
            bounds.Bottom <= letter.StartBaseLine.Y && bounds.Top >= letter.StartBaseLine.Y);
        return path.FillColor.ToRGBValues();
    }

    private static string? ReadShading(WordDocument document, TableCell cell) {
        Paragraph paragraph = cell.Elements<Paragraph>().Single();
        return WordRunFormatting.ResolveRunShadingFillColorHex(new WordParagraph(document, paragraph, paragraph.Elements<Run>().Single()));
    }

    private static WordDocument CreateTableBackground(string source, bool highlight) {
        WordDocument document = WordDocument.Create();
        Styles styles = document.MainDocumentPartRoot.StyleDefinitionsPart!.Styles!;
        styles.DocDefaults = new DocDefaults(new RunPropertiesDefault(highlight
            ? new RunPropertiesBaseStyle(new Highlight { Val = HighlightColorValues.Yellow })
            : new RunPropertiesBaseStyle(Shade("FF0000"))));
        Shading shading = source.EndsWith("nil", StringComparison.Ordinal) ? new Shading { Val = ShadingPatternValues.Nil, Fill = "0000FF" }
            : Shade(source == "clear" ? "auto" : "0000FF");
        var properties = new RunPropertiesBaseStyle();
        if (highlight) properties.Append(new Highlight { Val = source.EndsWith("none", StringComparison.Ordinal) ? HighlightColorValues.None : HighlightColorValues.Blue });
        else properties.Append(shading);
        var style = new Style { Type = StyleValues.Table, StyleId = "TableBackground" };
        bool conditional = source.StartsWith("first-row", StringComparison.Ordinal) || source == "inherited";
        if (conditional)
            style.Append(new TableStyleProperties(properties) { Type = TableStyleOverrideValues.FirstRow });
        else style.StyleRunProperties = new StyleRunProperties(properties.ChildElements.Select(child => child.CloneNode(true)));
        styles.Append(style);
        string styleId = "TableBackground";
        if (source == "inherited") {
            styles.Append(new Style(new BasedOn { Val = styleId }, new StyleRunProperties(Shade("00FF00"))) {
                Type = StyleValues.Table, StyleId = "DerivedTableBackground"
            });
            styleId = "DerivedTableBackground";
        }
        WordTable table = document.AddTable(1, 1);
        table._tableProperties!.TableStyle = new TableStyle { Val = styleId };
        table._tableProperties.TableLook = new TableLook { Val = "0600", FirstRow = conditional };
        if (source == "default") {
            foreach (Style other in styles.Elements<Style>().Where(candidate => candidate.Type?.Value == StyleValues.Table)) other.Default = false;
            style.Default = true;
            table._tableProperties.TableStyle = null;
        }
        WordParagraph paragraph = table.Rows[0].Cells[0].Paragraphs[0];
        paragraph.Text = "BACKGROUND";
        paragraph.FontFamily = "Arial";
        paragraph.FontSize = 12;
        return document;
    }

    private static Shading Shade(string fill) => new Shading { Val = ShadingPatternValues.Clear, Fill = fill };
}
