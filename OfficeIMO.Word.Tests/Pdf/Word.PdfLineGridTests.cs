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
    [InlineData(11, WordLineSpacingRule.Auto, 240, 18D)]
    [InlineData(20, WordLineSpacingRule.Auto, 240, 36D)]
    [InlineData(11, WordLineSpacingRule.Auto, 360, 27D)]
    [InlineData(20, WordLineSpacingRule.Auto, 360, 36D)]
    [InlineData(11, WordLineSpacingRule.AtLeast, 600, 30D)]
    [InlineData(20, WordLineSpacingRule.AtLeast, 600, 36D)]
    [InlineData(11, WordLineSpacingRule.Exact, 240, 12D)]
    public void SaveAsPdf_LineGridUsesNaturalLineBoxesAndAuthoredSpacing(int size, WordLineSpacingRule rule, int line, double expected) {
        double[] baselines = ExportGridBaselines(size, rule, line, "body");
        Assert.Equal(expected, baselines[0] - baselines[1], 3);
        Assert.Equal(expected, baselines[1] - baselines[2], 3);
    }

    [Theory]
    [InlineData("disabled", 16.5D)]
    [InlineData("style-disabled", 16.5D)]
    [InlineData("direct-enabled", 18D)]
    [InlineData("table", 16.5D)]
    [InlineData("table-adjust", 18D)]
    [InlineData("columns", 18D)]
    [InlineData("table-style-disabled", 16.5D)]
    [InlineData("table-direct-enabled", 18D)]
    [InlineData("table-conditional-disabled", 16.5D)]
    public void SaveAsPdf_LineGridRespectsParagraphOverridesAndTableCompatibility(string route, double expected) {
        double[] baselines = ExportGridBaselines(11, WordLineSpacingRule.Auto, 240, route);
        Assert.Equal(expected, baselines[0] - baselines[1], 3);
        Assert.Equal(expected, baselines[1] - baselines[2], 3);
    }

    [Fact]
    public void SaveAsPdf_LineGridPositionsTheBaselineInsideTheExpandedLineBox() {
        double baseline = ExportGridBaselines(11, WordLineSpacingRule.Auto, 240, "disabled")[0];
        Assert.Equal(.75D, baseline - ExportGridBaselines(11, WordLineSpacingRule.Auto, 240, "body")[0], 3);
        Assert.Equal(12.75D, baseline - ExportGridBaselines(11, WordLineSpacingRule.AtLeast, 600, "body")[0], 3);
    }

    private static double[] ExportGridBaselines(int size, WordLineSpacingRule rule, int line, string route) {
        using WordDocument source = WordDocument.Create();
        source.Sections[0]._sectionProperties.Append(new DocGrid { Type = DocGridValues.Lines, LinePitch = 360 });
        if (route == "columns") source.Sections[0].ColumnCount = 2;
        WordTable? table = route.StartsWith("table", StringComparison.Ordinal) ? source.AddTable(1, 1) : null;
        WordParagraph paragraph = table?.Rows[0].Cells[0].Paragraphs[0] ?? source.AddParagraph();
        paragraph._paragraph.RemoveAllChildren<Run>();
        paragraph._paragraph.Append(new Run(new RunProperties(new FontSize { Val = (size * 2).ToString() },
            new RunFonts { Ascii = "Arial", HighAnsi = "Arial" }), new Text("A"), new Break(), new Text("B"), new Break(), new Text("C")));
        paragraph.LineSpacing = (rule) switch { global::OfficeIMO.Word.WordLineSpacingRule.Exact => global::OfficeIMO.Word.WordLineSpacing.Exact(global::OfficeIMO.OfficeLength.Emus(checked((long)(line) * 635))), global::OfficeIMO.Word.WordLineSpacingRule.AtLeast => global::OfficeIMO.Word.WordLineSpacing.AtLeast(global::OfficeIMO.OfficeLength.Emus(checked((long)(line) * 635))), _ => global::OfficeIMO.Word.WordLineSpacing.Multiple((line) / 240D) };
        paragraph.LineSpacingBefore = global::OfficeIMO.OfficeLength.Zero; paragraph.LineSpacingAfter = global::OfficeIMO.OfficeLength.Zero;
        if (route == "disabled") paragraph._paragraph.ParagraphProperties!.Append(new SnapToGrid { Val = false });
        if (route is "style-disabled" or "direct-enabled") {
            Styles styles = source._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
            Style normal = styles.Elements<Style>().Single(style => style.StyleId?.Value == "Normal");
            (normal.StyleParagraphProperties ??= new StyleParagraphProperties()).Append(new SnapToGrid { Val = false });
            if (route == "direct-enabled") paragraph._paragraph.ParagraphProperties!.Append(new SnapToGrid { Val = true });
        }
        if (route is "table-adjust" or "table-style-disabled" or "table-direct-enabled" or "table-conditional-disabled") source._wordprocessingDocument.MainDocumentPart!.DocumentSettingsPart!.Settings!
            .GetFirstChild<Compatibility>()!.Append(new AdjustLineHeightInTable());
        if (route is "table-style-disabled" or "table-direct-enabled" or "table-conditional-disabled") {
            var gridStyle = new Style {
                Type = StyleValues.Table, StyleId = "GridTest", StyleName = new StyleName { Val = "Grid Test" },
                StyleParagraphProperties = new StyleParagraphProperties(new SnapToGrid { Val = false })
            };
            if (route == "table-conditional-disabled") {
                gridStyle.StyleParagraphProperties = null;
                gridStyle.Append(new TableStyleProperties(new StyleParagraphProperties(new SnapToGrid { Val = false })) {
                    Type = TableStyleOverrideValues.FirstRow
                });
                table!.ConditionalFormattingFirstRow = true;
            }
            source._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!.Append(gridStyle);
            (table!._table.TableProperties ??= new TableProperties()).TableStyle = new TableStyle { Val = "GridTest" };
            if (route == "table-direct-enabled") paragraph._paragraph.ParagraphProperties!.Append(new SnapToGrid { Val = true });
        }
        return ExportGridDocumentBaselines(source);
    }

    private static double[] ExportGridDocumentBaselines(WordDocument source) {
        var options = new PdfCreateOptions { DefaultFont = PdfStandardFont.Helvetica };
        options.EmbedStandardFont(PdfStandardFont.Helvetica,
            ManagedTextShapingTestAssets.CreateFontWithLineBoxMetrics(1200, -300, 0, 300,
                Enumerable.Range(32, 95).ToArray()), "OfficeIMO-Grid-Regular");
        string original = source._wordprocessingDocument.MainDocumentPart!.Document.OuterXml;
        using var pdf = PdfPigDocument.Open(source.ToPdf(new WordPdfExportOptions {
            IncludePageNumbers = false, ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic(), PdfOptions = options
        }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        Assert.Equal(original, source._wordprocessingDocument.MainDocumentPart.Document.OuterXml);
        var letters = pdf.GetPages().SelectMany(page => page.Letters).Where(letter => letter.Value is "A" or "B" or "C").ToArray();
        Assert.Equal("ABC", string.Concat(letters.Select(letter => letter.Value)));
        return letters.Select(letter => letter.StartBaseLine.Y).ToArray();
    }
}
