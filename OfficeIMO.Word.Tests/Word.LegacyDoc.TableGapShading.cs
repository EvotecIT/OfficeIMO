using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData("style")]
    [InlineData("inherited")]
    [InlineData("direct-spacing")]
    [InlineData("direct-shading")]
    public void LegacyDoc_TableGapShading_RejectsLossBeforeCreatingFile(string mode) {
        using WordDocument document = WordDocument.Create();
        WordTable table = CreateNativeGapShadingTable(document, mode);
        table.Rows[0].Cells[0].Paragraphs[0].Text = "Visible gaps";
        string output = Path.Combine(Path.GetTempPath(), "OfficeIMO-gap-" + Guid.NewGuid().ToString("N") + ".doc");
        try {
            var error = Assert.Throws<NotSupportedException>(() => document.Save(output));
            Assert.Contains("table gap shading", error.Message);
            Assert.False(File.Exists(output));
        } finally {
            if (File.Exists(output)) File.Delete(output);
        }
    }

    [Theory]
    [InlineData("zero-spacing")]
    [InlineData("zero-style-spacing")]
    [InlineData("clear-style")]
    public void LegacyDoc_TableGapShading_EffectiveOverridesPermitSupportedSave(string mode) {
        using WordDocument document = WordDocument.Create();
        WordTable table = CreateNativeGapShadingTable(document, "inherited");
        if (mode == "zero-spacing") table.StyleDetails!.CellSpacing = 0;
        else {
            Styles styles = document._wordprocessingDocument!.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
            var properties = mode == "zero-style-spacing"
                ? new StyleTableProperties(new TableCellSpacing { Width = "0", Type = TableWidthUnitValues.Dxa })
                : new StyleTableProperties(new Shading { Val = ShadingPatternValues.Clear, Fill = "auto" });
            styles.Elements<Style>().Single(s => s.StyleId?.Value == "GapChild").Append(properties);
        }
        table.Rows[0].Cells[0].Paragraphs[0].Text = "Supported cells";
        using WordDocument restored = WordDocument.Load(new MemoryStream(document.ToBytes(WordFileFormat.Doc)));
        Assert.Equal("Supported cells", restored.Tables[0].Rows[0].Cells[0].Paragraphs[0].Text);
        Assert.Equal(string.Empty, restored.Tables[0].Rows[0].Cells[0].ShadingFillColorHex);
        Assert.Equal(mode.Contains("spacing") ? (short)0 : (short)120, restored.Tables[0].StyleDetails!.CellSpacing);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void LegacyDoc_TableGapShading_ConditionalRegionsRespectTableLook(bool active) {
        using WordDocument document = WordDocument.Create();
        WordTable table = document.AddTable(1, 1, WordTableStyle.TableGrid);
        table.Rows[0].Cells[0].Paragraphs[0].Text = "Conditional gap";
        table.StyleDetails!.CellSpacing = 120;
        table._tableProperties!.TableLook = new TableLook { FirstRow = active };
        Styles styles = document._wordprocessingDocument!.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        Style grid = styles.Elements<Style>().Single(s => s.StyleId?.Value == "TableGrid");
        grid.Append(new TableStyleProperties(new TableStyleConditionalFormattingTableProperties(
            new Shading { Val = ShadingPatternValues.Clear, Fill = "FF0000" })) { Type = TableStyleOverrideValues.FirstRow });
        if (active) Assert.Throws<NotSupportedException>(() => document.ToBytes(WordFileFormat.Doc));
        else {
            using WordDocument restored = WordDocument.Load(new MemoryStream(document.ToBytes(WordFileFormat.Doc)));
            Assert.Equal("Conditional gap", restored.Tables[0].Rows[0].Cells[0].Paragraphs[0].Text);
        }
    }

    private static WordTable CreateNativeGapShadingTable(WordDocument document, string mode) {
        WordTable table = document.AddTable(1, 1, WordTableStyle.TableGrid);
        Styles styles = document._wordprocessingDocument!.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        Style grid = styles.Elements<Style>().Single(s => s.StyleId?.Value == "TableGrid");
        StyleTableProperties properties = grid.GetFirstChild<StyleTableProperties>()!;
        if (mode == "direct-shading") table._tableProperties!.Append(
            new Shading { Val = ShadingPatternValues.Clear, Fill = "FF0000" });
        else properties.Shading = new Shading { Val = ShadingPatternValues.Clear, Fill = "FF0000" };
        if (mode is "direct-spacing" or "direct-shading") table.StyleDetails!.CellSpacing = 120;
        else properties.Append(new TableCellSpacing { Width = "120", Type = TableWidthUnitValues.Dxa });
        if (mode == "inherited") {
            styles.Append(new Style(new StyleName { Val = "Gap Child" }, new BasedOn { Val = "TableGrid" }) {
                Type = StyleValues.Table, StyleId = "GapChild", CustomStyle = true
            });
            table._tableProperties!.TableStyle = new TableStyle { Val = "GapChild" };
        }
        return table;
    }
}
