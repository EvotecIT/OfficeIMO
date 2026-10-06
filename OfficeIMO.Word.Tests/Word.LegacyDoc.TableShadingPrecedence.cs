using OfficeIMO.Word;
using Xunit;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData("child")]
    [InlineData("grandchild")]
    [InlineData("direct")]
    [InlineData("conditional")]
    [InlineData("conditional-child")]
    public void LegacyDoc_TableShading_AutomaticCellClearsStopLowerPriorityFill(string kind) {
        using WordDocument source = WordDocument.Create();
        bool child = kind is "child" or "grandchild" or "conditional-child";
        WordTable table = CreateTableShadingPrecedenceFixture(source, child);
        table.StyleDetails!.ShadingFillColorHex = "auto";
        table.StyleDetails.CellSpacing = 120;
        W.Styles styles = source._wordprocessingDocument!.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        W.Style grid = styles.Elements<W.Style>().Single(s => s.StyleId?.Value == "TableGrid");
        W.Style own = child ? styles.Elements<W.Style>().Single(s => s.StyleId?.Value == "ShadingChild") : grid;
        if (kind is "child" or "grandchild") {
            own.AddChild(new W.StyleTableCellProperties(new W.Shading {
                Val = W.ShadingPatternValues.Clear, Fill = "auto" }), true);
            if (kind == "grandchild") {
                styles.Append(new W.Style(new W.StyleName { Val = "Shading grandchild" }, new W.BasedOn { Val = "ShadingChild" }) {
                    Type = W.StyleValues.Table, StyleId = "ShadingGrandchild", CustomStyle = true
                });
                table._tableProperties!.TableStyle = new W.TableStyle { Val = "ShadingGrandchild" };
            }
        } else if (kind == "direct") {
            WordTableCell cell = table.Rows[0].Cells[0];
            cell.ShadingFillColorHex = "auto";
            cell.ShadingFillColorHex = cell.ShadingFillColorHex;
        }
        else {
            if (kind == "conditional-child") grid.Append(new W.TableStyleProperties(
                new W.TableStyleConditionalFormattingTableCellProperties(new W.Shading {
                    Val = W.ShadingPatternValues.Clear, Fill = "00FF00" })) { Type = W.TableStyleOverrideValues.FirstRow });
            own.Append(new W.TableStyleProperties(new W.TableStyleConditionalFormattingTableCellProperties(
                new W.Shading { Val = W.ShadingPatternValues.Clear, Fill = "auto" })) { Type = W.TableStyleOverrideValues.FirstRow });
        }
        Assert.Empty(source.DocumentValidationErrors);
        using WordDocument restored = WordDocument.Load(new MemoryStream(source.ToBytes(WordFileFormat.Doc)));
        for (int row = 0; row < 2; row++) for (int column = 0; column < 2; column++) {
            bool cleared = kind is "child" or "grandchild" || (row == 0 && (kind != "direct" || column == 0));
            Assert.Equal(cleared ? string.Empty : "FFFF00", restored.Tables[0].Rows[row].Cells[column].ShadingFillColorHex);
        }
        Assert.Empty(restored.DocumentValidationErrors);
    }

    [Theory]
    [InlineData("auto", false)]
    [InlineData("auto", true)]
    [InlineData("foreground", false)]
    [InlineData("foreground", true)]
    [InlineData("red", false)]
    [InlineData("red", true)]
    public void LegacyDoc_TableShading_CellDefaultsSurviveDirectTableGapFormatting(string direct, bool inherited) {
        using WordDocument source = WordDocument.Create();
        WordTable table = CreateTableShadingPrecedenceFixture(source, inherited);
        if (direct == "foreground") table.StyleDetails!.ShadingColorHex = "FF0000";
        else table.StyleDetails!.ShadingFillColorHex = direct == "auto" ? "auto" : "FF0000";
        table.StyleDetails!.CellSpacing = direct == "red" ? (short)0 : (short)120;
        using WordDocument restored = WordDocument.Load(new MemoryStream(source.ToBytes(WordFileFormat.Doc)));
        Assert.All(restored.Tables[0].Rows.SelectMany(row => row.Cells), cell => Assert.Equal("FFFF00", cell.ShadingFillColorHex));
        Assert.Equal(direct == "red" ? (short)0 : (short)120, restored.Tables[0].StyleDetails!.CellSpacing);
        Assert.Empty(restored.DocumentValidationErrors);
    }

    [Fact]
    public void LegacyDoc_TableShading_ZeroSpacingDoesNotPromoteTableFillIntoCells() {
        using WordDocument source = WordDocument.Create();
        WordTable table = source.AddTable(1, 1, WordTableStyle.TableGrid);
        table.StyleDetails!.CellSpacing = 0;
        table.StyleDetails.ShadingFillColorHex = "FF0000";
        using WordDocument restored = WordDocument.Load(new MemoryStream(source.ToBytes(WordFileFormat.Doc)));
        Assert.Equal(string.Empty, restored.Tables[0].Rows[0].Cells[0].ShadingFillColorHex);
    }

    [Fact]
    public void LegacyDoc_TableShading_DirectAndConditionalCellColorsRetainPriority() {
        using WordDocument source = WordDocument.Create();
        WordTable table = CreateTableShadingPrecedenceFixture(source, false);
        table.StyleDetails!.CellSpacing = 120;
        table.StyleDetails.ShadingFillColorHex = "auto";
        table.Rows[0].Cells[0].ShadingFillColorHex = "FFFFFF";
        W.Style grid = source._wordprocessingDocument!.MainDocumentPart!.StyleDefinitionsPart!.Styles!
            .Elements<W.Style>().Single(s => s.StyleId?.Value == "TableGrid");
        grid.Append(new W.TableStyleProperties(new W.TableStyleConditionalFormattingTableCellProperties(
            new W.Shading { Val = W.ShadingPatternValues.Clear, Fill = "00FF00" })) { Type = W.TableStyleOverrideValues.FirstRow });
        using WordDocument restored = WordDocument.Load(new MemoryStream(source.ToBytes(WordFileFormat.Doc)));
        Assert.Equal("FFFFFF", restored.Tables[0].Rows[0].Cells[0].ShadingFillColorHex);
        Assert.Equal("00FF00", restored.Tables[0].Rows[0].Cells[1].ShadingFillColorHex);
        Assert.All(restored.Tables[0].Rows[1].Cells, cell => Assert.Equal("FFFF00", cell.ShadingFillColorHex));
    }

    private static WordTable CreateTableShadingPrecedenceFixture(WordDocument document, bool inherited) {
        WordTable table = document.AddTable(2, 2, WordTableStyle.TableGrid);
        W.Styles styles = document._wordprocessingDocument!.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        W.Style grid = styles.Elements<W.Style>().Single(s => s.StyleId?.Value == "TableGrid");
        grid.GetFirstChild<W.StyleTableProperties>()!.AddChild(new W.Shading { Val = W.ShadingPatternValues.Clear, Fill = "0000FF" }, true);
        var cellProperties = grid.GetFirstChild<W.StyleTableCellProperties>();
        if (cellProperties == null) { cellProperties = new W.StyleTableCellProperties(); grid.AddChild(cellProperties, true); }
        cellProperties.AddChild(new W.Shading { Val = W.ShadingPatternValues.Clear, Fill = "FFFF00" }, true);
        if (inherited) {
            styles.Append(new W.Style(new W.StyleName { Val = "Shading child" }, new W.BasedOn { Val = "TableGrid" }) {
                Type = W.StyleValues.Table, StyleId = "ShadingChild", CustomStyle = true
            });
            table._tableProperties!.TableStyle = new W.TableStyle { Val = "ShadingChild" };
        }
        return table;
    }
}
