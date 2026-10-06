using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void LegacyDoc_TableWordCompatibility_CellShadingDefaultsDoNotUseTableGapShading(bool derived) {
        using WordDocument document = WordDocument.Create();
        WordTable table = document.AddTable(1, 1, WordTableStyle.TableGrid);
        Styles styles = document._wordprocessingDocument!.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        Style grid = styles.Elements<Style>().Single(s => s.StyleId?.Value == "TableGrid");
        grid.GetFirstChild<StyleTableProperties>()!.Shading = new Shading { Val = ShadingPatternValues.Clear, Fill = "FF0000" };
        grid.Append(new StyleTableCellProperties(new Shading { Val = ShadingPatternValues.Clear, Fill = "00FF00" }));
        if (derived) {
            styles.Append(new Style(new StyleName { Val = "Cell Shading Child" }, new BasedOn { Val = "TableGrid" }) {
                Type = StyleValues.Table, StyleId = "CellShadingChild", CustomStyle = true
            });
            table._tableProperties!.TableStyle = new TableStyle { Val = "CellShadingChild" };
        }
        table.Rows[0].Cells[0].Paragraphs[0].Text = "Cell fill";
        using WordDocument restored = WordDocument.Load(new MemoryStream(document.ToBytes(WordFileFormat.Doc)));
        Assert.Equal("00FF00", restored.Tables[0].Rows[0].Cells[0].ShadingFillColorHex);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void LegacyDoc_TableWordCompatibility_AbsentAndZeroBandSizesDisableBanding(int mode) {
        using WordDocument document = WordDocument.Create();
        WordTable table = document.AddTable(2, 2, WordTableStyle.TableGrid);
        Styles styles = document._wordprocessingDocument!.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        Style grid = styles.Elements<Style>().Single(s => s.StyleId?.Value == "TableGrid");
        grid.Append(new TableStyleProperties(new StyleRunProperties(new Color { Val = "FF0000" })) {
            Type = TableStyleOverrideValues.Band1Horizontal
        });
        grid.Append(new TableStyleProperties(new StyleRunProperties(new Color { Val = "0000FF" })) {
            Type = TableStyleOverrideValues.Band1Vertical
        });
        if (mode > 0) {
            grid.GetFirstChild<StyleTableProperties>()!.TableStyleRowBandSize = new TableStyleRowBandSize { Val = mode == 1 ? 0 : 1 };
            grid.GetFirstChild<StyleTableProperties>()!.TableStyleColumnBandSize = new TableStyleColumnBandSize { Val = mode == 1 ? 0 : 1 };
        }
        if (mode == 2) {
            styles.Append(new Style(new StyleName { Val = "Disabled Bands" }, new BasedOn { Val = "TableGrid" },
                new StyleTableProperties(new TableStyleRowBandSize { Val = 0 }, new TableStyleColumnBandSize { Val = 0 })) {
                Type = StyleValues.Table, StyleId = "DisabledBands", CustomStyle = true
            });
            table._tableProperties!.TableStyle = new TableStyle { Val = "DisabledBands" };
        }
        table._tableProperties!.TableLook = new TableLook { FirstRow = false, FirstColumn = false, LastRow = false, LastColumn = false, NoHorizontalBand = false, NoVerticalBand = false };
        for (int row = 0; row < 2; row++) for (int col = 0; col < 2; col++) table.Rows[row].Cells[col].Paragraphs[0].Text = "Disabled band";
        using WordDocument restored = WordDocument.Load(new MemoryStream(document.ToBytes(WordFileFormat.Doc)));
        foreach (WordTableRow row in restored.Tables[0].Rows) foreach (WordTableCell cell in row.Cells) Assert.Equal(string.Empty, cell.Paragraphs[0].ColorHex);
    }
}
