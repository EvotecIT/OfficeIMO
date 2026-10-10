using System.Text.Json;
using OfficeIMO.IWork;
using OfficeIMO.Word;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Fact]
    public void Native_Pages_bordered_table_matches_all_saved_and_native_DOCX_cell_edges() {
        string source = CorpusFixture("native-exports/pages-table-borders-v15.4.pages");
        string reference = CorpusFixture("native-exports/pages-table-borders-v15.4.docx");
        using var decoded = JsonDocument.Parse(File.ReadAllText(CorpusFixture("native-exports/pages-table-borders-v15.4-decode.json")));
        var expected = decoded.RootElement;
        Assert.Equal(expected.GetProperty("sourceSha256").GetString(), BorderCorpusHash(source));
        Assert.Equal(expected.GetProperty("nativeDocxSha256").GetString(), BorderCorpusHash(reference));
        using var result = WordIWorkConverter.ConvertPagesToWordResult(source,
            conversionOptions: new IWorkConversionOptions { AllowPartialEditableReconstruction = true });
        Assert.True(result.Report.IsPartialEditableReconstruction);
        Assert.DoesNotContain(result.Report.Diagnostics, d => d.Code == "IWORK_PAGES_DOCX_PRECISION");
        Assert.Contains(result.Report.Diagnostics, d => d.Code == "IWORK_PAGES_TABLE_LAYOUT_APPROXIMATED");
        Assert.DoesNotContain(result.Report.Diagnostics, d => d.Code == "IWORK_TABLE_BORDERS_UNSUPPORTED");
        Assert.Throws<InvalidOperationException>(() => result.Report.RequireCompleteEditableReconstruction());
        Assert.False(result.IsVisualFallback);
        var table = Assert.Single(result.Projection.Tables);
        Assert.Equal(4, table.RowCount); Assert.Equal(4, table.ColumnCount);
        Assert.Equal(1, table.HeaderRowCount); Assert.Equal(1, table.HeaderColumnCount); Assert.Equal(1, table.FooterRowCount);
        Assert.False(table.AutoResizeRows);
        Assert.True(table.HasBorderDeclarations);

        using var saved = new MemoryStream(); result.Value.Save(saved); saved.Position = 0;
        using var reopened = WordDocument.Load(saved);
        using var native = WordDocument.Load(reference);
        var output = Assert.Single(reopened.Tables);
        var nativeTable = Assert.Single(native.Tables);
        Assert.Equal(4, output.Rows.Count);
        int edges = 0;
        foreach (var cell in expected.GetProperty("cells").EnumerateArray()) {
            int row = cell.GetProperty("row").GetInt32(), column = cell.GetProperty("column").GetInt32();
            string value = cell.GetProperty("value").GetString()!;
            Assert.Equal(value, table.GetCell(row, column)?.DisplayText ?? string.Empty);
            var target = output.Rows[row - 1].Cells[column - 1];
            var apple = nativeTable.Rows[row - 1].Cells[column - 1];
            Assert.Equal(value, string.Join("", target.Paragraphs.Select(p => p.Text)));
            Assert.Equal(value, string.Join("", apple.Paragraphs.Select(p => p.Text)));
            var borders = table.GetBorders(row, column);
            foreach (var edge in cell.GetProperty("edges").EnumerateObject()) {
                var side = edge.Name switch {
                    "left" => borders.Left, "right" => borders.Right, "top" => borders.Top, "bottom" => borders.Bottom,
                    _ => throw new InvalidDataException("Unknown reference edge: " + edge.Name)
                };
                double width = edge.Value.GetProperty("widthPoints").GetDouble();
                string rgb = edge.Value.GetProperty("rgb").GetString()!;
                AssertBorder(side, rgb, width);
                var actual = PagesBorderSide(target.Borders, edge.Name);
                var exported = PagesBorderSide(apple.Borders, edge.Name);
                Assert.Equal(width == 0 ? WordBorderStyle.Nil : WordBorderStyle.Single, actual.Style);
                Assert.Equal(exported.Style, actual.Style);
                if (width > 0) {
                    Assert.Equal(checked((uint)(width * 8)), actual.Size);
                    Assert.Equal(exported.Size, actual.Size);
                    Assert.Equal(rgb, actual.Color, ignoreCase: true);
                    Assert.Equal(exported.Color, actual.Color, ignoreCase: true);
                }
                edges++;
            }
        }
        Assert.Equal(64, edges);
        Assert.Empty(reopened.ValidateDocument());
        Assert.Equal(expected.GetProperty("sourceSha256").GetString(), BorderCorpusHash(source));
    }

    private static (WordBorderStyle? Style, uint? Size, string? Color) PagesBorderSide(WordTableCellBorder border, string side) => side switch {
        "left" => (border.LeftStyle, border.LeftSize, border.LeftColorHex),
        "right" => (border.RightStyle, border.RightSize, border.RightColorHex),
        "top" => (border.TopStyle, border.TopSize, border.TopColorHex),
        "bottom" => (border.BottomStyle, border.BottomSize, border.BottomColorHex),
        _ => throw new InvalidDataException("Unknown reference edge: " + side)
    };
}
