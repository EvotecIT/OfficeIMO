using System.Security.Cryptography;
using System.Text.Json;
using OfficeIMO.Excel;
using OfficeIMO.IWork;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Fact]
    public void Native_Numbers_border_regions_hidden_edges_and_later_edits_match_all_saved_XLSX_edges() {
        string source = CorpusFixture("native-exports/numbers-table-borders-v14.5.numbers");
        string reference = CorpusFixture("native-exports/numbers-table-borders-v14.5.xlsx");
        using var decoded = JsonDocument.Parse(File.ReadAllText(CorpusFixture("native-exports/numbers-table-borders-v14.5-decode.json")));
        var expected = decoded.RootElement;
        Assert.Equal(expected.GetProperty("sourceSha256").GetString(), BorderCorpusHash(source));
        Assert.Equal(expected.GetProperty("nativeXlsxSha256").GetString(), BorderCorpusHash(reference));
        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(source,
            conversionOptions: new IWorkConversionOptions { AllowPartialEditableReconstruction = true });
        Assert.True(result.Report.IsPartialEditableReconstruction);
        Assert.Contains(result.Report.Diagnostics, d => d.Code == "IWORK_NUMBERS_CELL_PADDING_OMITTED"
            && d.LossKind == global::OfficeIMO.OfficeConversionLossKind.Omission);
        Assert.Throws<InvalidOperationException>(() => result.Report.RequireCompleteEditableReconstruction());
        Assert.False(result.IsVisualFallback);
        var table = Assert.Single(Assert.Single(result.Projection.Sheets).Tables);
        Assert.True(table.HasBorderDeclarations);
        Assert.Equal(4, table.RowCount); Assert.Equal(4, table.ColumnCount);
        Assert.Equal(1, table.HeaderRowCount); Assert.Equal(1, table.HeaderColumnCount); Assert.Equal(1, table.FooterRowCount);
        Assert.Equal(IWorkCellKind.Empty, table.GetCell(3, 4)!.Kind);
        Assert.False(result.Value.Sheets[0].TryGetCellValueSnapshot(3, 4, out _));
        Assert.DoesNotContain(result.Report.Diagnostics, d => d.Code is "IWORK_TABLE_BORDERS_UNSUPPORTED" or "IWORK_TABLE_CELL_FEATURES_UNASSESSED");

        using var saved = new MemoryStream(); result.Value.Save(saved); saved.Position = 0;
        using var reopened = ExcelDocument.Load(saved);
        using var native = ExcelDocument.Load(reference);
        var output = Assert.Single(reopened.Sheets);
        var nativeSheet = Assert.Single(native.Sheets);
        int edges = 0;
        foreach (var cell in expected.GetProperty("cells").EnumerateArray()) {
            int row = cell.GetProperty("row").GetInt32(), column = cell.GetProperty("column").GetInt32();
            string? value = cell.GetProperty("value").GetString();
            Assert.Equal(value, table.GetCell(row, column)?.Value);
            Assert.Equal(value ?? string.Empty, output.CellAt(row, column).GetValue<string>() ?? string.Empty);
            Assert.Equal(nativeSheet.CellAt(row, column).GetValue<string>() ?? string.Empty,
                output.CellAt(row, column).GetValue<string>() ?? string.Empty);
            var borders = table.GetBorders(row, column);
            var savedBorders = Assert.IsType<ExcelCellBorderSnapshot>(output.GetCellStyle(row, column).Border);
            var nativeBorders = Assert.IsType<ExcelCellBorderSnapshot>(nativeSheet.GetCellStyle(row, column).Border);
            foreach (var edge in cell.GetProperty("edges").EnumerateObject()) {
                var side = edge.Name switch {
                    "left" => borders.Left, "right" => borders.Right, "top" => borders.Top, "bottom" => borders.Bottom,
                    _ => throw new InvalidDataException("Unknown reference edge: " + edge.Name)
                };
                double width = edge.Value.GetProperty("widthPoints").GetDouble();
                string rgb = edge.Value.GetProperty("rgb").GetString()!;
                AssertBorder(side, rgb, width);
                var savedSide = BorderCorpusSide(savedBorders, edge.Name);
                var nativeSide = BorderCorpusSide(nativeBorders, edge.Name);
                Assert.Equal(width == 0 ? "none" : width == 1 ? "thin" : "medium", savedSide?.Style ?? "none");
                Assert.Equal(nativeSide?.Style ?? "none", savedSide?.Style ?? "none");
                if (width > 0) {
                    Assert.Equal("FF" + rgb, savedSide!.ColorArgb, ignoreCase: true);
                    Assert.Equal(nativeSide!.ColorArgb, savedSide.ColorArgb, ignoreCase: true);
                }
                edges++;
            }
        }
        Assert.Equal(64, edges);
        Assert.False(nativeSheet.TryGetCellValueSnapshot(3, 4, out _));
        Assert.False(output.TryGetCellValueSnapshot(3, 4, out _));
        Assert.NotNull(output.GetCellStyle(3, 4).Border); // Empty D3 retains its four styled edges.
        Assert.Empty(reopened.ValidateOpenXml());
        Assert.Equal(expected.GetProperty("sourceSha256").GetString(), BorderCorpusHash(source));
    }

    private static ExcelBorderSideSnapshot? BorderCorpusSide(ExcelCellBorderSnapshot border, string name) => name switch {
        "left" => border.Left, "right" => border.Right, "top" => border.Top, "bottom" => border.Bottom,
        _ => throw new InvalidDataException("Unknown reference edge: " + name)
    };

    private static string BorderCorpusHash(string path) => Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))).ToLowerInvariant();
}
