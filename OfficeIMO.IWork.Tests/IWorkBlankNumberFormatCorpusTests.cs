using System.Security.Cryptography;
using System.Text.Json;
using OfficeIMO.IWork;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Fact]
    public void Independent_blank_number_formats_preserve_selected_source_metadata() {
        string root = Path.Combine(AppContext.BaseDirectory, "Documents", "IWorkCorpus", "numbers-parser");
        using var manifest = JsonDocument.Parse(File.ReadAllText(Path.Combine(root, "blank-number-formats.json")));
        int qualified = 0;
        foreach (JsonElement package in manifest.RootElement.GetProperty("packages").EnumerateArray()) {
            JsonElement[] cells = package.GetProperty("cells").EnumerateArray()
                .Where(cell => !cell.GetProperty("hasOtherScalarSelector").GetBoolean()).ToArray();
            if (cells.Length == 0) continue;
            string path = Path.Combine(root, package.GetProperty("source").GetString()!);
            Assert.Equal(package.GetProperty("sourceSha256").GetString(),
                Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))).ToLowerInvariant());
            IWorkNumbersProjection projection = IWorkSourceDocument.Open(path).ReadNumbers();
            foreach (JsonElement expected in cells) {
                var sheet = Assert.Single(projection.Sheets, sheet => sheet.Name == expected.GetProperty("sheet").GetString());
                var table = Assert.Single(sheet.Tables, table => table.Name == expected.GetProperty("table").GetString());
                var cell = table.GetCell(expected.GetProperty("row").GetInt32(), expected.GetProperty("column").GetInt32());
                Assert.NotNull(cell);
                Assert.Equal(IWorkCellKind.Empty, cell.Kind);
                Assert.Null(cell.Value);
                Assert.NotNull(cell.NumberFormat);
                Assert.Equal(expected.GetProperty("formatType").GetInt32() == 258
                    ? IWorkNumberFormatKind.Percentage : IWorkNumberFormatKind.Number, cell.NumberFormat.Kind);
                int decimals = expected.GetProperty("decimalPlaces").GetInt32();
                Assert.Equal(decimals == 253 ? (int?)null : decimals, cell.NumberFormat.DecimalPlaces);
                Assert.Equal(expected.GetProperty("thousandsSeparator").GetBoolean(), cell.NumberFormat.ThousandsSeparator);
                qualified++;
            }
        }
        Assert.Equal(26, qualified);
    }

    [Fact]
    public void Independent_blank_number_formats_survive_saved_xlsx() {
        string path = Path.Combine(AppContext.BaseDirectory, "Documents", "IWorkCorpus", "numbers-parser", "issue-102-v15.1.numbers");
        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(path,
            conversionOptions: new IWorkConversionOptions { AllowPartialEditableReconstruction = true, NormalizeWorksheetNames = true });
        Assert.False(result.IsVisualFallback);
        var mapping = Assert.Single(result.WorksheetMappings, mapping => mapping.SourceSheetName == "Cats" && mapping.SourceTableName == "Cats");
        var table = Assert.Single(Assert.Single(result.Projection.Sheets, sheet => sheet.Name == "Cats").Tables, table => table.Name == "Cats");
        using var saved = new MemoryStream(); result.Value.Save(saved); saved.Position = 0;
        using var reopened = OfficeIMO.Excel.ExcelDocument.Load(saved);
        var sheet = Assert.Single(reopened.Sheets, sheet => sheet.Name == mapping.DestinationName);
        var blanks = table.Cells.Where(cell => cell.Kind == IWorkCellKind.Empty && cell.NumberFormat != null).ToArray();
        Assert.Equal(21, blanks.Length);
        foreach (var cell in blanks) {
            Assert.Equal("#,##0.###############", sheet.CellAt(cell.Row, cell.Column).GetStyle().NumberFormatCode);
            Assert.Equal("", sheet.CellAt(cell.Row, cell.Column).GetValue<string>());
        }
        Assert.Empty(reopened.ValidateOpenXml());
    }
}
