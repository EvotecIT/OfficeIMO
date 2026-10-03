using System.Security.Cryptography;
using System.Text.Json;
using OfficeIMO.IWork;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Fact]
    public void Independent_blank_number_formats_preserve_selected_source_metadata() {
        string root = Path.Combine(AppContext.BaseDirectory, "Documents", "IWorkCorpus", "numbers-parser");
        using var manifest = JsonDocument.Parse(File.ReadAllText(Path.Combine(root, "blank-number-formats.json")));
        int qualified = 0, textSelected = 0, inactive = 0;
        foreach (JsonElement package in manifest.RootElement.GetProperty("packages").EnumerateArray()) {
            JsonElement[] cells = package.GetProperty("cells").EnumerateArray().ToArray();
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
                if (expected.GetProperty("hasOtherScalarSelector").GetBoolean()) {
                    Assert.Equal(IWorkNumberFormatKind.Text, cell.NumberFormat!.Kind);
                    Assert.Equal(IWorkCellUnsupportedFeatures.None, cell.UnsupportedFeatures);
                    textSelected++;
                    continue;
                }
                if ((expected.GetProperty("flags").GetUInt32() & (1u << 12)) == 0) {
                    Assert.Null(cell.NumberFormat);
                    inactive++;
                    continue;
                }
                Assert.NotNull(cell.NumberFormat);
                Assert.Equal(expected.GetProperty("formatType").GetInt32() == 258
                    ? IWorkNumberFormatKind.Percentage : IWorkNumberFormatKind.Number, cell.NumberFormat.Kind);
                int decimals = expected.GetProperty("decimalPlaces").GetInt32();
                Assert.Equal(decimals == 253 ? (int?)null : decimals, cell.NumberFormat.DecimalPlaces);
                Assert.Equal(expected.GetProperty("thousandsSeparator").GetBoolean(), cell.NumberFormat.ThousandsSeparator);
                qualified++;
            }
        }
        Assert.Equal(25, qualified);
        Assert.Equal(1, inactive);
        Assert.Equal(9, textSelected);
    }

    [Fact]
    public void Native_explicit_selection_activates_dormant_blank_percentage_format() {
        string root = Path.Combine(AppContext.BaseDirectory, "Documents", "IWorkCorpus");
        foreach (bool explicitlySelected in new[] { false, true }) {
            string path = Path.Combine(root, explicitlySelected
                ? "native-exports/numbers-blank-explicit-percentage-v14.5.numbers"
                : "numbers-parser/cross-table-formulas.numbers");
            using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(path,
                conversionOptions: new IWorkConversionOptions { Mode = IWorkConversionMode.EditableOnly,
                    AllowPartialEditableReconstruction = true, NormalizeWorksheetNames = true });
            var table = result.Projection.Sheets.Single(s => s.Name == "Main Sheet").Tables.Single(t => t.Name == "Extra Headers");
            var cell = table.GetCell(8, 1)!;
            Assert.Equal(IWorkCellKind.Empty, cell.Kind);
            if (explicitlySelected) Assert.Equal(IWorkNumberFormatKind.Percentage, cell.NumberFormat!.Kind);
            else Assert.Null(cell.NumberFormat);
            var mapping = result.WorksheetMappings.Single(m => m.SourceSheetName == "Main Sheet" && m.SourceTableName == "Extra Headers");
            using var saved = new MemoryStream(); result.Value.Save(saved); saved.Position = 0;
            using var reopened = OfficeIMO.Excel.ExcelDocument.Load(saved);
            var target = reopened.Sheets.Single(s => s.Name == mapping.DestinationName).CellAt(8, 1);
            Assert.Equal("", target.GetValue<string>());
            Assert.Equal(explicitlySelected, target.GetStyle().NumberFormatCode?.Contains('%') == true);
        }
    }

    [Fact]
    public void Explicit_blank_text_formats_match_native_xlsx_after_save_and_reopen() {
        string root = Path.Combine(AppContext.BaseDirectory, "Documents", "IWorkCorpus");
        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(
            Path.Combine(root, "numbers-parser", "cross-table-formulas.numbers"),
            conversionOptions: new IWorkConversionOptions { Mode = IWorkConversionMode.EditableOnly,
                AllowPartialEditableReconstruction = true, NormalizeWorksheetNames = true });
        var mapping = result.WorksheetMappings.Single(m => m.SourceSheetName == "Main Sheet" && m.SourceTableName == "Extra Headers");
        using var saved = new MemoryStream(); result.Value.Save(saved); saved.Position = 0;
        using var reopened = OfficeIMO.Excel.ExcelDocument.Load(saved);
        using var native = OfficeIMO.Excel.ExcelDocument.Load(Path.Combine(root, "native-exports", "numbers-blank-selectors-v14.5.xlsx"));
        var sheet = reopened.Sheets.Single(s => s.Name == mapping.DestinationName);
        var reference = native.Sheets.Single(s => s.Name == "Main Sheet - Extra Headers");
        for (int row = 9; row <= 17; row++) {
            Assert.Equal("@", reference.CellAt(row + 1, 1).GetStyle().NumberFormatCode);
            Assert.Equal(reference.CellAt(row + 1, 1).GetStyle().NumberFormatCode,
                sheet.CellAt(row, 1).GetStyle().NumberFormatCode);
            Assert.Equal("", sheet.CellAt(row, 1).GetValue<string>());
        }
        Assert.Empty(reopened.ValidateOpenXml());
    }

    [Theory]
    [InlineData("currency", IWorkCellUnsupportedFeatures.None)]
    [InlineData("date", IWorkCellUnsupportedFeatures.DateFormat)]
    [InlineData("duration", IWorkCellUnsupportedFeatures.DurationFormat)]
    [InlineData("automatic", IWorkCellUnsupportedFeatures.None)]
    public void Native_blank_scalar_selection_ignores_retained_inactive_catalogs(string format,
        IWorkCellUnsupportedFeatures expectedFeatures) {
        string root = Path.Combine(AppContext.BaseDirectory, "Documents", "IWorkCorpus", "native-exports");
        using var manifest = JsonDocument.Parse(File.ReadAllText(Path.Combine(root, "numbers-blank-selectors-v14.5.json")));
        var snapshot = manifest.RootElement.GetProperty("nativeUiEvidence").GetProperty("scalarSequence")
            .GetProperty("snapshots").EnumerateArray().Single(s => s.GetProperty("format").GetString() == format);
        string path = Path.Combine(root, snapshot.GetProperty("source").GetString()!);
        Assert.Equal(snapshot.GetProperty("sha256").GetString(),
            Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))).ToLowerInvariant());
        var projection = IWorkSourceDocument.Open(path).ReadNumbers();
        var cell = projection.Sheets.Single(s => s.Name == "Main Sheet").Tables.Single(t => t.Name == "Extra Headers").GetCell(8, 1)!;
        Assert.Equal(IWorkCellKind.Empty, cell.Kind);
        Assert.Null(cell.Value);
        Assert.Equal(expectedFeatures, cell.UnsupportedFeatures);
        if (format == "currency") Assert.Equal(IWorkNumberFormatKind.Currency, cell.NumberFormat!.Kind);
        else Assert.Null(cell.NumberFormat);
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
