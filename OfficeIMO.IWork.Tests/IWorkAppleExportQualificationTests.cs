using System.Globalization;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Xml.Linq;
using OfficeIMO.Excel;
using OfficeIMO.IWork;

namespace OfficeIMO.IWork.Tests;

public sealed class IWorkAppleExportQualificationTests {
    [Fact]
    public void Numbers_formula_expressions_and_stable_caches_match_independent_Apple_export() {
        string corpus = Path.Combine(AppContext.BaseDirectory, "Documents", "IWorkCorpus");
        string references = Path.Combine(corpus, "native-exports");
        using JsonDocument manifest = JsonDocument.Parse(File.ReadAllText(
            Path.Combine(references, "numbers-formulas-v14.5.json")));
        JsonElement evidence = manifest.RootElement;
        string sourcePath = Path.Combine(corpus, evidence.GetProperty("sourceFixture").GetString()!);
        Assert.Equal(evidence.GetProperty("sourceSha256").GetString(), Hash(sourcePath));
        foreach (JsonElement artifact in evidence.GetProperty("artifacts").EnumerateArray()) {
            Assert.Equal(artifact.GetProperty("sha256").GetString(),
                Hash(Path.Combine(references, artifact.GetProperty("path").GetString()!)));
        }
        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(sourcePath);
        Assert.False(result.IsVisualFallback);
        Assert.Equal(2, result.WorksheetMappings.Count);
        Assert.Contains(result.Report.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_NUMBERS_ERROR_VALUE_APPROXIMATED"
            && diagnostic.LossKind == global::OfficeIMO.OfficeConversionLossKind.Approximation);
        using var saved = new MemoryStream();
        result.Value.Save(saved);
        saved.Position = 0;
        using ExcelDocument reopened = ExcelDocument.Load(saved);
        Assert.Equal(2, reopened.Sheets.Count);
        saved.Position = 0;
        using var converted = new ZipArchive(saved, ZipArchiveMode.Read, leaveOpen: true);
        using var native = ZipFile.OpenRead(Path.Combine(references, "numbers-formulas-v14.5.xlsx"));
        JsonElement[] expectations = evidence.GetProperty("formulaExpectations").EnumerateArray().ToArray();
        Assert.Equal(28, expectations.Length);
        Assert.Equal(expectations.Length, result.Report.FormulaSummary.TotalCount);
        Assert.Equal(expectations.Length, result.Report.FormulaSummary.CompleteExpressionCount);
        var actualSheets = Enumerable.Range(1, 2).ToDictionary(index => index, index => ReadCells(converted, index));
        var nativeSheets = Enumerable.Range(1, 2).ToDictionary(index => index, index => ReadCells(native, index));
        Assert.Equal(28, actualSheets.Values.Sum(cells => cells.Values.Count(cell => cell.Element(Spreadsheet + "f") != null)));
        foreach (JsonElement expected in expectations) {
            int tableIndex = expected.GetProperty("tableIndex").GetInt32();
            XElement actual = actualSheets[tableIndex][expected.GetProperty("sourceCell").GetString()!];
            XElement oracle = nativeSheets[tableIndex][expected.GetProperty("nativeExportCell").GetString()!];
            Assert.Equal(expected.GetProperty("nativeExportFormula").GetString(), oracle.Element(Spreadsheet + "f")!.Value);
            Assert.Equal(expected.GetProperty("sourceFormula").GetString(),
                NormalizeFormula(actual.Element(Spreadsheet + "f")!.Value));
            if (!expected.GetProperty("compareCache").GetBoolean()) continue;
            string nativeValue = oracle.Element(Spreadsheet + "v")!.Value;
            Assert.Equal(expected.GetProperty("nativeCachedValue").GetString(), nativeValue);
            string actualValue = actual.Element(Spreadsheet + "v")!.Value;
            string nativeType = expected.GetProperty("nativeCacheType").GetString()!;
            string actualType = actual.Attribute("t")?.Value ?? "n";
            Assert.Equal(nativeType, actualType);
            if (nativeType == "n") {
                Assert.Equal(double.Parse(nativeValue, CultureInfo.InvariantCulture),
                    double.Parse(actualValue, CultureInfo.InvariantCulture), precision: 8);
            } else {
                Assert.Equal(nativeValue, actualValue);
            }
        }
    }

    private static readonly XNamespace Spreadsheet = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    [Fact]
    public void Numbers_declared_sizes_stay_within_the_pinned_native_grid_comparison_bounds() {
        string corpus = Path.Combine(AppContext.BaseDirectory, "Documents", "IWorkCorpus");
        string references = Path.Combine(corpus, "native-exports");
        using JsonDocument manifest = JsonDocument.Parse(File.ReadAllText(
            Path.Combine(references, "numbers-formulas-v14.5.json")));
        JsonElement evidence = manifest.RootElement;
        string sourcePath = Path.Combine(corpus, evidence.GetProperty("sourceFixture").GetString()!);
        Assert.Equal(evidence.GetProperty("sourceSha256").GetString(), Hash(sourcePath));
        foreach (JsonElement artifact in evidence.GetProperty("artifacts").EnumerateArray())
            Assert.Equal(artifact.GetProperty("sha256").GetString(),
                Hash(Path.Combine(references, artifact.GetProperty("path").GetString()!)));
        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(sourcePath);
        Assert.False(result.IsVisualFallback);
        var tables = result.Projection.Sheets.SelectMany(sheet => sheet.Tables).ToArray();
        using var saved = new MemoryStream();
        result.Value.Save(saved); saved.Position = 0;
        using ExcelDocument reopened = ExcelDocument.Load(saved);
        Assert.Equal(tables.Length, reopened.Sheets.Count);
        saved.Position = 0;
        using var converted = new ZipArchive(saved, ZipArchiveMode.Read, leaveOpen: true);
        using var native = ZipFile.OpenRead(Path.Combine(references, "numbers-formulas-v14.5.xlsx"));
        JsonElement[] grids = evidence.GetProperty("pdf").GetProperty("tableGeometry")
            .GetProperty("tables").EnumerateArray().ToArray();
        Assert.Equal(tables.Length, grids.Length);
        int titleOffset = evidence.GetProperty("exportOptions").GetProperty("nativeTableTitleRowOffset").GetInt32();
        foreach (JsonElement grid in grids) {
            int index = grid.GetProperty("tableIndex").GetInt32();
            IWork.IWorkTable table = tables[index - 1];
            double[] widths = grid.GetProperty("columnWidthsPoints").EnumerateArray().Select(value => value.GetDouble()).ToArray();
            double[] heights = grid.GetProperty("rowHeightsPoints").EnumerateArray().Select(value => value.GetDouble()).ToArray();
            Assert.Equal(table.ColumnCount, widths.Length);
            Assert.Equal(table.RowCount, heights.Length);
            for (int column = 1; column <= table.ColumnCount; column++)
                Assert.Equal(widths[column - 1], table.GetColumnWidth(column)!.Value, precision: 3);
            Dictionary<int, double> actualRows = ReadRowHeights(converted, index);
            Dictionary<int, double> nativeRows = ReadRowHeights(native, index);
            for (int row = 1; row <= table.RowCount; row++) {
                double declared = table.GetRowHeight(row)!.Value;
                Assert.Equal(declared, actualRows[row], precision: 6);
                // This fixture's declared heights differ from Apple's rendered
                // heights. Bound the observed difference without claiming equivalent automatic row sizing.
                Assert.InRange(Math.Abs(declared - heights[row - 1]), 0d, 0.15d);
                Assert.InRange(Math.Abs(nativeRows[row + titleOffset] - heights[row - 1]), 0d, 0.025d);
            }
        }
    }

    private static Dictionary<int, double> ReadRowHeights(ZipArchive archive, int sheetIndex) {
        using Stream stream = archive.GetEntry($"xl/worksheets/sheet{sheetIndex}.xml")!.Open();
        XElement sheet = XDocument.Load(stream).Root!;
        double defaultHeight = double.Parse(sheet.Element(Spreadsheet + "sheetFormatPr")!
            .Attribute("defaultRowHeight")!.Value, CultureInfo.InvariantCulture);
        return sheet.Descendants(Spreadsheet + "row").ToDictionary(
            row => int.Parse(row.Attribute("r")!.Value, CultureInfo.InvariantCulture),
            row => row.Attribute("ht") is XAttribute height
                ? double.Parse(height.Value, CultureInfo.InvariantCulture) : defaultHeight);
    }

    private static Dictionary<string, XElement> ReadCells(ZipArchive archive, int sheetIndex) {
        using Stream stream = archive.GetEntry($"xl/worksheets/sheet{sheetIndex}.xml")!.Open();
        return XDocument.Load(stream).Descendants(Spreadsheet + "c")
            .ToDictionary(cell => cell.Attribute("r")!.Value);
    }

    private static string Hash(string path) => Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))).ToLowerInvariant();

    private static string NormalizeFormula(string value) {
        var text = new StringBuilder();
        bool quoted = false;
        foreach (char character in value) {
            if (character == '"') quoted = !quoted;
            if (quoted || !char.IsWhiteSpace(character)) text.Append(character);
        }
        return text.ToString();
    }
}
