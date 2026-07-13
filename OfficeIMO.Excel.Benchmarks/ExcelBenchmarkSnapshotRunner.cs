using System.Globalization;
using System.Text.Json;
using System.Text.Json.Nodes;
using ClosedXML.Excel;

namespace OfficeIMO.Excel.Benchmarks;

internal static class ExcelBenchmarkSnapshotRunner {
    private const int DefaultRowCount = 2500;
    private const int WarmupIterations = 2;
    private const int MeasuredIterations = 5;

    internal static string WriteSnapshot(string outputPath, int rowCount = DefaultRowCount, string? websiteBenchmarkDataPath = null) {
        if (string.IsNullOrWhiteSpace(outputPath)) {
            throw new ArgumentException("Output path must not be empty.", nameof(outputPath));
        }
        if (rowCount <= 0) {
            throw new ArgumentOutOfRangeException(nameof(rowCount));
        }

        var rows = ExcelBenchmarkScenarioFactory.CreateSalesRecords(rowCount);
        var workbookBytes = ExcelBenchmarkScenarioFactory.CreateWorkbookBytes(rows);
        string rowLabel = rowCount.ToString("N0", CultureInfo.InvariantCulture);
        string dataRange = ExcelBenchmarkScenarioFactory.BuildDataRange(rowCount);

        var snapshot = new ExcelBenchmarkSnapshot {
            GeneratedAtUtc = DateTime.UtcNow,
            Framework = System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription,
            MachineName = Environment.MachineName,
            RowCount = rowCount,
            Scenarios = [
                Measure($"Excel: Write report ({rowLabel} rows)", "OfficeIMO.Excel", "Insert objects, table, and AutoFit on a single worksheet.",
                    () => OfficeImoWrite(rows)),
                Measure($"Excel: Write report ({rowLabel} rows)", "ClosedXML", "Insert table and adjust columns on a single worksheet.",
                    () => ClosedXmlWrite(rows)),
                Measure($"Excel: Read objects ({rowLabel} rows)", "OfficeIMO.Excel", "Reader path materializing dictionary-style rows.",
                    () => OfficeImoReadObjects(workbookBytes, dataRange)),
                Measure($"Excel: Read typed objects ({rowLabel} rows)", "OfficeIMO.Excel", "Reader path mapping rows directly to typed objects.",
                    () => OfficeImoReadObjectsAs(workbookBytes, dataRange)),
                Measure($"Excel: Read DataTable ({rowLabel} rows)", "OfficeIMO.Excel", "Reader path materializing a DataTable from the same workbook.",
                    () => OfficeImoReadDataTable(workbookBytes, dataRange)),
                Measure($"Excel: Read rows ({rowLabel} rows)", "ClosedXML", "Worksheet row iteration over the same workbook payload.",
                    () => ClosedXmlReadRows(workbookBytes)),
                Measure($"Excel: Load/edit/save ({rowLabel} rows)", "OfficeIMO.Excel", "Load workbook, add review column, save to memory.",
                    () => OfficeImoRoundTrip(workbookBytes, rowCount)),
                Measure($"Excel: Load/edit/save ({rowLabel} rows)", "ClosedXML", "Load workbook, add review column, save to memory.",
                    () => ClosedXmlRoundTrip(workbookBytes, rowCount))
            ]
        };

        string? directory = Path.GetDirectoryName(outputPath);
        if (!string.IsNullOrEmpty(directory)) {
            Directory.CreateDirectory(directory);
        }

        var options = new JsonSerializerOptions { WriteIndented = true };
        File.WriteAllText(outputPath, JsonSerializer.Serialize(snapshot, options));
        if (!string.IsNullOrWhiteSpace(websiteBenchmarkDataPath)) {
            WriteWebsiteBenchmarkData(websiteBenchmarkDataPath, snapshot, DateOnly.FromDateTime(DateTime.Now));
        }

        return outputPath;
    }

    private static void WriteWebsiteBenchmarkData(string path, ExcelBenchmarkSnapshot snapshot, DateOnly snapshotDate) {
        if (!File.Exists(path)) {
            throw new FileNotFoundException("Website benchmark data file was not found.", path);
        }

        JsonNode root = JsonNode.Parse(File.ReadAllText(path)) ?? throw new InvalidOperationException("Website benchmark data is empty.");
        JsonArray scenarios = root["scenarios"]?.AsArray() ?? throw new InvalidOperationException("Website benchmark data does not contain a scenarios array.");
        var snapshotScenarios = snapshot.Scenarios.ToDictionary(
            scenario => BuildScenarioKey(scenario.Name, scenario.Library),
            scenario => scenario,
            StringComparer.Ordinal);
        string dateText = snapshotDate.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);

        foreach (JsonNode? scenarioNode in scenarios) {
            if (scenarioNode is null) {
                continue;
            }

            string? name = scenarioNode["name"]?.GetValue<string>();
            string? library = scenarioNode["library"]?.GetValue<string>();
            if (name is null || library is null || !snapshotScenarios.TryGetValue(BuildScenarioKey(name, library), out var snapshotScenario)) {
                continue;
            }

            scenarioNode["result"] = snapshotScenario.Result;
            scenarioNode["notes"] = string.Create(
                CultureInfo.InvariantCulture,
                $"5-sample snapshot dated {dateText}. Median {snapshotScenario.MedianMilliseconds:F1} ms. {snapshotScenario.Notes}");
        }

        string? directory = Path.GetDirectoryName(path);
        if (!string.IsNullOrEmpty(directory)) {
            Directory.CreateDirectory(directory);
        }

        File.WriteAllText(path, root.ToJsonString(new JsonSerializerOptions { WriteIndented = true }));
    }

    private static string BuildScenarioKey(string name, string library)
        => string.Concat(name, "\u001f", library);

    private static ExcelBenchmarkSnapshotScenario Measure(string name, string library, string notes, Func<int> action) {
        var measurement = BenchmarkMeasurement.Measure(WarmupIterations, MeasuredIterations, action);

        return new ExcelBenchmarkSnapshotScenario {
            Name = name,
            Library = library,
            Result = string.Create(CultureInfo.InvariantCulture, $"{measurement.AverageMilliseconds:F1} ms avg"),
            Notes = notes,
            OutputMetric = measurement.OutputMetric,
            AverageMilliseconds = measurement.AverageMilliseconds,
            MedianMilliseconds = measurement.MedianMilliseconds,
            SamplesMilliseconds = measurement.SamplesMilliseconds.ToList()
        };
    }

    private static int OfficeImoWrite(IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows) {
        using var stream = new MemoryStream();
        using (var document = ExcelDocument.Create(stream)) {
            document.Execution.SaveWorksheetAfterAutoFit = false;
            var sheet = document.AddWorksheet("Data");
            ExcelBenchmarkScenarioFactory.PopulateOfficeImoWorksheet(sheet, rows);
        }

        return checked((int)stream.Length);
    }

    private static int ClosedXmlWrite(IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows) {
        using var stream = new MemoryStream();
        using var workbook = new XLWorkbook();
        var worksheet = workbook.Worksheets.Add("Data");
        ExcelBenchmarkScenarioFactory.PopulateClosedXmlWorksheet(worksheet, rows);
        workbook.SaveAs(stream);
        return checked((int)stream.Length);
    }

    private static int OfficeImoReadObjects(byte[] workbookBytes, string dataRange) {
        using var stream = new MemoryStream(workbookBytes, writable: false);
        using var reader = ExcelDocumentReader.Open(stream);
        return reader.GetSheet("Data").ReadObjects(dataRange).Count();
    }

    private static int OfficeImoReadObjectsAs(byte[] workbookBytes, string dataRange) {
        using var stream = new MemoryStream(workbookBytes, writable: false);
        using var reader = ExcelDocumentReader.Open(stream);
        return reader.GetSheet("Data").ReadObjects<ReadSalesRecord>(dataRange).Count();
    }

    private static int OfficeImoReadDataTable(byte[] workbookBytes, string dataRange) {
        using var stream = new MemoryStream(workbookBytes, writable: false);
        using var reader = ExcelDocumentReader.Open(stream);
        return reader.GetSheet("Data").ReadRangeAsDataTable(dataRange, headersInFirstRow: true).Rows.Count;
    }

    private static int ClosedXmlReadRows(byte[] workbookBytes) {
        using var stream = new MemoryStream(workbookBytes, writable: false);
        using var workbook = new XLWorkbook(stream);
        var worksheet = workbook.Worksheet("Data");
        int lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 0;
        int count = 0;

        for (int row = 2; row <= lastRow; row++) {
            _ = worksheet.Cell(row, 1).GetValue<int>();
            _ = worksheet.Cell(row, 5).GetValue<double>();
            count++;
        }

        return count;
    }

    private static int OfficeImoRoundTrip(byte[] workbookBytes, int rowCount) {
        using var input = new MemoryStream(workbookBytes, writable: false);
        using var output = new MemoryStream();

        using (var document = ExcelDocument.Load(input)) {
            var sheet = document["Data"];
            sheet.CellValue(1, 9, "ReviewStatus");

            int limit = Math.Min(rowCount + 1, 101);
            for (int row = 2; row <= limit; row++) {
                sheet.CellValue(row, 9, "Reviewed");
            }

            document.Save(output);
        }

        return checked((int)output.Length);
    }

    private static int ClosedXmlRoundTrip(byte[] workbookBytes, int rowCount) {
        using var input = new MemoryStream(workbookBytes, writable: false);
        using var output = new MemoryStream();
        using var workbook = new XLWorkbook(input);
        var worksheet = workbook.Worksheet("Data");

        worksheet.Cell(1, 9).Value = "ReviewStatus";
        int limit = Math.Min(rowCount + 1, 101);
        for (int row = 2; row <= limit; row++) {
            worksheet.Cell(row, 9).Value = "Reviewed";
        }

        workbook.SaveAs(output);
        return checked((int)output.Length);
    }

    private sealed class ExcelBenchmarkSnapshot {
        public DateTime GeneratedAtUtc { get; init; }
        public string Framework { get; init; } = string.Empty;
        public string MachineName { get; init; } = string.Empty;
        public int RowCount { get; init; }
        public List<ExcelBenchmarkSnapshotScenario> Scenarios { get; init; } = [];
    }

    private sealed class ExcelBenchmarkSnapshotScenario {
        public string Name { get; init; } = string.Empty;
        public string Library { get; init; } = string.Empty;
        public string Result { get; init; } = string.Empty;
        public string Notes { get; init; } = string.Empty;
        public int OutputMetric { get; init; }
        public double AverageMilliseconds { get; init; }
        public double MedianMilliseconds { get; init; }
        public List<double> SamplesMilliseconds { get; init; } = [];
    }

    private sealed class ReadSalesRecord {
        public int Id { get; set; }
        public string Region { get; set; } = string.Empty;
        public string Owner { get; set; } = string.Empty;
        public DateTime CreatedOn { get; set; }
        public double Amount { get; set; }
        public int Units { get; set; }
        public bool Active { get; set; }
        public string Notes { get; set; } = string.Empty;
    }
}
