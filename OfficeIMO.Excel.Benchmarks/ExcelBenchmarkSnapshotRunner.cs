using System.Text.Json;
using ClosedXML.Excel;

namespace OfficeIMO.Excel.Benchmarks;

internal static class ExcelBenchmarkSnapshotRunner {
    private const int SnapshotRowCount = 2500;
    private const int WarmupIterations = 2;
    private const int MeasuredIterations = 5;

    internal static string WriteSnapshot(string outputPath) {
        if (string.IsNullOrWhiteSpace(outputPath)) {
            throw new ArgumentException("Output path must not be empty.", nameof(outputPath));
        }

        var rows = ExcelBenchmarkScenarioFactory.CreateSalesRecords(SnapshotRowCount);
        var workbookBytes = ExcelBenchmarkScenarioFactory.CreateWorkbookBytes(rows);

        var snapshot = new ExcelBenchmarkSnapshot {
            GeneratedAtUtc = DateTime.UtcNow,
            Framework = System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription,
            MachineName = Environment.MachineName,
            RowCount = SnapshotRowCount,
            Scenarios = [
                Measure("Excel: Write report (2,500 rows)", "OfficeIMO.Excel", "Insert objects, table, and AutoFit on a single worksheet.",
                    () => OfficeImoWrite(rows)),
                Measure("Excel: Write report (2,500 rows)", "ClosedXML", "Insert table and adjust columns on a single worksheet.",
                    () => ClosedXmlWrite(rows)),
                Measure("Excel: Read objects (2,500 rows)", "OfficeIMO.Excel", "Reader path materializing dictionary-style rows.",
                    () => OfficeImoReadObjects(workbookBytes)),
                Measure("Excel: Read DataTable (2,500 rows)", "OfficeIMO.Excel", "Reader path materializing a DataTable from the same workbook.",
                    () => OfficeImoReadDataTable(workbookBytes)),
                Measure("Excel: Read rows (2,500 rows)", "ClosedXML", "Worksheet row iteration over the same workbook payload.",
                    () => ClosedXmlReadRows(workbookBytes)),
                Measure("Excel: Load/edit/save (2,500 rows)", "OfficeIMO.Excel", "Load workbook, add review column, save to memory.",
                    () => OfficeImoRoundTrip(workbookBytes)),
                Measure("Excel: Load/edit/save (2,500 rows)", "ClosedXML", "Load workbook, add review column, save to memory.",
                    () => ClosedXmlRoundTrip(workbookBytes))
            ]
        };

        string? directory = Path.GetDirectoryName(outputPath);
        if (!string.IsNullOrEmpty(directory)) {
            Directory.CreateDirectory(directory);
        }

        var options = new JsonSerializerOptions { WriteIndented = true };
        File.WriteAllText(outputPath, JsonSerializer.Serialize(snapshot, options));
        return outputPath;
    }

    private static ExcelBenchmarkSnapshotScenario Measure(string name, string library, string notes, Func<int> action) {
        var measurement = BenchmarkMeasurement.Measure(WarmupIterations, MeasuredIterations, action);

        return new ExcelBenchmarkSnapshotScenario {
            Name = name,
            Library = library,
            Result = $"{measurement.AverageMilliseconds:F1} ms avg",
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
            var sheet = document.AddWorkSheet("Data");
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

    private static int OfficeImoReadObjects(byte[] workbookBytes) {
        using var stream = new MemoryStream(workbookBytes, writable: false);
        using var reader = ExcelDocumentReader.Open(stream);
        return reader.GetSheet("Data").ReadObjects(ExcelBenchmarkScenarioFactory.BuildDataRange(SnapshotRowCount)).Count();
    }

    private static int OfficeImoReadDataTable(byte[] workbookBytes) {
        using var stream = new MemoryStream(workbookBytes, writable: false);
        using var reader = ExcelDocumentReader.Open(stream);
        return reader.GetSheet("Data").ReadRangeAsDataTable(ExcelBenchmarkScenarioFactory.BuildDataRange(SnapshotRowCount), headersInFirstRow: true).Rows.Count;
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

    private static int OfficeImoRoundTrip(byte[] workbookBytes) {
        using var input = new MemoryStream(workbookBytes, writable: false);
        using var output = new MemoryStream();

        using (var document = ExcelDocument.Load(input)) {
            var sheet = document["Data"];
            sheet.CellValue(1, 9, "ReviewStatus");

            int limit = Math.Min(SnapshotRowCount + 1, 101);
            for (int row = 2; row <= limit; row++) {
                sheet.CellValue(row, 9, "Reviewed");
            }

            document.Save(output);
        }

        return checked((int)output.Length);
    }

    private static int ClosedXmlRoundTrip(byte[] workbookBytes) {
        using var input = new MemoryStream(workbookBytes, writable: false);
        using var output = new MemoryStream();
        using var workbook = new XLWorkbook(input);
        var worksheet = workbook.Worksheet("Data");

        worksheet.Cell(1, 9).Value = "ReviewStatus";
        int limit = Math.Min(SnapshotRowCount + 1, 101);
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
}
