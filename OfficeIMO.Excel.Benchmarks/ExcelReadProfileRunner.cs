using System.Text.Json;
using ClosedXML.Excel;

namespace OfficeIMO.Excel.Benchmarks;

internal static class ExcelReadProfileRunner {
    private const int DefaultRowCount = 2500;
    private const int WarmupIterations = 2;
    private const int MeasuredIterations = 5;

    internal static string WriteProfile(string outputPath, int rowCount = DefaultRowCount) {
        if (string.IsNullOrWhiteSpace(outputPath)) {
            throw new ArgumentException("Output path must not be empty.", nameof(outputPath));
        }
        if (rowCount <= 0) {
            throw new ArgumentOutOfRangeException(nameof(rowCount));
        }

        var rows = ExcelBenchmarkScenarioFactory.CreateSalesRecords(rowCount);
        byte[] workbookBytes = ExcelBenchmarkScenarioFactory.CreateWorkbookBytes(rows);
        string dataRange = ExcelBenchmarkScenarioFactory.BuildDataRange(rowCount);

        var profile = new ExcelReadProfile {
            GeneratedAtUtc = DateTime.UtcNow,
            Framework = System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription,
            MachineName = Environment.MachineName,
            RowCount = rowCount,
            Scenarios = [
                Measure("OfficeIMO.Excel", "ReadObjects", "Automatic execution policy.", () => OfficeImoReadObjects(workbookBytes, dataRange, null)),
                Measure("OfficeIMO.Excel", "ReadObjects.Sequential", "Forced sequential range conversion.", () => OfficeImoReadObjects(workbookBytes, dataRange, ExecutionMode.Sequential)),
                Measure("OfficeIMO.Excel", "ReadObjects.Parallel", "Forced parallel range conversion.", () => OfficeImoReadObjects(workbookBytes, dataRange, ExecutionMode.Parallel)),
                Measure("OfficeIMO.Excel", "ReadObjectsAs", "Typed object materialization with automatic execution policy.", () => OfficeImoReadObjectsAs(workbookBytes, dataRange, null)),
                Measure("OfficeIMO.Excel", "ReadObjectsAs.Sequential", "Typed object materialization with forced sequential range conversion.", () => OfficeImoReadObjectsAs(workbookBytes, dataRange, ExecutionMode.Sequential)),
                Measure("OfficeIMO.Excel", "ReadObjectsAs.Parallel", "Typed object materialization with forced parallel range conversion.", () => OfficeImoReadObjectsAs(workbookBytes, dataRange, ExecutionMode.Parallel)),
                Measure("OfficeIMO.Excel", "ReadRangeAsDataTable", "Automatic execution policy.", () => OfficeImoReadDataTable(workbookBytes, dataRange, null)),
                Measure("OfficeIMO.Excel", "ReadRangeAsDataTable.Sequential", "Forced sequential range conversion.", () => OfficeImoReadDataTable(workbookBytes, dataRange, ExecutionMode.Sequential)),
                Measure("OfficeIMO.Excel", "ReadRangeAsDataTable.Parallel", "Forced parallel range conversion.", () => OfficeImoReadDataTable(workbookBytes, dataRange, ExecutionMode.Parallel)),
                Measure("ClosedXML", "ReadRows", "Worksheet row iteration over the same workbook payload.", () => ClosedXmlReadRows(workbookBytes))
            ]
        };

        string? directory = Path.GetDirectoryName(outputPath);
        if (!string.IsNullOrEmpty(directory)) {
            Directory.CreateDirectory(directory);
        }

        var options = new JsonSerializerOptions { WriteIndented = true };
        File.WriteAllText(outputPath, JsonSerializer.Serialize(profile, options));
        return outputPath;
    }

    private static ExcelReadProfileScenario Measure(string library, string name, string notes, Func<int> action) {
        var measurement = BenchmarkMeasurement.Measure(WarmupIterations, MeasuredIterations, action);

        return new ExcelReadProfileScenario {
            Library = library,
            Name = name,
            Notes = notes,
            OutputMetric = measurement.OutputMetric,
            AverageMilliseconds = measurement.AverageMilliseconds,
            MedianMilliseconds = measurement.MedianMilliseconds,
            SamplesMilliseconds = measurement.SamplesMilliseconds.ToList()
        };
    }

    private static int OfficeImoReadObjects(byte[] workbookBytes, string dataRange, ExecutionMode? mode) {
        using var stream = new MemoryStream(workbookBytes, writable: false);
        using var reader = ExcelDocumentReader.Open(stream);
        return reader.GetSheet("Data").ReadObjects(dataRange, mode).Count();
    }

    private static int OfficeImoReadObjectsAs(byte[] workbookBytes, string dataRange, ExecutionMode? mode) {
        using var stream = new MemoryStream(workbookBytes, writable: false);
        using var reader = ExcelDocumentReader.Open(stream);
        return reader.GetSheet("Data").ReadObjects<ReadSalesRecord>(dataRange, mode).Count();
    }

    private static int OfficeImoReadDataTable(byte[] workbookBytes, string dataRange, ExecutionMode? mode) {
        using var stream = new MemoryStream(workbookBytes, writable: false);
        using var reader = ExcelDocumentReader.Open(stream);
        return reader.GetSheet("Data").ReadRangeAsDataTable(dataRange, headersInFirstRow: true, mode: mode).Rows.Count;
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

    private sealed class ExcelReadProfile {
        public DateTime GeneratedAtUtc { get; init; }
        public string Framework { get; init; } = string.Empty;
        public string MachineName { get; init; } = string.Empty;
        public int RowCount { get; init; }
        public List<ExcelReadProfileScenario> Scenarios { get; init; } = [];
    }

    private sealed class ExcelReadProfileScenario {
        public string Library { get; init; } = string.Empty;
        public string Name { get; init; } = string.Empty;
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
