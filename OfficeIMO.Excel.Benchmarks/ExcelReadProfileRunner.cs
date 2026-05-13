using System.Diagnostics;
using System.Text.Json;
using ClosedXML.Excel;

namespace OfficeIMO.Excel.Benchmarks;

internal static class ExcelReadProfileRunner {
    private const int DefaultRowCount = 2500;
    private const int SparseLastRow = 100_001;
#if DEBUG
    private const string BuildConfiguration = "Debug";
#else
    private const string BuildConfiguration = "Release";
#endif
    internal const int DefaultWarmupIterations = 2;
    internal const int DefaultMeasuredIterations = 5;

    internal static string WriteProfile(
        string outputPath,
        int rowCount = DefaultRowCount,
        int warmupIterations = DefaultWarmupIterations,
        int measuredIterations = DefaultMeasuredIterations) {
        if (string.IsNullOrWhiteSpace(outputPath)) {
            throw new ArgumentException("Output path must not be empty.", nameof(outputPath));
        }
        if (rowCount <= 0) {
            throw new ArgumentOutOfRangeException(nameof(rowCount));
        }
        if (warmupIterations <= 0) {
            throw new ArgumentOutOfRangeException(nameof(warmupIterations));
        }
        if (measuredIterations <= 0) {
            throw new ArgumentOutOfRangeException(nameof(measuredIterations));
        }

        var rows = ExcelBenchmarkScenarioFactory.CreateSalesRecords(rowCount);
        byte[] workbookBytes = ExcelBenchmarkScenarioFactory.CreateWorkbookBytes(rows);
        string dataRange = ExcelBenchmarkScenarioFactory.BuildDataRange(rowCount);
        byte[] sparseWorkbookBytes = CreateSparseWorkbookBytes(SparseLastRow);
        string sparseRange = $"A1:A{SparseLastRow}";

        List<ExcelReadProfileScenario> scenarios = [];
        scenarios.AddRange(MeasureGroup("OfficeIMO.Excel", [
            new ReadProfileCase("ReadObjects", "Automatic execution policy.", () => OfficeImoReadObjects(workbookBytes, dataRange, null)),
            new ReadProfileCase("ReadObjects.Sequential", "Forced sequential range conversion.", () => OfficeImoReadObjects(workbookBytes, dataRange, ExecutionMode.Sequential)),
            new ReadProfileCase("ReadObjects.Parallel", "Forced parallel range conversion.", () => OfficeImoReadObjects(workbookBytes, dataRange, ExecutionMode.Parallel))
        ], warmupIterations, measuredIterations));
        scenarios.AddRange(MeasureGroup("OfficeIMO.Excel", [
            new ReadProfileCase("ReadObjectsAs", "Typed object materialization with automatic execution policy.", () => OfficeImoReadObjectsAs(workbookBytes, dataRange, null)),
            new ReadProfileCase("ReadObjectsAs.Sequential", "Typed object materialization with forced sequential range conversion.", () => OfficeImoReadObjectsAs(workbookBytes, dataRange, ExecutionMode.Sequential)),
            new ReadProfileCase("ReadObjectsAs.Parallel", "Typed object materialization with forced parallel range conversion.", () => OfficeImoReadObjectsAs(workbookBytes, dataRange, ExecutionMode.Parallel))
        ], warmupIterations, measuredIterations));
        scenarios.AddRange(MeasureGroup("OfficeIMO.Excel", [
            new ReadProfileCase("ReadRange", "Dense 2D array read with automatic execution policy.", () => OfficeImoReadRange(workbookBytes, dataRange, null)),
            new ReadProfileCase("ReadRange.Sequential", "Dense 2D array read with forced sequential conversion.", () => OfficeImoReadRange(workbookBytes, dataRange, ExecutionMode.Sequential)),
            new ReadProfileCase("ReadRange.Parallel", "Dense 2D array read with forced parallel conversion.", () => OfficeImoReadRange(workbookBytes, dataRange, ExecutionMode.Parallel))
        ], warmupIterations, measuredIterations));
        scenarios.AddRange(MeasureGroup("OfficeIMO.Excel", [
            new ReadProfileCase("ReadRangeAsDataTable", "Automatic execution policy.", () => OfficeImoReadDataTable(workbookBytes, dataRange, null)),
            new ReadProfileCase("ReadRangeAsDataTable.Sequential", "Forced sequential range conversion.", () => OfficeImoReadDataTable(workbookBytes, dataRange, ExecutionMode.Sequential)),
            new ReadProfileCase("ReadRangeAsDataTable.Parallel", "Forced parallel range conversion.", () => OfficeImoReadDataTable(workbookBytes, dataRange, ExecutionMode.Parallel))
        ], warmupIterations, measuredIterations));
        scenarios.AddRange(MeasureGroup("OfficeIMO.Excel", [
            new ReadProfileCase("ReadRangeStream", "Streaming row chunks with automatic execution policy.", () => OfficeImoReadRangeStream(workbookBytes, dataRange, null)),
            new ReadProfileCase("ReadRangeStream.Sequential", "Streaming row chunks with forced sequential conversion.", () => OfficeImoReadRangeStream(workbookBytes, dataRange, ExecutionMode.Sequential)),
            new ReadProfileCase("ReadRangeStream.Parallel", "Streaming row chunks with forced parallel conversion.", () => OfficeImoReadRangeStream(workbookBytes, dataRange, ExecutionMode.Parallel))
        ], warmupIterations, measuredIterations));
        scenarios.Add(Measure("OfficeIMO.Excel", "ReadColumn.LargeSparse", "Sparse A1:A100001 read with only the first and last rows populated.", () => OfficeImoReadSparseColumn(sparseWorkbookBytes, sparseRange, SparseLastRow), warmupIterations, measuredIterations));
        scenarios.Add(Measure("OfficeIMO.Excel", "ReadRows.LargeSparse", "Sparse A1:A100001 row read with only the first and last rows populated.", () => OfficeImoReadSparseRows(sparseWorkbookBytes, sparseRange, SparseLastRow), warmupIterations, measuredIterations));
        scenarios.Add(Measure("ClosedXML", "ReadRows", "Worksheet row iteration over the same workbook payload.", () => ClosedXmlReadRows(workbookBytes), warmupIterations, measuredIterations));

        var profile = new ExcelReadProfile {
            GeneratedAtUtc = DateTime.UtcNow,
            Framework = System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription,
            MachineName = Environment.MachineName,
            BuildConfiguration = BuildConfiguration,
            RowCount = rowCount,
            WarmupIterations = warmupIterations,
            MeasuredIterations = measuredIterations,
            Scenarios = scenarios
        };

        string? directory = Path.GetDirectoryName(outputPath);
        if (!string.IsNullOrEmpty(directory)) {
            Directory.CreateDirectory(directory);
        }

        var options = new JsonSerializerOptions { WriteIndented = true };
        File.WriteAllText(outputPath, JsonSerializer.Serialize(profile, options));
        return outputPath;
    }

    private static ExcelReadProfileScenario Measure(
        string library,
        string name,
        string notes,
        Func<int> action,
        int warmupIterations,
        int measuredIterations) {
        var measurement = BenchmarkMeasurement.Measure(warmupIterations, measuredIterations, action);

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

    private static IReadOnlyList<ExcelReadProfileScenario> MeasureGroup(
        string library,
        IReadOnlyList<ReadProfileCase> cases,
        int warmupIterations,
        int measuredIterations) {
        if (cases.Count == 0) {
            return [];
        }

        var samples = new List<double>[cases.Count];
        var outputMetrics = new int[cases.Count];
        for (int i = 0; i < cases.Count; i++) {
            if (cases[i].Action == null) {
                throw new ArgumentNullException(nameof(cases));
            }

            samples[i] = new List<double>(measuredIterations);
        }

        for (int warmup = 0; warmup < warmupIterations; warmup++) {
            foreach (int index in GetRotatedOrder(cases.Count, warmup)) {
                cases[index].Action();
            }
        }

        for (int iteration = 0; iteration < measuredIterations; iteration++) {
            foreach (int index in GetRotatedOrder(cases.Count, iteration)) {
                BenchmarkMeasurement.PrepareForMeasurement();

                var stopwatch = Stopwatch.StartNew();
                outputMetrics[index] = cases[index].Action();
                stopwatch.Stop();

                samples[index].Add(stopwatch.Elapsed.TotalMilliseconds);
            }
        }

        var scenarios = new List<ExcelReadProfileScenario>(cases.Count);
        for (int i = 0; i < cases.Count; i++) {
            var measurement = new BenchmarkMeasurement.BenchmarkMeasurementResult(outputMetrics[i], samples[i]);
            scenarios.Add(new ExcelReadProfileScenario {
                Library = library,
                Name = cases[i].Name,
                Notes = cases[i].Notes,
                OutputMetric = measurement.OutputMetric,
                AverageMilliseconds = measurement.AverageMilliseconds,
                MedianMilliseconds = measurement.MedianMilliseconds,
                SamplesMilliseconds = measurement.SamplesMilliseconds.ToList()
            });
        }

        return scenarios;
    }

    private static IEnumerable<int> GetRotatedOrder(int count, int offset) {
        for (int i = 0; i < count; i++) {
            yield return (i + offset) % count;
        }
    }

    private static int OfficeImoReadObjects(byte[] workbookBytes, string dataRange, ExecutionMode? mode) {
        using var reader = ExcelDocumentReader.Open(workbookBytes);
        return reader.GetSheet("Data").ReadObjects(dataRange, mode).Count();
    }

    private static int OfficeImoReadObjectsAs(byte[] workbookBytes, string dataRange, ExecutionMode? mode) {
        using var reader = ExcelDocumentReader.Open(workbookBytes);
        return reader.GetSheet("Data").ReadObjects<ReadSalesRecord>(dataRange, mode).Count();
    }

    private static int OfficeImoReadRange(byte[] workbookBytes, string dataRange, ExecutionMode? mode) {
        using var reader = ExcelDocumentReader.Open(workbookBytes);
        var values = reader.GetSheet("Data").ReadRange(dataRange, mode);
        return values.GetLength(0) * values.GetLength(1);
    }

    private static int OfficeImoReadDataTable(byte[] workbookBytes, string dataRange, ExecutionMode? mode) {
        using var reader = ExcelDocumentReader.Open(workbookBytes);
        return reader.GetSheet("Data").ReadRangeAsDataTable(dataRange, headersInFirstRow: true, mode: mode).Rows.Count;
    }

    private static int OfficeImoReadRangeStream(byte[] workbookBytes, string dataRange, ExecutionMode? mode) {
        using var reader = ExcelDocumentReader.Open(workbookBytes);
        int rows = 0;
        foreach (var chunk in reader.GetSheet("Data").ReadRangeStream(dataRange, chunkRows: 512, mode: mode)) {
            rows += chunk.RowCount;
        }

        return rows;
    }

    private static int OfficeImoReadSparseColumn(byte[] workbookBytes, string sparseRange, int expectedRows) {
        using var reader = ExcelDocumentReader.Open(workbookBytes);
        int rowIndex = 0;

        foreach (object? value in reader.GetSheet("Data").ReadColumn(sparseRange)) {
            rowIndex++;
            ValidateSparseCell(rowIndex, expectedRows, value);
        }

        if (rowIndex != expectedRows) {
            throw new InvalidOperationException($"Expected {expectedRows} sparse column rows, got {rowIndex}.");
        }

        return rowIndex;
    }

    private static int OfficeImoReadSparseRows(byte[] workbookBytes, string sparseRange, int expectedRows) {
        using var reader = ExcelDocumentReader.Open(workbookBytes);
        int rowIndex = 0;

        foreach (object?[]? row in reader.GetSheet("Data").ReadRows(sparseRange)) {
            rowIndex++;
            ValidateSparseCell(rowIndex, expectedRows, row?[0]);
        }

        if (rowIndex != expectedRows) {
            throw new InvalidOperationException($"Expected {expectedRows} sparse rows, got {rowIndex}.");
        }

        return rowIndex;
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

    private static byte[] CreateSparseWorkbookBytes(int lastRow) {
        using var stream = new MemoryStream();
        using (var document = ExcelDocument.Create(stream)) {
            var sheet = document.AddWorkSheet("Data");
            sheet.CellValue(1, 1, "Header");
            sheet.CellValue(lastRow, 1, "Tail");
        }

        return stream.ToArray();
    }

    private static void ValidateSparseCell(int rowIndex, int expectedRows, object? value) {
        if (rowIndex == 1) {
            if (!Equals("Header", value)) {
                throw new InvalidOperationException("Sparse read did not return the first row value.");
            }

            return;
        }

        if (rowIndex == expectedRows) {
            if (!Equals("Tail", value)) {
                throw new InvalidOperationException("Sparse read did not return the last row value.");
            }

            return;
        }

        if (value != null) {
            throw new InvalidOperationException($"Sparse read returned an unexpected value at row {rowIndex}.");
        }
    }

    private sealed class ExcelReadProfile {
        public DateTime GeneratedAtUtc { get; init; }
        public string Framework { get; init; } = string.Empty;
        public string MachineName { get; init; } = string.Empty;
        public string BuildConfiguration { get; init; } = string.Empty;
        public int RowCount { get; init; }
        public int WarmupIterations { get; init; }
        public int MeasuredIterations { get; init; }
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

    private sealed record ReadProfileCase(string Name, string Notes, Func<int> Action);

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
