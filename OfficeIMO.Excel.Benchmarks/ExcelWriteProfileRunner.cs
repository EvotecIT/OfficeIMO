using System.Diagnostics;
using System.Text.Json;
using ClosedXML.Excel;

namespace OfficeIMO.Excel.Benchmarks;

internal static class ExcelWriteProfileRunner {
    private const int DefaultRowCount = 2500;
    private const int WarmupIterations = 2;
    private const int MeasuredIterations = 5;

    internal static string WriteProfile(string outputPath, int rowCount = DefaultRowCount) {
        if (string.IsNullOrWhiteSpace(outputPath)) {
            throw new ArgumentException("Output path must not be empty.", nameof(outputPath));
        }

        var rows = ExcelBenchmarkScenarioFactory.CreateSalesRecords(rowCount);
        var profile = new ExcelWriteProfile {
            GeneratedAtUtc = DateTime.UtcNow,
            Framework = System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription,
            MachineName = Environment.MachineName,
            RowCount = rowCount,
            Libraries = [
                MeasureOfficeImo(rows),
                MeasureClosedXml(rows)
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

    private static ExcelWriteProfileLibrary MeasureOfficeImo(IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows) {
        int lastOutputBytes = 0;
        var samplesByStage = CreateStageSamples();

        for (int i = 0; i < WarmupIterations; i++) {
            lastOutputBytes = MeasureOfficeImoIteration(rows, totals: null);
        }

        for (int i = 0; i < MeasuredIterations; i++) {
            var iterationTotals = CreateStageAccumulator();
            BenchmarkMeasurement.PrepareForMeasurement();
            lastOutputBytes = MeasureOfficeImoIteration(rows, iterationTotals);
            AddStageSamples(samplesByStage, iterationTotals);
        }

        return BuildLibraryProfile("OfficeIMO.Excel", samplesByStage, lastOutputBytes);
    }

    private static ExcelWriteProfileLibrary MeasureClosedXml(IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows) {
        int lastOutputBytes = 0;
        var samplesByStage = CreateStageSamples();

        for (int i = 0; i < WarmupIterations; i++) {
            lastOutputBytes = MeasureClosedXmlIteration(rows, totals: null);
        }

        for (int i = 0; i < MeasuredIterations; i++) {
            var iterationTotals = CreateStageAccumulator();
            BenchmarkMeasurement.PrepareForMeasurement();
            lastOutputBytes = MeasureClosedXmlIteration(rows, iterationTotals);
            AddStageSamples(samplesByStage, iterationTotals);
        }

        return BuildLibraryProfile("ClosedXML", samplesByStage, lastOutputBytes);
    }

    private static int MeasureOfficeImoIteration(IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows, Dictionary<string, double>? totals) {
        using var stream = new MemoryStream();
        var totalWatch = Stopwatch.StartNew();
        var stageWatch = Stopwatch.StartNew();
        var document = ExcelDocument.Create(stream);

        try {
            var sheet = document.AddWorkSheet("Data");

            stageWatch.Restart();
            ExcelBenchmarkScenarioFactory.InsertOfficeImoObjects(sheet, rows);
            AddStage(totals, "InsertObjects", stageWatch.Elapsed.TotalMilliseconds);

            stageWatch.Restart();
            ExcelBenchmarkScenarioFactory.AddOfficeImoTable(sheet, rows.Count);
            AddStage(totals, "AddTable", stageWatch.Elapsed.TotalMilliseconds);

            stageWatch.Restart();
            ExcelBenchmarkScenarioFactory.AutoFitOfficeImoColumns(sheet);
            AddStage(totals, "AutoFitColumns", stageWatch.Elapsed.TotalMilliseconds);

            stageWatch.Restart();
            document.Dispose();
            AddStage(totals, "DisposeAndSave", stageWatch.Elapsed.TotalMilliseconds);
        } finally {
            GC.SuppressFinalize(document);
        }

        totalWatch.Stop();
        AddStage(totals, "Total", totalWatch.Elapsed.TotalMilliseconds);
        return checked((int)stream.Length);
    }

    private static int MeasureClosedXmlIteration(IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows, Dictionary<string, double>? totals) {
        using var stream = new MemoryStream();
        using var workbook = new XLWorkbook();
        var worksheet = workbook.Worksheets.Add("Data");
        var totalWatch = Stopwatch.StartNew();
        var stageWatch = Stopwatch.StartNew();

        stageWatch.Restart();
        var table = ExcelBenchmarkScenarioFactory.InsertClosedXmlTable(worksheet, rows);
        AddStage(totals, "InsertTable", stageWatch.Elapsed.TotalMilliseconds);

        stageWatch.Restart();
        ExcelBenchmarkScenarioFactory.StyleClosedXmlTable(table);
        AddStage(totals, "StyleTable", stageWatch.Elapsed.TotalMilliseconds);

        stageWatch.Restart();
        ExcelBenchmarkScenarioFactory.AutoFitClosedXmlColumns(worksheet);
        AddStage(totals, "AutoFitColumns", stageWatch.Elapsed.TotalMilliseconds);

        stageWatch.Restart();
        workbook.SaveAs(stream);
        AddStage(totals, "SaveAs", stageWatch.Elapsed.TotalMilliseconds);

        totalWatch.Stop();
        AddStage(totals, "Total", totalWatch.Elapsed.TotalMilliseconds);
        return checked((int)stream.Length);
    }

    private static ExcelWriteProfileLibrary BuildLibraryProfile(string library, Dictionary<string, List<double>> samplesByStage, int lastOutputBytes) {
        return new ExcelWriteProfileLibrary {
            Library = library,
            OutputBytes = lastOutputBytes,
            Stages = samplesByStage
                .Select(stage => new ExcelWriteProfileStage {
                    Name = stage.Key,
                    AverageMilliseconds = stage.Value.Count == 0 ? 0 : stage.Value.Average(),
                    MedianMilliseconds = CalculateMedian(stage.Value),
                    SamplesMilliseconds = stage.Value.ToList()
                })
                .ToList()
        };
    }

    private static Dictionary<string, double> CreateStageAccumulator() {
        return new Dictionary<string, double>(StringComparer.Ordinal);
    }

    private static Dictionary<string, List<double>> CreateStageSamples() {
        return new Dictionary<string, List<double>>(StringComparer.Ordinal);
    }

    private static void AddStageSamples(Dictionary<string, List<double>> totals, Dictionary<string, double> iterationTotals) {
        foreach (var stage in iterationTotals) {
            if (!totals.TryGetValue(stage.Key, out var samples)) {
                samples = [];
                totals[stage.Key] = samples;
            }

            samples.Add(stage.Value);
        }
    }

    private static void AddStage(Dictionary<string, double>? totals, string stageName, double elapsedMilliseconds) {
        if (totals is null) {
            return;
        }

        totals.TryGetValue(stageName, out double currentValue);
        totals[stageName] = currentValue + elapsedMilliseconds;
    }

    private static double CalculateMedian(IReadOnlyList<double> samples) {
        if (samples.Count == 0) {
            return 0;
        }

        var ordered = samples.OrderBy(v => v).ToArray();
        int middle = ordered.Length / 2;
        if ((ordered.Length & 1) == 1) {
            return ordered[middle];
        }

        return (ordered[middle - 1] + ordered[middle]) / 2.0;
    }

    private sealed class ExcelWriteProfile {
        public DateTime GeneratedAtUtc { get; init; }
        public string Framework { get; init; } = string.Empty;
        public string MachineName { get; init; } = string.Empty;
        public int RowCount { get; init; }
        public List<ExcelWriteProfileLibrary> Libraries { get; init; } = [];
    }

    private sealed class ExcelWriteProfileLibrary {
        public string Library { get; init; } = string.Empty;
        public int OutputBytes { get; init; }
        public List<ExcelWriteProfileStage> Stages { get; init; } = [];
    }

    private sealed class ExcelWriteProfileStage {
        public string Name { get; init; } = string.Empty;
        public double AverageMilliseconds { get; init; }
        public double MedianMilliseconds { get; init; }
        public List<double> SamplesMilliseconds { get; init; } = [];
    }
}
