using System.Diagnostics;
using System.Text.Json;
using OfficeIMO.Drawing;

namespace OfficeIMO.Excel.Benchmarks;

internal static class ExcelChartProfileRunner {
    internal const int DefaultWarmupIterations = 1;
    internal const int DefaultMeasuredIterations = 3;

    internal static string WriteProfile(string outputPath, int rowCount, int warmupIterations, int measuredIterations) {
        if (string.IsNullOrWhiteSpace(outputPath)) {
            throw new ArgumentException("Output path must not be empty.", nameof(outputPath));
        }

        var rows = ExcelBenchmarkScenarioFactory.CreateSalesRecords(rowCount);
        var samplesByStage = new Dictionary<string, List<double>>(StringComparer.Ordinal);
        var diagnostics = new List<ChartProfileDiagnostics>();
        int outputBytes = 0;

        for (int i = 0; i < warmupIterations; i++) {
            outputBytes = MeasureOfficeImoIteration(rows, null, null);
        }

        for (int i = 0; i < measuredIterations; i++) {
            var totals = new Dictionary<string, double>(StringComparer.Ordinal);
            BenchmarkMeasurement.PrepareForMeasurement();
            outputBytes = MeasureOfficeImoIteration(rows, totals, diagnostics);
            foreach (var stage in totals) {
                if (!samplesByStage.TryGetValue(stage.Key, out var samples)) {
                    samples = [];
                    samplesByStage[stage.Key] = samples;
                }

                samples.Add(stage.Value);
            }
        }

        var profile = new ChartProfile {
            GeneratedAtUtc = DateTime.UtcNow,
            Framework = System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription,
            MachineName = Environment.MachineName,
            RowCount = rowCount,
            WarmupIterations = warmupIterations,
            MeasuredIterations = measuredIterations,
            OutputBytes = outputBytes,
            Diagnostics = diagnostics,
            Stages = samplesByStage
                .Select(stage => new ChartProfileStage {
                    Name = stage.Key,
                    AverageMilliseconds = stage.Value.Average(),
                    MedianMilliseconds = CalculateMedian(stage.Value),
                    SamplesMilliseconds = stage.Value.ToList()
                })
                .OrderBy(stage => stage.Name, StringComparer.Ordinal)
                .ToList()
        };

        string? directory = Path.GetDirectoryName(outputPath);
        if (!string.IsNullOrEmpty(directory)) {
            Directory.CreateDirectory(directory);
        }

        File.WriteAllText(outputPath, JsonSerializer.Serialize(profile, new JsonSerializerOptions { WriteIndented = true }));
        return outputPath;
    }

    private static int MeasureOfficeImoIteration(
        IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows,
        Dictionary<string, double>? totals,
        List<ChartProfileDiagnostics>? diagnostics) {
        using var stream = new MemoryStream();
        using var document = ExcelDocument.Create(stream, autoSave: false);
        var totalWatch = Stopwatch.StartNew();
        var stageWatch = Stopwatch.StartNew();

        document.Execution.SaveWorksheetAfterAutoFit = false;
        if (totals != null) {
            document.Execution.OnTiming = (operation, elapsed) => AddStage(totals, operation, elapsed.TotalMilliseconds);
        }

        stageWatch.Restart();
        var sheet = document.AddWorkSheet("Data");
        AddStage(totals, "AddWorksheet", stageWatch.Elapsed.TotalMilliseconds);

        stageWatch.Restart();
        ExcelBenchmarkScenarioFactory.InsertOfficeImoObjects(sheet, rows);
        AddStage(totals, "InsertObjects", stageWatch.Elapsed.TotalMilliseconds);

        stageWatch.Restart();
        var summaries = BuildRegionSummaries(rows);
        AddStage(totals, "BuildRegionSummaries", stageWatch.Elapsed.TotalMilliseconds);

        stageWatch.Restart();
        var chartData = new ExcelChartData(
            summaries.Select(static item => item.Region),
            new[] {
                new ExcelChartSeries("Amount", summaries.Select(static item => item.Amount)),
                new ExcelChartSeries("Units", summaries.Select(static item => (double)item.Units))
            });
        AddStage(totals, "BuildChartData", stageWatch.Elapsed.TotalMilliseconds);

        stageWatch.Restart();
        sheet.AddChart(chartData, row: 18, column: 10, widthPixels: 720, heightPixels: 360, type: ExcelChartType.ColumnClustered, title: "Regional Sales");
        AddStage(totals, "AddChart", stageWatch.Elapsed.TotalMilliseconds);

        stageWatch.Restart();
        document.Save(stream);
        AddStage(totals, "Save", stageWatch.Elapsed.TotalMilliseconds);

        totalWatch.Stop();
        AddStage(totals, "Total", totalWatch.Elapsed.TotalMilliseconds);
        diagnostics?.Add(new ChartProfileDiagnostics {
            Writer = document.LastSaveDiagnostics.Writer.ToString(),
            UsedFastPackageWriter = document.LastSaveDiagnostics.UsedFastPackageWriter,
            FastPackageSkipReason = document.LastSaveDiagnostics.FastPackageSkipReason
        });

        return checked((int)stream.Length);
    }

    private static IReadOnlyList<RegionSummary> BuildRegionSummaries(IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows)
        => rows
            .GroupBy(static row => row.Region, StringComparer.Ordinal)
            .OrderBy(static group => group.Key, StringComparer.Ordinal)
            .Select(static group => new RegionSummary(
                group.Key,
                Math.Round(group.Sum(static row => row.Amount), 2),
                group.Sum(static row => row.Units)))
            .ToList();

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

        var ordered = samples.OrderBy(value => value).ToArray();
        int middle = ordered.Length / 2;
        return (ordered.Length & 1) == 1
            ? ordered[middle]
            : (ordered[middle - 1] + ordered[middle]) / 2.0;
    }

    private sealed record RegionSummary(string Region, double Amount, int Units);

    private sealed class ChartProfile {
        public DateTime GeneratedAtUtc { get; init; }
        public string Framework { get; init; } = string.Empty;
        public string MachineName { get; init; } = string.Empty;
        public int RowCount { get; init; }
        public int WarmupIterations { get; init; }
        public int MeasuredIterations { get; init; }
        public int OutputBytes { get; init; }
        public List<ChartProfileDiagnostics> Diagnostics { get; init; } = [];
        public List<ChartProfileStage> Stages { get; init; } = [];
    }

    private sealed class ChartProfileDiagnostics {
        public string Writer { get; init; } = string.Empty;
        public bool UsedFastPackageWriter { get; init; }
        public string? FastPackageSkipReason { get; init; }
    }

    private sealed class ChartProfileStage {
        public string Name { get; init; } = string.Empty;
        public double AverageMilliseconds { get; init; }
        public double MedianMilliseconds { get; init; }
        public List<double> SamplesMilliseconds { get; init; } = [];
    }
}
