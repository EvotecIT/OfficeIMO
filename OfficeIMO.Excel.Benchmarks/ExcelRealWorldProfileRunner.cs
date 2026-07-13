using System.Diagnostics;
using System.Globalization;
using System.Text.Json;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Drawing;

namespace OfficeIMO.Excel.Benchmarks;

internal static class ExcelRealWorldProfileRunner {
    internal const int DefaultWarmupIterations = 1;
    internal const int DefaultMeasuredIterations = 3;

    internal static string WriteProfile(string outputPath, int rowCount, int warmupIterations, int measuredIterations) {
        if (string.IsNullOrWhiteSpace(outputPath)) {
            throw new ArgumentException("Output path must not be empty.", nameof(outputPath));
        }

        var rows = ExcelBenchmarkScenarioFactory.CreateSalesRecords(rowCount);
        var samplesByStage = new Dictionary<string, List<double>>(StringComparer.Ordinal);
        var diagnostics = new List<RealWorldProfileDiagnostics>();
        int outputBytes = 0;

        for (int i = 0; i < warmupIterations; i++) {
            outputBytes = MeasureIteration(rows, null, null);
        }

        for (int i = 0; i < measuredIterations; i++) {
            var totals = new Dictionary<string, double>(StringComparer.Ordinal);
            BenchmarkMeasurement.PrepareForMeasurement();
            outputBytes = MeasureIteration(rows, totals, diagnostics);
            foreach (var stage in totals) {
                if (!samplesByStage.TryGetValue(stage.Key, out var samples)) {
                    samples = [];
                    samplesByStage[stage.Key] = samples;
                }

                samples.Add(stage.Value);
            }
        }

        var profile = new RealWorldProfile {
            GeneratedAtUtc = DateTime.UtcNow,
            Framework = System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription,
            MachineName = Environment.MachineName,
            RowCount = rowCount,
            WarmupIterations = warmupIterations,
            MeasuredIterations = measuredIterations,
            OutputBytes = outputBytes,
            Diagnostics = diagnostics,
            Stages = samplesByStage
                .Select(stage => new RealWorldProfileStage {
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

    private static int MeasureIteration(
        IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows,
        Dictionary<string, double>? totals,
        List<RealWorldProfileDiagnostics>? diagnostics) {
        using var stream = new MemoryStream();
        using var document = ExcelDocument.Create(stream);
        var totalWatch = Stopwatch.StartNew();
        var stageWatch = Stopwatch.StartNew();

        document.Execution.SaveWorksheetAfterAutoFit = false;
        if (totals != null) {
            document.Execution.OnTiming = (operation, elapsed) => AddStage(totals, operation, elapsed.TotalMilliseconds);
        }

        stageWatch.Restart();
        var sheet = document.AddWorksheet("Data");
        AddStage(totals, "AddWorksheet", stageWatch.Elapsed.TotalMilliseconds);

        stageWatch.Restart();
        ExcelBenchmarkScenarioFactory.InsertOfficeImoObjects(sheet, rows);
        AddStage(totals, "InsertObjects", stageWatch.Elapsed.TotalMilliseconds);

        stageWatch.Restart();
        ApplyTable(sheet, rows.Count);
        AddStage(totals, "TableAndAutoFit", stageWatch.Elapsed.TotalMilliseconds);

        stageWatch.Restart();
        ApplyNavigation(sheet, rows.Count);
        AddStage(totals, "Navigation", stageWatch.Elapsed.TotalMilliseconds);

        stageWatch.Restart();
        ApplyConditionalFormatting(sheet, rows.Count);
        AddStage(totals, "ConditionalFormatting", stageWatch.Elapsed.TotalMilliseconds);

        stageWatch.Restart();
        ApplyDataValidation(sheet, rows.Count);
        AddStage(totals, "DataValidation", stageWatch.Elapsed.TotalMilliseconds);

        stageWatch.Restart();
        AddPivotTable(sheet, rows.Count);
        AddStage(totals, "AddPivotTable", stageWatch.Elapsed.TotalMilliseconds);

        stageWatch.Restart();
        AddRegionalChart(sheet, rows);
        AddStage(totals, "AddChart", stageWatch.Elapsed.TotalMilliseconds);

        stageWatch.Restart();
        document.Save(stream);
        AddStage(totals, "Save", stageWatch.Elapsed.TotalMilliseconds);

        totalWatch.Stop();
        AddStage(totals, "Total", totalWatch.Elapsed.TotalMilliseconds);
        diagnostics?.Add(new RealWorldProfileDiagnostics {
            Writer = document.LastSaveDiagnostics.Writer.ToString(),
            UsedFastPackageWriter = document.LastSaveDiagnostics.UsedFastPackageWriter,
            FastPackageSkipReason = document.LastSaveDiagnostics.FastPackageSkipReason
        });

        return checked((int)stream.Length);
    }

    private static void ApplyTable(ExcelSheet sheet, int rowCount) {
        sheet.AddTable(BuildSalesRange(rowCount), hasHeader: true, name: "SalesData", style: TableStyle.TableStyleMedium2, includeAutoFilter: true);
        sheet.AutoFitColumns();
    }

    private static void ApplyNavigation(ExcelSheet sheet, int rowCount) {
        sheet.Freeze(topRows: 1, leftCols: 1);
        sheet.AddAutoFilter(BuildSalesRange(rowCount));
    }

    private static void ApplyConditionalFormatting(ExcelSheet sheet, int rowCount) {
        int lastRow = rowCount + 1;
        sheet.AddConditionalRule($"E2:E{lastRow.ToString(CultureInfo.InvariantCulture)}", ConditionalFormattingOperatorValues.GreaterThan, "3000");
        sheet.AddConditionalRule($"F2:F{lastRow.ToString(CultureInfo.InvariantCulture)}", ConditionalFormattingOperatorValues.LessThan, "5");
        sheet.AddConditionalColorScale($"E2:E{lastRow.ToString(CultureInfo.InvariantCulture)}", OfficeColor.LightPink, OfficeColor.LightGreen);
        sheet.AddConditionalDataBar($"F2:F{lastRow.ToString(CultureInfo.InvariantCulture)}", OfficeColor.SteelBlue);
    }

    private static void ApplyDataValidation(ExcelSheet sheet, int rowCount) {
        int lastRow = rowCount + 1;
        sheet.ValidationWholeNumber($"F2:F{lastRow.ToString(CultureInfo.InvariantCulture)}", DataValidationOperatorValues.Between, 1, 24);
    }

    private static void AddPivotTable(ExcelSheet sheet, int rowCount) {
        sheet.AddPivotTable(
            sourceRange: BuildSalesRange(rowCount),
            destinationCell: "J3",
            name: "SalesPivot",
            rowFields: new[] { "Region" },
            columnFields: new[] { "Owner" },
            dataFields: new[] { new ExcelPivotDataField("Amount", DataConsolidateFunctionValues.Sum, "Total Amount") },
            pivotStyleName: "PivotStyleMedium9");
    }

    private static void AddRegionalChart(ExcelSheet sheet, IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows) {
        var summaries = rows
            .GroupBy(static row => row.Region, StringComparer.Ordinal)
            .OrderBy(static group => group.Key, StringComparer.Ordinal)
            .Select(static group => new {
                Region = group.Key,
                Amount = Math.Round(group.Sum(static row => row.Amount), 2),
                Units = group.Sum(static row => row.Units)
            })
            .ToArray();

        var chartData = new ExcelChartData(
            summaries.Select(static item => item.Region),
            new[] {
                new ExcelChartSeries("Amount", summaries.Select(static item => item.Amount)),
                new ExcelChartSeries("Units", summaries.Select(static item => (double)item.Units))
            });

        sheet.AddChart(chartData, row: 18, column: 10, widthPixels: 720, heightPixels: 360, type: ExcelChartType.ColumnClustered, title: "Regional Sales");
    }

    private static string BuildSalesRange(int rowCount)
        => "A1:H" + (rowCount + 1).ToString(CultureInfo.InvariantCulture);

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

    private sealed class RealWorldProfile {
        public DateTime GeneratedAtUtc { get; init; }
        public string Framework { get; init; } = string.Empty;
        public string MachineName { get; init; } = string.Empty;
        public int RowCount { get; init; }
        public int WarmupIterations { get; init; }
        public int MeasuredIterations { get; init; }
        public int OutputBytes { get; init; }
        public List<RealWorldProfileDiagnostics> Diagnostics { get; init; } = [];
        public List<RealWorldProfileStage> Stages { get; init; } = [];
    }

    private sealed class RealWorldProfileDiagnostics {
        public string Writer { get; init; } = string.Empty;
        public bool UsedFastPackageWriter { get; init; }
        public string? FastPackageSkipReason { get; init; }
    }

    private sealed class RealWorldProfileStage {
        public string Name { get; init; } = string.Empty;
        public double AverageMilliseconds { get; init; }
        public double MedianMilliseconds { get; init; }
        public List<double> SamplesMilliseconds { get; init; } = [];
    }
}
