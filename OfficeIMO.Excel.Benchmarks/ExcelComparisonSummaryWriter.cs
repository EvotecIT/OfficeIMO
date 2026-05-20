using System.Globalization;
using System.Text;
using System.Text.Json;

namespace OfficeIMO.Excel.Benchmarks;

internal sealed record ExcelComparisonSummaryInput(string Kind, int RowCount, string Path);

internal sealed record ExcelComparisonSummaryOutput(string MarkdownPath, string CsvPath, string JsonPath);

internal static class ExcelComparisonSummaryWriter {
    private const double TieThresholdRatio = 0.02;

    internal static ExcelComparisonSummaryOutput WriteSummary(
        string outputDirectory,
        IEnumerable<ExcelComparisonSummaryInput> artifacts) {
        if (string.IsNullOrWhiteSpace(outputDirectory)) {
            throw new ArgumentException("Output directory must not be empty.", nameof(outputDirectory));
        }

        var rows = new List<ComparisonSummaryRow>();
        foreach (var artifact in artifacts) {
            if (!File.Exists(artifact.Path) || IsSummaryArtifact(artifact.Kind)) {
                continue;
            }

            rows.AddRange(ReadArtifactRows(artifact));
        }

        ApplyScenarioComparisons(rows);

        Directory.CreateDirectory(outputDirectory);
        string jsonPath = Path.Combine(outputDirectory, "officeimo.excel.comparison-summary.json");
        string csvPath = Path.Combine(outputDirectory, "officeimo.excel.comparison-summary.csv");
        string markdownPath = Path.Combine(outputDirectory, "officeimo.excel.comparison-summary.md");

        File.WriteAllText(jsonPath, JsonSerializer.Serialize(new ComparisonSummaryDocument {
            GeneratedAtUtc = DateTime.UtcNow,
            Notes = "Mean/stddev/stderr are from the lightweight rotated local runner. Allocations use GC.GetAllocatedBytesForCurrentThread. Use BenchmarkDotNet-specific classes for publication-grade statistical Error columns.",
            Rows = rows
        }, new JsonSerializerOptions { WriteIndented = true }));
        File.WriteAllText(csvPath, BuildCsv(rows));
        File.WriteAllText(markdownPath, BuildMarkdown(rows));

        return new ExcelComparisonSummaryOutput(markdownPath, csvPath, jsonPath);
    }

    private static bool IsSummaryArtifact(string kind)
        => kind.Contains("summary", StringComparison.OrdinalIgnoreCase);

    private static IEnumerable<ComparisonSummaryRow> ReadArtifactRows(ExcelComparisonSummaryInput artifact) {
        using var document = JsonDocument.Parse(File.ReadAllText(artifact.Path));
        if (!document.RootElement.TryGetProperty("Scenarios", out JsonElement scenarios)
            || scenarios.ValueKind != JsonValueKind.Array) {
            yield break;
        }

        foreach (JsonElement scenario in scenarios.EnumerateArray()) {
            string scenarioName = GetString(scenario, "Scenario");
            string library = GetString(scenario, "Library");
            var row = new ComparisonSummaryRow {
                ArtifactKind = artifact.Kind,
                RowCount = artifact.RowCount,
                Workload = GetWorkloadKind(scenarioName, artifact.Kind),
                Scenario = scenarioName,
                Library = library,
                Notes = GetString(scenario, "Notes"),
                OutputMetric = GetNullableInt(scenario, "OutputMetric"),
                MeanMilliseconds = GetDouble(scenario, "AverageMilliseconds"),
                MedianMilliseconds = GetDouble(scenario, "MedianMilliseconds"),
                StandardDeviationMilliseconds = GetNullableDouble(scenario, "StandardDeviationMilliseconds")
                    ?? CalculateStandardDeviation(GetDoubleArray(scenario, "SamplesMilliseconds")),
                StandardErrorMilliseconds = GetNullableDouble(scenario, "StandardErrorMilliseconds")
                    ?? CalculateStandardError(GetDoubleArray(scenario, "SamplesMilliseconds")),
                MeanAllocatedBytes = GetNullableDouble(scenario, "AverageAllocatedBytes"),
                MedianAllocatedBytes = GetNullableDouble(scenario, "MedianAllocatedBytes"),
                ArtifactPath = artifact.Path
            };

            if (scenario.TryGetProperty("Package", out JsonElement package)
                && package.ValueKind == JsonValueKind.Object) {
                row.PackageBytes = GetNullableLong(package, "FileSizeBytes");
                row.WorksheetCompressedBytes = GetNullableLong(package, "WorksheetCompressedBytes");
                row.SharedStringsCompressedBytes = GetNullableLong(package, "SharedStringsCompressedBytes");
                row.StylesCompressedBytes = GetNullableLong(package, "StylesCompressedBytes");
                row.TablesCompressedBytes = GetNullableLong(package, "TablesCompressedBytes");
                row.WorksheetCellCount = GetNullableInt(package, "WorksheetCellCount");
                row.WorksheetRowCount = GetNullableInt(package, "WorksheetRowCount");
                row.SharedStringCount = GetNullableInt(package, "SharedStringCount");
                row.UniqueSharedStringCount = GetNullableInt(package, "UniqueSharedStringCount");
                row.CellStyleCount = GetNullableInt(package, "CellStyleCount");
            }

            yield return row;
        }
    }

    private static void ApplyScenarioComparisons(List<ComparisonSummaryRow> rows) {
        foreach (var group in rows.GroupBy(row => (row.ArtifactKind, row.RowCount, row.Scenario))) {
            double bestMean = group.Min(row => row.MeanMilliseconds);
            var bestRows = group
                .Where(row => IsTie(row.MeanMilliseconds, bestMean))
                .OrderBy(row => row.MeanMilliseconds)
                .ThenBy(row => row.Library, StringComparer.OrdinalIgnoreCase)
                .ToArray();
            string bestLibraries = string.Join(", ", bestRows.Select(row => row.Library));
            var office = group.FirstOrDefault(row => IsOfficeImo(row.Library));

            double? officeMean = office?.MeanMilliseconds;
            double? officeAllocated = office?.MeanAllocatedBytes;
            long? officePackageBytes = office?.PackageBytes;

            foreach (var row in group) {
                row.BestLibrary = bestLibraries;
                row.BestMeanMilliseconds = bestMean;
                row.RatioToBest = bestMean <= 0 ? null : row.MeanMilliseconds / bestMean;
                row.RatioToOfficeImo = officeMean is > 0 ? row.MeanMilliseconds / officeMean.Value : null;
                row.AllocatedRatioToOfficeImo = officeAllocated is > 0 && row.MeanAllocatedBytes is > 0
                    ? row.MeanAllocatedBytes.Value / officeAllocated.Value
                    : null;
                row.PackageRatioToOfficeImo = officePackageBytes is > 0 && row.PackageBytes is > 0
                    ? (double)row.PackageBytes.Value / officePackageBytes.Value
                    : null;
                row.Outcome = GetOutcome(row, bestMean, officeMean);
            }
        }
    }

    private static string GetOutcome(ComparisonSummaryRow row, double bestMean, double? officeMean) {
        if (IsOfficeImo(row.Library)) {
            if (IsTie(row.MeanMilliseconds, bestMean)) {
                return "Win";
            }

            double lossPercent = bestMean <= 0 ? 0 : ((row.MeanMilliseconds / bestMean) - 1) * 100;
            return string.Create(CultureInfo.InvariantCulture, $"Loss +{lossPercent:F1}%");
        }

        if (officeMean is not > 0) {
            return string.Empty;
        }

        if (IsTie(row.MeanMilliseconds, officeMean.Value)) {
            return "Tie vs OfficeIMO";
        }

        if (row.MeanMilliseconds < officeMean.Value) {
            double fasterPercent = (1 - (row.MeanMilliseconds / officeMean.Value)) * 100;
            return string.Create(CultureInfo.InvariantCulture, $"{fasterPercent:F1}% faster than OfficeIMO");
        }

        double slowerPercent = ((row.MeanMilliseconds / officeMean.Value) - 1) * 100;
        return string.Create(CultureInfo.InvariantCulture, $"{slowerPercent:F1}% slower than OfficeIMO");
    }

    private static bool IsTie(double value, double baseline)
        => baseline <= 0 || Math.Abs(value - baseline) / baseline <= TieThresholdRatio;

    private static bool IsOfficeImo(string library)
        => string.Equals(library, "OfficeIMO.Excel", StringComparison.OrdinalIgnoreCase);

    private static string GetWorkloadKind(string scenario, string artifactKind) {
        if (artifactKind.Contains("package", StringComparison.OrdinalIgnoreCase)) {
            return "package";
        }

        if (scenario.StartsWith("write-", StringComparison.OrdinalIgnoreCase)
            || string.Equals(scenario, "append-plain-rows", StringComparison.OrdinalIgnoreCase)
            || string.Equals(scenario, "large-shared-strings", StringComparison.OrdinalIgnoreCase)) {
            return "write";
        }

        if (scenario.Contains("read", StringComparison.OrdinalIgnoreCase)) {
            return "read";
        }

        if (scenario.Contains("autofit", StringComparison.OrdinalIgnoreCase)) {
            return "mutate";
        }

        return "other";
    }

    private static string BuildCsv(IReadOnlyList<ComparisonSummaryRow> rows) {
        var builder = new StringBuilder();
        builder.AppendLine(string.Join(",", CsvHeaders.Select(EscapeCsv)));
        foreach (var row in rows.OrderBy(row => row.RowCount).ThenBy(row => row.ArtifactKind).ThenBy(row => row.Scenario).ThenBy(row => row.MeanMilliseconds)) {
            builder.AppendLine(string.Join(",", GetCsvValues(row).Select(EscapeCsv)));
        }

        return builder.ToString();
    }

    private static string BuildMarkdown(IReadOnlyList<ComparisonSummaryRow> rows) {
        var builder = new StringBuilder();
        builder.AppendLine("# OfficeIMO.Excel Comparison Summary");
        builder.AppendLine();
        builder.AppendLine("This is the suite-level decision table. Mean, standard deviation, standard error, ratios, and allocations come from the lightweight rotated local runner; they are meant for engineering direction. Use the BenchmarkDotNet benchmark classes when a publication-grade `Error` column is required.");
        builder.AppendLine();

        AppendAtAGlance(builder, rows);
        AppendDecisionTable(builder, rows);
        AppendFullTable(builder, rows);
        return builder.ToString();
    }

    private static void AppendAtAGlance(StringBuilder builder, IReadOnlyList<ComparisonSummaryRow> rows) {
        var officeRows = rows.Where(row => IsOfficeImo(row.Library)).ToArray();
        builder.AppendLine("## At a glance");
        builder.AppendLine();
        builder.AppendLine("| Row count | Artifact | Workload | OfficeIMO wins | OfficeIMO losses | Biggest loss |");
        builder.AppendLine("| ---: | --- | --- | ---: | ---: | --- |");

        foreach (var group in officeRows.GroupBy(row => (row.RowCount, row.ArtifactKind, row.Workload)).OrderBy(group => group.Key.RowCount).ThenBy(group => group.Key.ArtifactKind).ThenBy(group => group.Key.Workload)) {
            int wins = group.Count(row => row.Outcome == "Win");
            var losses = group.Where(row => row.Outcome.StartsWith("Loss", StringComparison.Ordinal)).ToArray();
            var biggestLoss = losses
                .OrderByDescending(row => row.RatioToBest ?? 0)
                .FirstOrDefault();

            builder.Append("| ");
            builder.Append(group.Key.RowCount.ToString(CultureInfo.InvariantCulture));
            builder.Append(" | ");
            builder.Append(EscapeMarkdown(group.Key.ArtifactKind));
            builder.Append(" | ");
            builder.Append(EscapeMarkdown(group.Key.Workload));
            builder.Append(" | ");
            builder.Append(wins.ToString(CultureInfo.InvariantCulture));
            builder.Append(" | ");
            builder.Append(losses.Length.ToString(CultureInfo.InvariantCulture));
            builder.Append(" | ");
            builder.Append(biggestLoss == null
                ? string.Empty
                : EscapeMarkdown($"{biggestLoss.Scenario}: {biggestLoss.Outcome} vs {biggestLoss.BestLibrary}"));
            builder.AppendLine(" |");
        }

        builder.AppendLine();
    }

    private static void AppendDecisionTable(StringBuilder builder, IReadOnlyList<ComparisonSummaryRow> rows) {
        var officeRows = rows.Where(row => IsOfficeImo(row.Library)).ToArray();
        builder.AppendLine("## OfficeIMO decision table");
        builder.AppendLine();
        builder.AppendLine("| Row count | Artifact | Workload | Scenario | OfficeIMO mean | Best | OfficeIMO vs best | Alloc | Package |");
        builder.AppendLine("| ---: | --- | --- | --- | ---: | --- | ---: | ---: | ---: |");

        foreach (var row in officeRows.OrderBy(row => row.RowCount).ThenBy(row => row.ArtifactKind).ThenBy(row => row.Workload).ThenBy(row => row.Scenario)) {
            builder.Append("| ");
            builder.Append(row.RowCount.ToString(CultureInfo.InvariantCulture));
            builder.Append(" | ");
            builder.Append(EscapeMarkdown(row.ArtifactKind));
            builder.Append(" | ");
            builder.Append(EscapeMarkdown(row.Workload));
            builder.Append(" | ");
            builder.Append(EscapeMarkdown(row.Scenario));
            builder.Append(" | ");
            builder.Append(FormatMilliseconds(row.MeanMilliseconds));
            builder.Append(" | ");
            builder.Append(EscapeMarkdown(row.BestLibrary));
            builder.Append(" | ");
            builder.Append(EscapeMarkdown(row.Outcome));
            builder.Append(" | ");
            builder.Append(FormatKilobytes(row.MeanAllocatedBytes));
            builder.Append(" | ");
            builder.Append(FormatKilobytes(row.PackageBytes));
            builder.AppendLine(" |");
        }

        builder.AppendLine();
    }

    private static void AppendFullTable(StringBuilder builder, IReadOnlyList<ComparisonSummaryRow> rows) {
        builder.AppendLine("## Full comparison table");
        builder.AppendLine();
        builder.AppendLine("| Row count | Artifact | Scenario | Library | Mean | StdDev | StdErr | Ratio to OfficeIMO | Ratio to best | Alloc | Alloc ratio | Package | Package ratio | Outcome |");
        builder.AppendLine("| ---: | --- | --- | --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | --- |");

        foreach (var row in rows.OrderBy(row => row.RowCount).ThenBy(row => row.ArtifactKind).ThenBy(row => row.Scenario).ThenBy(row => row.MeanMilliseconds)) {
            builder.Append("| ");
            builder.Append(row.RowCount.ToString(CultureInfo.InvariantCulture));
            builder.Append(" | ");
            builder.Append(EscapeMarkdown(row.ArtifactKind));
            builder.Append(" | ");
            builder.Append(EscapeMarkdown(row.Scenario));
            builder.Append(" | ");
            builder.Append(EscapeMarkdown(row.Library));
            builder.Append(" | ");
            builder.Append(FormatMilliseconds(row.MeanMilliseconds));
            builder.Append(" | ");
            builder.Append(FormatMilliseconds(row.StandardDeviationMilliseconds));
            builder.Append(" | ");
            builder.Append(FormatMilliseconds(row.StandardErrorMilliseconds));
            builder.Append(" | ");
            builder.Append(FormatRatio(row.RatioToOfficeImo));
            builder.Append(" | ");
            builder.Append(FormatRatio(row.RatioToBest));
            builder.Append(" | ");
            builder.Append(FormatKilobytes(row.MeanAllocatedBytes));
            builder.Append(" | ");
            builder.Append(FormatRatio(row.AllocatedRatioToOfficeImo));
            builder.Append(" | ");
            builder.Append(FormatKilobytes(row.PackageBytes));
            builder.Append(" | ");
            builder.Append(FormatRatio(row.PackageRatioToOfficeImo));
            builder.Append(" | ");
            builder.Append(EscapeMarkdown(row.Outcome));
            builder.AppendLine(" |");
        }
    }

    private static readonly string[] CsvHeaders = [
        "RowCount",
        "ArtifactKind",
        "Workload",
        "Scenario",
        "Library",
        "MeanMilliseconds",
        "MedianMilliseconds",
        "StandardDeviationMilliseconds",
        "StandardErrorMilliseconds",
        "RatioToOfficeIMO",
        "RatioToBest",
        "BestLibrary",
        "Outcome",
        "MeanAllocatedBytes",
        "MedianAllocatedBytes",
        "AllocatedRatioToOfficeIMO",
        "PackageBytes",
        "PackageRatioToOfficeIMO",
        "WorksheetCompressedBytes",
        "SharedStringsCompressedBytes",
        "StylesCompressedBytes",
        "TablesCompressedBytes",
        "WorksheetRowCount",
        "WorksheetCellCount",
        "SharedStringCount",
        "UniqueSharedStringCount",
        "CellStyleCount",
        "OutputMetric",
        "Notes",
        "ArtifactPath"
    ];

    private static IEnumerable<string> GetCsvValues(ComparisonSummaryRow row) {
        yield return row.RowCount.ToString(CultureInfo.InvariantCulture);
        yield return row.ArtifactKind;
        yield return row.Workload;
        yield return row.Scenario;
        yield return row.Library;
        yield return row.MeanMilliseconds.ToString("F6", CultureInfo.InvariantCulture);
        yield return row.MedianMilliseconds.ToString("F6", CultureInfo.InvariantCulture);
        yield return row.StandardDeviationMilliseconds.ToString("F6", CultureInfo.InvariantCulture);
        yield return row.StandardErrorMilliseconds.ToString("F6", CultureInfo.InvariantCulture);
        yield return FormatNullableNumber(row.RatioToOfficeImo);
        yield return FormatNullableNumber(row.RatioToBest);
        yield return row.BestLibrary;
        yield return row.Outcome;
        yield return FormatNullableNumber(row.MeanAllocatedBytes);
        yield return FormatNullableNumber(row.MedianAllocatedBytes);
        yield return FormatNullableNumber(row.AllocatedRatioToOfficeImo);
        yield return row.PackageBytes?.ToString(CultureInfo.InvariantCulture) ?? string.Empty;
        yield return FormatNullableNumber(row.PackageRatioToOfficeImo);
        yield return row.WorksheetCompressedBytes?.ToString(CultureInfo.InvariantCulture) ?? string.Empty;
        yield return row.SharedStringsCompressedBytes?.ToString(CultureInfo.InvariantCulture) ?? string.Empty;
        yield return row.StylesCompressedBytes?.ToString(CultureInfo.InvariantCulture) ?? string.Empty;
        yield return row.TablesCompressedBytes?.ToString(CultureInfo.InvariantCulture) ?? string.Empty;
        yield return row.WorksheetRowCount?.ToString(CultureInfo.InvariantCulture) ?? string.Empty;
        yield return row.WorksheetCellCount?.ToString(CultureInfo.InvariantCulture) ?? string.Empty;
        yield return row.SharedStringCount?.ToString(CultureInfo.InvariantCulture) ?? string.Empty;
        yield return row.UniqueSharedStringCount?.ToString(CultureInfo.InvariantCulture) ?? string.Empty;
        yield return row.CellStyleCount?.ToString(CultureInfo.InvariantCulture) ?? string.Empty;
        yield return row.OutputMetric?.ToString(CultureInfo.InvariantCulture) ?? string.Empty;
        yield return row.Notes;
        yield return row.ArtifactPath;
    }

    private static string FormatMilliseconds(double value)
        => value <= 0 ? string.Empty : string.Create(CultureInfo.InvariantCulture, $"{value:F2} ms");

    private static string FormatRatio(double? value)
        => value is > 0 ? value.Value.ToString("F2", CultureInfo.InvariantCulture) : string.Empty;

    private static string FormatKilobytes(double? value)
        => value is > 0 ? string.Create(CultureInfo.InvariantCulture, $"{value.Value / 1024.0:F1} KB") : string.Empty;

    private static string FormatKilobytes(long? value)
        => value is > 0 ? string.Create(CultureInfo.InvariantCulture, $"{value.Value / 1024.0:F1} KB") : string.Empty;

    private static string FormatNullableNumber(double? value)
        => value?.ToString("F6", CultureInfo.InvariantCulture) ?? string.Empty;

    private static string EscapeCsv(string value) {
        if (!value.Contains(',') && !value.Contains('"') && !value.Contains('\n') && !value.Contains('\r')) {
            return value;
        }

        return "\"" + value.Replace("\"", "\"\"", StringComparison.Ordinal) + "\"";
    }

    private static string EscapeMarkdown(string? value)
        => (value ?? string.Empty).Replace("|", "\\|", StringComparison.Ordinal);

    private static string GetString(JsonElement element, string propertyName)
        => element.TryGetProperty(propertyName, out JsonElement property) && property.ValueKind == JsonValueKind.String
            ? property.GetString() ?? string.Empty
            : string.Empty;

    private static double GetDouble(JsonElement element, string propertyName)
        => GetNullableDouble(element, propertyName) ?? 0;

    private static double? GetNullableDouble(JsonElement element, string propertyName) {
        if (!element.TryGetProperty(propertyName, out JsonElement property) || property.ValueKind != JsonValueKind.Number) {
            return null;
        }

        return property.GetDouble();
    }

    private static long? GetNullableLong(JsonElement element, string propertyName) {
        if (!element.TryGetProperty(propertyName, out JsonElement property) || property.ValueKind != JsonValueKind.Number) {
            return null;
        }

        return property.GetInt64();
    }

    private static int? GetNullableInt(JsonElement element, string propertyName) {
        if (!element.TryGetProperty(propertyName, out JsonElement property) || property.ValueKind != JsonValueKind.Number) {
            return null;
        }

        return property.GetInt32();
    }

    private static double[] GetDoubleArray(JsonElement element, string propertyName) {
        if (!element.TryGetProperty(propertyName, out JsonElement property) || property.ValueKind != JsonValueKind.Array) {
            return [];
        }

        return property.EnumerateArray()
            .Where(value => value.ValueKind == JsonValueKind.Number)
            .Select(value => value.GetDouble())
            .ToArray();
    }

    private static double CalculateStandardDeviation(IReadOnlyList<double> values) {
        if (values.Count <= 1) {
            return 0;
        }

        double average = values.Average();
        double variance = values.Sum(value => Math.Pow(value - average, 2)) / (values.Count - 1);
        return Math.Sqrt(variance);
    }

    private static double CalculateStandardError(IReadOnlyList<double> values)
        => values.Count == 0 ? 0 : CalculateStandardDeviation(values) / Math.Sqrt(values.Count);

    private sealed class ComparisonSummaryDocument {
        public DateTime GeneratedAtUtc { get; init; }
        public string Notes { get; init; } = string.Empty;
        public List<ComparisonSummaryRow> Rows { get; init; } = [];
    }

    private sealed class ComparisonSummaryRow {
        public string ArtifactKind { get; init; } = string.Empty;
        public int RowCount { get; init; }
        public string Workload { get; init; } = string.Empty;
        public string Scenario { get; init; } = string.Empty;
        public string Library { get; init; } = string.Empty;
        public string Notes { get; init; } = string.Empty;
        public int? OutputMetric { get; init; }
        public double MeanMilliseconds { get; init; }
        public double MedianMilliseconds { get; init; }
        public double StandardDeviationMilliseconds { get; init; }
        public double StandardErrorMilliseconds { get; init; }
        public double? MeanAllocatedBytes { get; init; }
        public double? MedianAllocatedBytes { get; init; }
        public long? PackageBytes { get; set; }
        public long? WorksheetCompressedBytes { get; set; }
        public long? SharedStringsCompressedBytes { get; set; }
        public long? StylesCompressedBytes { get; set; }
        public long? TablesCompressedBytes { get; set; }
        public int? WorksheetRowCount { get; set; }
        public int? WorksheetCellCount { get; set; }
        public int? SharedStringCount { get; set; }
        public int? UniqueSharedStringCount { get; set; }
        public int? CellStyleCount { get; set; }
        public string ArtifactPath { get; init; } = string.Empty;
        public string BestLibrary { get; set; } = string.Empty;
        public double? BestMeanMilliseconds { get; set; }
        public double? RatioToOfficeImo { get; set; }
        public double? RatioToBest { get; set; }
        public double? AllocatedRatioToOfficeImo { get; set; }
        public double? PackageRatioToOfficeImo { get; set; }
        public string Outcome { get; set; } = string.Empty;
    }
}
