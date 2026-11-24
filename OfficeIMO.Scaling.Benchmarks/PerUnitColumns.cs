using BenchmarkDotNet.Columns;
using BenchmarkDotNet.Configs;
using BenchmarkDotNet.Reports;
using BenchmarkDotNet.Running;

namespace OfficeIMO.Scaling.Benchmarks;

internal sealed class PerUnitConfig : ManualConfig {
    public PerUnitConfig() {
        AddColumn(new PerUnitColumn(allocation: false), new PerUnitColumn(allocation: true));
    }
}

/// <summary>Mean time or allocation divided by the <c>Units</c> parameter.</summary>
internal sealed class PerUnitColumn : IColumn {
    private readonly bool _allocation;

    public PerUnitColumn(bool allocation) {
        _allocation = allocation;
    }

    public string Id => _allocation ? "AllocPerUnit" : "TimePerUnit";
    public string ColumnName => _allocation ? "B/unit" : "µs/unit";
    public bool AlwaysShow => true;
    public ColumnCategory Category => ColumnCategory.Custom;
    public int PriorityInCategory => _allocation ? 1 : 0;
    public bool IsNumeric => true;
    public UnitType UnitType => UnitType.Dimensionless;
    public string Legend => _allocation ? "Allocated bytes per unit" : "Mean microseconds per unit";

    public string GetValue(Summary summary, BenchmarkCase benchmarkCase) => GetValue(summary, benchmarkCase, SummaryStyle.Default);

    public string GetValue(Summary summary, BenchmarkCase benchmarkCase, SummaryStyle style) {
        BenchmarkReport? report = summary[benchmarkCase];
        object? units = benchmarkCase.Parameters["Units"];
        if (report?.ResultStatistics == null || units is not int count || count <= 0) {
            return "-";
        }

        if (_allocation) {
            long? bytes = report.GcStats.GetBytesAllocatedPerOperation(benchmarkCase);
            return bytes == null ? "-" : (bytes.Value / (double)count).ToString("N0", System.Globalization.CultureInfo.InvariantCulture);
        }

        return (report.ResultStatistics.Mean / 1000d / count).ToString("0.00", System.Globalization.CultureInfo.InvariantCulture);
    }

    public bool IsAvailable(Summary summary) => true;
    public bool IsDefault(Summary summary, BenchmarkCase benchmarkCase) => false;
}
