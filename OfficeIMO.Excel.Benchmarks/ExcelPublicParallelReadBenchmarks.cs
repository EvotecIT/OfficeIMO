#nullable enable

using System.Data.Common;
using BenchmarkDotNet.Attributes;
using OfficeIMO.Benchmarks;
using OfficeIMO.Data;

namespace OfficeIMO.Excel.Benchmarks;

/// <summary>
/// Measures the exact public typed-row contract over the hash-pinned 65K workbook
/// in each native reader family. Both methods open, parse, map, and validate the
/// same fourteen columns; only the ordered mapping strategy differs.
/// </summary>
[MemoryDiagnoser]
[RankColumn]
public class ExcelPublicParallelReadBenchmarks {
    private readonly ParallelRowMappingOptions _parallelOptions = new() {
        MaxDegreeOfParallelism = 16,
        BatchSize = 1024
    };
    private string _path = string.Empty;
    private ExcelReadObservation _expected;

    [Params(ExcelPublicReadFormat.Xlsx, ExcelPublicReadFormat.Xlsb, ExcelPublicReadFormat.Xls)]
    public ExcelPublicReadFormat Format { get; set; }

    [GlobalSetup]
    public void Setup() {
        string fileName = Format switch {
            ExcelPublicReadFormat.Xlsx => MarkPflug65KFixture.XlsxFileName,
            ExcelPublicReadFormat.Xlsb => MarkPflug65KFixture.XlsbFileName,
            ExcelPublicReadFormat.Xls => MarkPflug65KFixture.XlsFileName,
            _ => throw new ArgumentOutOfRangeException(nameof(Format))
        };
        MarkPflug65KFixture.EnsureAuthentic(fileName);
        _path = Path.Combine(MarkPflug65KFixture.Root, fileName);
        _expected = MarkPflug65KXlsxBenchmarks.ExpectedObservation();
        Validate(nameof(Sequential), Sequential());
        Validate(nameof(OrderedParallel), OrderedParallel());
    }

    [Benchmark(Baseline = true)]
    public ExcelReadObservation Sequential() {
        using DbDataReader reader = OpenReader();
        return Observe(reader.RowsAs<MarkPflugSalesRow>());
    }

    [Benchmark]
    public ExcelReadObservation OrderedParallel() {
        using DbDataReader reader = OpenReader();
        return Observe(reader.RowsAsParallel<MarkPflugSalesRow>(_parallelOptions));
    }

    private DbDataReader OpenReader() => ExcelDocument.OpenDataReader(
        _path,
        new ExcelReadOptions { NumericAsDecimal = true });

    private static ExcelReadObservation Observe(IEnumerable<MarkPflugSalesRow> rows) {
        var observation = new ExcelObservationAccumulator();
        foreach (MarkPflugSalesRow row in rows) {
            observation.BeginRow();
            observation.Add(row.Region);
            observation.Add(row.Country);
            observation.Add(row.ItemType);
            observation.Add(row.SalesChannel);
            observation.Add(row.OrderPriority);
            observation.Add(row.OrderDate);
            observation.Add(row.OrderId);
            observation.Add(row.ShipDate);
            observation.Add(row.UnitsSold);
            observation.Add(row.UnitPrice);
            observation.Add(row.UnitCost);
            observation.Add(row.TotalRevenue);
            observation.Add(row.TotalCost);
            observation.Add(row.TotalProfit);
        }

        return observation.Build();
    }

    private void Validate(string method, ExcelReadObservation actual) {
        if (actual != _expected) {
            throw new InvalidDataException(
                $"{method} did not perform the complete {Format} typed-row workload. Expected {_expected}; actual {actual}.");
        }
    }

    public sealed class MarkPflugSalesRow {
        public string Region { get; set; } = string.Empty;
        public string Country { get; set; } = string.Empty;

        [ExcelColumn("Item Type")]
        public string ItemType { get; set; } = string.Empty;

        [ExcelColumn("Sales Channel")]
        public string SalesChannel { get; set; } = string.Empty;

        [ExcelColumn("Order Priority")]
        public string OrderPriority { get; set; } = string.Empty;

        [ExcelColumn("Order Date")]
        public DateTime OrderDate { get; set; }

        [ExcelColumn("Order ID")]
        public int OrderId { get; set; }

        [ExcelColumn("Ship Date")]
        public DateTime ShipDate { get; set; }

        [ExcelColumn("Units Sold")]
        public int UnitsSold { get; set; }

        [ExcelColumn("Unit Price")]
        public decimal UnitPrice { get; set; }

        [ExcelColumn("Unit Cost")]
        public decimal UnitCost { get; set; }

        [ExcelColumn("Total Revenue")]
        public decimal TotalRevenue { get; set; }

        [ExcelColumn("Total Cost")]
        public decimal TotalCost { get; set; }

        [ExcelColumn("Total Profit")]
        public decimal TotalProfit { get; set; }
    }
}

public enum ExcelPublicReadFormat {
    Xlsx,
    Xlsb,
    Xls
}
