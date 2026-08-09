#nullable enable

using System.Data.Common;
using System.Data;
using System.Numerics;
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
        new ExcelReadOptions {
            NumericAsDecimal = true,
            InferSchema = true
        });

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

/// <summary>
/// Measures the public ordered-parallel crossover when row projection performs enough
/// independent CPU work to repay snapshot and scheduling overhead. Both methods open the
/// same native format reader, read the same fields, execute the same deterministic projection,
/// preserve source order, and validate the complete result.
/// </summary>
[MemoryDiagnoser]
[RankColumn]
public class ExcelPublicParallelProjectionBenchmarks {
    private const int ExpectedProjectionRounds = 4096;
    private const ulong ExpectedProjectionChecksum = 11619009235369511762UL;
    private readonly ParallelRowMappingOptions _parallelOptions = new() {
        MaxDegreeOfParallelism = 16,
        BatchSize = 1024
    };
    private string _path = string.Empty;
    private ProjectionObservation _expected;

    [Params(ExcelPublicReadFormat.Xlsx, ExcelPublicReadFormat.Xlsb, ExcelPublicReadFormat.Xls)]
    public ExcelPublicReadFormat Format { get; set; }

    [Params(ExpectedProjectionRounds)]
    public int ProjectionRounds { get; set; }

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
        ValidateConcurrentProjection();
        _expected = Sequential();
        Validate(nameof(Sequential), _expected);
        Validate(nameof(OrderedParallel), OrderedParallel());
    }

    [Benchmark(Baseline = true)]
    public ProjectionObservation Sequential() {
        using DbDataReader reader = OpenReader();
        return Observe(reader.RowsAs(Project));
    }

    [Benchmark]
    public ProjectionObservation OrderedParallel() {
        using DbDataReader reader = OpenReader();
        return Observe(reader.RowsAsParallel(Project, _parallelOptions));
    }

    private DbDataReader OpenReader() => ExcelDocument.OpenDataReader(
        _path,
        new ExcelReadOptions {
            NumericAsDecimal = true,
            InferSchema = true
        });

    private void ValidateConcurrentProjection() {
        using DbDataReader reader = OpenReader();
        using var firstWorkers = new Barrier(2);
        int calls = 0;
        foreach (int _ in reader.RowsAsParallel(
                     row => {
                         if (Interlocked.Increment(ref calls) <= 2 &&
                             !firstWorkers.SignalAndWait(TimeSpan.FromSeconds(30))) {
                             throw new InvalidDataException(
                                 $"The {Format} ordered-parallel benchmark did not start two projection workers.");
                         }
                         return row.GetInt32(6);
                     },
                     _parallelOptions)) {
        }
    }

    private ProjectedRow Project(IDataRecord row) {
        string region = row.GetString(0);
        string country = row.GetString(1);
        int orderId = row.GetInt32(6);
        decimal totalProfit = row.GetDecimal(13);

        ulong hash = unchecked((ulong)(uint)orderId) ^ unchecked((ulong)decimal.GetBits(totalProfit)[0]);
        for (int index = 0; index < region.Length; index++) {
            hash = Mix(hash, region[index]);
        }
        for (int index = 0; index < country.Length; index++) {
            hash = Mix(hash, country[index]);
        }
        for (int round = 0; round < ProjectionRounds; round++) {
            hash = Mix(hash, unchecked((uint)(orderId + round)));
        }

        return new ProjectedRow(orderId, hash);
    }

    private static ulong Mix(ulong value, ulong input) =>
        unchecked(BitOperations.RotateLeft(value ^ input, 17) * 0x9E3779B185EBCA87UL);

    private static ProjectionObservation Observe(IEnumerable<ProjectedRow> rows) {
        int count = 0;
        ulong checksum = 1469598103934665603UL;
        foreach (ProjectedRow row in rows) {
            count++;
            checksum = Mix(checksum, unchecked((uint)row.OrderId));
            checksum = Mix(checksum, row.Hash);
        }
        return new ProjectionObservation(count, checksum);
    }

    private void Validate(string method, ProjectionObservation actual) {
        if (actual.RowCount != MarkPflug65KFixture.ExpectedRows
            || actual.Checksum != ExpectedProjectionChecksum
            || actual != _expected) {
            throw new InvalidDataException(
                $"{method} did not perform the complete {Format} projection workload. " +
                $"Expected {MarkPflug65KFixture.ExpectedRows} rows and checksum {ExpectedProjectionChecksum}; actual {actual}.");
        }
    }

    private readonly record struct ProjectedRow(int OrderId, ulong Hash);

    public readonly record struct ProjectionObservation(int RowCount, ulong Checksum);
}

public enum ExcelPublicReadFormat {
    Xlsx,
    Xlsb,
    Xls
}
