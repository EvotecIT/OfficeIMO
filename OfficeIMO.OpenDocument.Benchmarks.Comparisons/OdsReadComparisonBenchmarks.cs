using BenchmarkDotNet.Attributes;

namespace OfficeIMO.OpenDocument.Benchmarks.Comparisons;

/// <summary>Compares complete ODS open-and-enumerate workflows over the same package.</summary>
[MemoryDiagnoser]
public class OdsReadComparisonBenchmarks {
    private byte[] _package = Array.Empty<byte>();

    /// <summary>Gets or sets the deterministic corpus scale.</summary>
    [Params("Small", "Normal")]
    public string Scale { get; set; } = string.Empty;

    /// <summary>Creates the shared source package and validates both readers.</summary>
    [GlobalSetup]
    public void Setup() {
        OdsComparisonScale scale = OdsComparisonCorpus.Get(Scale);
        _package = OdsComparisonWorkflows.CreateOfficeIMO(scale);
        long expected = checked((long)scale.Rows * scale.Columns * OdsComparisonCorpus.Cell(0, 0).Length);
        long officeIMO = OdsComparisonWorkflows.ReadOfficeIMO(_package);
        long openStandardLibrary = OdsComparisonWorkflows.ReadOpenStandardLibrary(_package).GetAwaiter().GetResult();
        if (officeIMO != expected || openStandardLibrary != expected) {
            throw new InvalidOperationException(
                $"ODS read validation failed for {Scale}: expected {expected}; " +
                $"OfficeIMO {officeIMO}; OpenStandardLibrary {openStandardLibrary}.");
        }
    }

    /// <summary>Opens and enumerates every populated cell through OfficeIMO.</summary>
    [Benchmark(Baseline = true, Description = "OfficeIMO")]
    public long OfficeIMO() => OdsComparisonWorkflows.ReadOfficeIMO(_package);

    /// <summary>Opens and enumerates every populated cell through OpenStandardLibrary.</summary>
    [Benchmark(Description = "OpenStandardLibrary")]
    public Task<long> OpenStandardLibrary() => OdsComparisonWorkflows.ReadOpenStandardLibrary(_package);
}
