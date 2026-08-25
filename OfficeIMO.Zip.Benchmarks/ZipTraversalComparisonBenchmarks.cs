using BenchmarkDotNet.Attributes;

namespace OfficeIMO.Zip.Benchmarks;

/// <summary>Measures guarded OfficeIMO ZIP traversal against an equivalent raw metadata projection.</summary>
[MemoryDiagnoser]
public class ZipTraversalComparisonBenchmarks {
    private byte[] _input = Array.Empty<byte>();

    /// <summary>Gets or sets the deterministic archive scale.</summary>
    [ParamsSource(nameof(Scales))]
    public string Scale { get; set; } = string.Empty;

    /// <summary>Returns the corpus scales shared by both implementations.</summary>
    public IEnumerable<string> Scales() => ZipComparisonCorpus.ScaleNames;

    /// <summary>Creates and validates the shared archive before measurement.</summary>
    [GlobalSetup]
    public void Setup() {
        ZipBenchmarkScale scale = ZipComparisonCorpus.Get(Scale);
        _input = ZipComparisonCorpus.CreateArchive(scale);
        ZipComparisonValidation.Validate(scale, _input);
    }

    /// <summary>Traverses metadata with OfficeIMO path and expansion-safety policy.</summary>
    [Benchmark(Baseline = true, Description = "OfficeIMO")]
    public ZipTraversalResult OfficeIMO() => ZipComparisonWorkflows.TraverseOffice(_input);

    /// <summary>Traverses the same safe archive through raw platform metadata projection.</summary>
    [Benchmark(Description = "System.IO.Compression")]
    public IReadOnlyList<ZipProjectionDescriptor> Platform() => ZipComparisonWorkflows.TraversePlatform(_input);
}
