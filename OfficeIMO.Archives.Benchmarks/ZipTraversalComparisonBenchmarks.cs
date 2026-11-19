using BenchmarkDotNet.Attributes;
using OfficeIMO.Benchmarks;

namespace OfficeIMO.Archives.Benchmarks;

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
        ZipComparisonReport report;
        try {
            report = ZipComparisonValidation.Validate(scale, _input);
        } catch (InvalidOperationException exception) {
            BenchmarkOutputValidationLog.Failed("OfficeIMO", $"Scale={Scale}; {exception.Message}");
            BenchmarkOutputValidationLog.Failed("System.IO.Compression", $"Scale={Scale}; {exception.Message}");
            throw;
        }

        // Both projections carry the expected file count and uncompressed bytes and the same names, depths, sizes and
        // timestamps (one structural fingerprint).
        string detail = $"Scale={Scale}; {report.EntryCount} files, {report.TotalUncompressedBytes} uncompressed bytes, fingerprint {report.StructuralFingerprint}";
        BenchmarkOutputValidationLog.Passed("OfficeIMO", detail);
        BenchmarkOutputValidationLog.Passed("System.IO.Compression", detail);
    }

    /// <summary>Traverses metadata with OfficeIMO path and expansion-safety policy.</summary>
    [Benchmark(Baseline = true, Description = "OfficeIMO")]
    [BenchmarkCategory("library:OfficeIMO", "case:Default")]
    public ArchiveTraversalResult OfficeIMO() => ZipComparisonWorkflows.TraverseOffice(_input);

    /// <summary>Traverses the same safe archive through raw platform metadata projection.</summary>
    [Benchmark(Description = "System.IO.Compression")]
    [BenchmarkCategory("library:System.IO.Compression", "case:Default")]
    public IReadOnlyList<ZipProjectionDescriptor> Platform() => ZipComparisonWorkflows.TraversePlatform(_input);
#if NET10_0_OR_GREATER

    /// <summary>Lists the same archive through OneArchive with its default reading limits.</summary>
    [Benchmark(Description = "OneArchive")]
    [BenchmarkCategory("library:onearchive")]
    public IReadOnlyList<ZipProjectionDescriptor> OneArchive() => ZipComparisonWorkflows.TraverseOneArchive(_input);
#endif
}
