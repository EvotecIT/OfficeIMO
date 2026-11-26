using BenchmarkDotNet.Attributes;
using OfficeIMO.Benchmarks;

namespace OfficeIMO.OpenDocument.Benchmarks.Comparisons;

/// <summary>Compares complete, output-validated ODS creation in memory.</summary>
[MemoryDiagnoser]
public class OdsCreateComparisonBenchmarks {
    private OdsComparisonScale _scale = null!;

    /// <summary>Gets or sets the deterministic corpus scale.</summary>
    [Params("Small", "Normal")]
    public string Scale { get; set; } = string.Empty;

    /// <summary>Prepares and validates both generated packages before measurement begins.</summary>
    [GlobalSetup]
    public void Setup() {
        _scale = OdsComparisonCorpus.Get(Scale);
        OdsComparisonReport report = OdsComparisonValidation.Validate(Scale);
        foreach (OdsOutputEvidence output in new[] { report.OfficeIMO, report.OpenStandardLibrary }) {
            BenchmarkOutputValidationLog.Passed(output.Implementation,
                $"Scale={Scale}; OfficeIMO reread all {output.RecordCount} cell values by position; characters={output.ContentLength}");
        }
    }

    /// <summary>Creates the complete ODS package through OfficeIMO.</summary>
    [Benchmark(Baseline = true, Description = "OfficeIMO")]
    [BenchmarkCategory("library:OfficeIMO", "case:Default")]
    public byte[] OfficeIMO() => OdsComparisonWorkflows.CreateOfficeIMO(_scale);

    /// <summary>Creates the same complete ODS package through OpenStandardLibrary.</summary>
    [Benchmark(Description = "OpenStandardLibrary")]
    [BenchmarkCategory("library:OpenStandardLibrary", "case:Default")]
    public Task<byte[]> OpenStandardLibrary() => OdsComparisonWorkflows.CreateOpenStandardLibrary(_scale);
}
