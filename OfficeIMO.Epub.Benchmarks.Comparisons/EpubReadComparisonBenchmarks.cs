using BenchmarkDotNet.Attributes;

namespace OfficeIMO.Epub.Benchmarks.Comparisons;

/// <summary>Compares complete, output-validated EPUB open-and-enumerate workflows.</summary>
[MemoryDiagnoser]
public class EpubReadComparisonBenchmarks {
    private byte[] _package = Array.Empty<byte>();

    /// <summary>Gets or sets the deterministic corpus scale.</summary>
    [Params("Small", "Normal")]
    public string Scale { get; set; } = string.Empty;

    /// <summary>Creates the shared source package and validates both readers.</summary>
    [GlobalSetup]
    public void Setup() {
        EpubComparisonScale scale = EpubComparisonCorpus.Get(Scale);
        _package = EpubComparisonCorpus.CreatePackage(scale);
        EpubComparisonValidation.Validate(Scale);
        long officeIMO = EpubComparisonWorkflows.ReadOfficeIMO(_package);
        long versOne = EpubComparisonWorkflows.ReadVersOne(_package);
        if (officeIMO != versOne) {
            throw new InvalidOperationException(
                $"EPUB read observation differs for {Scale}: OfficeIMO {officeIMO}; VersOne.Epub {versOne}.");
        }
    }

    /// <summary>Opens the package and enumerates its reading order through OfficeIMO.</summary>
    [Benchmark(Baseline = true, Description = "OfficeIMO")]
    public long OfficeIMO() => EpubComparisonWorkflows.ReadOfficeIMO(_package);

    /// <summary>Opens the same package and enumerates its reading order through VersOne.Epub.</summary>
    [Benchmark(Description = "VersOne.Epub+HAP")]
    public long VersOneEpub() => EpubComparisonWorkflows.ReadVersOne(_package);
}
