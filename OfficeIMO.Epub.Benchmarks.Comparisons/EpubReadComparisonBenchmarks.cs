using BenchmarkDotNet.Attributes;
using OfficeIMO.Benchmarks;

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
        ValidateLane("OfficeIMO", () => EpubComparisonValidation.ValidateLane(scale, EpubComparisonWorkflows.RetainOfficeIMO(_package)));
        ValidateLane("VersOne.Epub", () => EpubComparisonValidation.ValidateLane(scale, EpubComparisonWorkflows.RetainVersOne(_package)));
        try {
            EpubComparisonValidation.Validate(Scale);
            long officeIMO = EpubComparisonWorkflows.ReadOfficeIMO(_package);
            long versOne = EpubComparisonWorkflows.ReadVersOne(_package);
            if (officeIMO != versOne) {
                throw new InvalidOperationException(
                    $"EPUB read observation differs for {Scale}: OfficeIMO {officeIMO}; VersOne.Epub {versOne}.");
            }
        } catch (InvalidOperationException exception) {
            BenchmarkOutputValidationLog.Failed("OfficeIMO", $"Scale={Scale}; {exception.Message}");
            BenchmarkOutputValidationLog.Failed("VersOne.Epub", $"Scale={Scale}; {exception.Message}");
            throw;
        }
    }

    private void ValidateLane(string library, Func<long> validate) {
        long characters;
        try {
            characters = validate();
        } catch (InvalidOperationException exception) {
            BenchmarkOutputValidationLog.Failed(library, $"Scale={Scale}; {exception.Message}");
            throw;
        }

        EpubComparisonScale scale = EpubComparisonCorpus.Get(Scale);
        BenchmarkOutputValidationLog.Passed(
            library,
            $"Scale={Scale}; title, creator, language, {scale.Chapters} chapters in reading order with expected paths and text ({characters} characters)");
    }

    /// <summary>Opens the package and enumerates its reading order through OfficeIMO.</summary>
    [Benchmark(Baseline = true, Description = "OfficeIMO")]
    [BenchmarkCategory("library:OfficeIMO", "case:Default")]
    public long OfficeIMO() => EpubComparisonWorkflows.ReadOfficeIMO(_package);

    /// <summary>Opens the same package and enumerates its reading order through VersOne.Epub.</summary>
    [Benchmark(Description = "VersOne.Epub+HAP")]
    [BenchmarkCategory("library:VersOne.Epub", "case:Default")]
    public long VersOneEpub() => EpubComparisonWorkflows.ReadVersOne(_package);
}
