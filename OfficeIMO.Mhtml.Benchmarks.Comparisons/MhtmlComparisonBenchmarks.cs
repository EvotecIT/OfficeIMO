using BenchmarkDotNet.Attributes;
using MimeKit;

namespace OfficeIMO.Mhtml.Benchmarks.Comparisons;

/// <summary>Compares complete MHTML parsing, HTML DOM creation, and decoded-resource materialization.</summary>
[MemoryDiagnoser]
public class MhtmlReadBenchmarks {
    private byte[] _input = Array.Empty<byte>();

    /// <summary>Gets or sets the deterministic archive scale.</summary>
    [ParamsSource(nameof(Scales))]
    public string Scale { get; set; } = string.Empty;

    /// <summary>Returns the scales shared by both implementations.</summary>
    public IEnumerable<string> Scales() => MhtmlComparisonCorpus.ScaleNames;

    /// <summary>Creates and validates the shared MimeKit-produced archive.</summary>
    [GlobalSetup]
    public void Setup() {
        MhtmlBenchmarkScale scale = MhtmlComparisonCorpus.Get(Scale);
        using MimeMessage message = MhtmlComparisonCorpus.CreateMimeMessage(scale);
        _input = MhtmlComparisonCorpus.WriteMimeKit(message);
        MhtmlComparisonValidation.Validate(Scale);
    }

    /// <summary>Loads the archive through OfficeIMO.</summary>
    [Benchmark(Baseline = true, Description = "OfficeIMO")]
    public MhtmlDocument OfficeIMO() => MhtmlComparisonValidation.LoadOffice(_input);

    /// <summary>Loads MIME through MimeKit, parses HTML through AngleSharp, and retains decoded resources.</summary>
    [Benchmark(Description = "MimeKit + AngleSharp")]
    public object MimeKit() => MhtmlComparisonValidation.LoadMimeKit(_input);
}

/// <summary>Compares complete in-memory MHTML serialization from equivalent prepared models.</summary>
[MemoryDiagnoser]
public class MhtmlWriteBenchmarks {
    private MhtmlDocument _officeDocument = null!;
    private MimeMessage _mimeMessage = null!;

    /// <summary>Gets or sets the deterministic archive scale.</summary>
    [ParamsSource(nameof(Scales))]
    public string Scale { get; set; } = string.Empty;

    /// <summary>Returns the scales shared by both implementations.</summary>
    public IEnumerable<string> Scales() => MhtmlComparisonCorpus.ScaleNames;

    /// <summary>Creates equivalent prepared models and validates both outputs.</summary>
    [GlobalSetup]
    public void Setup() {
        MhtmlBenchmarkScale scale = MhtmlComparisonCorpus.Get(Scale);
        _officeDocument = MhtmlComparisonCorpus.CreateOfficeDocument(scale);
        _mimeMessage = MhtmlComparisonCorpus.CreateMimeMessage(scale);
        MhtmlComparisonValidation.Validate(Scale);
    }

    /// <summary>Serializes the prepared OfficeIMO archive.</summary>
    [Benchmark(Baseline = true, Description = "OfficeIMO")]
    public byte[] OfficeIMO() => _officeDocument.ToBytes();

    /// <summary>Serializes the equivalent prepared MimeKit archive.</summary>
    [Benchmark(Description = "MimeKit")]
    public byte[] MimeKit() => MhtmlComparisonCorpus.WriteMimeKit(_mimeMessage);

    /// <summary>Releases the prepared MimeKit model.</summary>
    [GlobalCleanup]
    public void Cleanup() => _mimeMessage.Dispose();
}
