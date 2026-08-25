using BenchmarkDotNet.Attributes;
using MimeKit;

namespace OfficeIMO.Email.Benchmarks.Comparisons;

/// <summary>Compares complete EML parsing with decoded body and attachment materialization.</summary>
[MemoryDiagnoser]
public class EmailMimeReadBenchmarks {
    private byte[] _input = Array.Empty<byte>();

    /// <summary>Gets or sets the deterministic MIME corpus scale.</summary>
    [ParamsSource(nameof(Scales))]
    public string Scale { get; set; } = string.Empty;

    /// <summary>Returns the corpus scales shared by both implementations.</summary>
    public IEnumerable<string> Scales() => EmailMimeComparisonCorpus.ScaleNames;

    /// <summary>Creates and validates the shared input before measurement.</summary>
    [GlobalSetup]
    public void Setup() {
        EmailMimeBenchmarkScale scale = EmailMimeComparisonCorpus.Get(Scale);
        using MimeMessage message = EmailMimeComparisonCorpus.CreateMimeMessage(scale);
        _input = EmailMimeComparisonCorpus.WriteMimeKit(message);
        EmailMimeComparisonValidation.Validate(Scale);
    }

    /// <summary>Reads the complete EML through OfficeIMO and consumes all decoded payloads.</summary>
    [Benchmark(Baseline = true, Description = "OfficeIMO")]
    public int OfficeIMO() => EmailMimeComparisonValidation.ConsumeOffice(_input);

    /// <summary>Reads the same EML through MimeKit and consumes all decoded payloads.</summary>
    [Benchmark(Description = "MimeKit")]
    public int MimeKit() => EmailMimeComparisonValidation.ConsumeMimeKit(_input);
}

/// <summary>Compares complete in-memory EML serialization from equivalent prepared models.</summary>
[MemoryDiagnoser]
public class EmailMimeWriteBenchmarks {
    private EmailDocument _officeDocument = null!;
    private MimeMessage _mimeMessage = null!;

    /// <summary>Gets or sets the deterministic MIME corpus scale.</summary>
    [ParamsSource(nameof(Scales))]
    public string Scale { get; set; } = string.Empty;

    /// <summary>Returns the corpus scales shared by both implementations.</summary>
    public IEnumerable<string> Scales() => EmailMimeComparisonCorpus.ScaleNames;

    /// <summary>Creates equivalent prepared models and validates both outputs before measurement.</summary>
    [GlobalSetup]
    public void Setup() {
        EmailMimeBenchmarkScale scale = EmailMimeComparisonCorpus.Get(Scale);
        _officeDocument = EmailMimeComparisonCorpus.CreateOfficeDocument(scale);
        _mimeMessage = EmailMimeComparisonCorpus.CreateMimeMessage(scale);
        EmailMimeComparisonValidation.Validate(Scale);
    }

    /// <summary>Serializes the prepared model through OfficeIMO.</summary>
    [Benchmark(Baseline = true, Description = "OfficeIMO")]
    public byte[] OfficeIMO() => _officeDocument.ToBytes(EmailFileFormat.Eml);

    /// <summary>Serializes the equivalent prepared model through MimeKit.</summary>
    [Benchmark(Description = "MimeKit")]
    public byte[] MimeKit() => EmailMimeComparisonCorpus.WriteMimeKit(_mimeMessage);

    /// <summary>Releases the prepared MimeKit model.</summary>
    [GlobalCleanup]
    public void Cleanup() => _mimeMessage.Dispose();
}
