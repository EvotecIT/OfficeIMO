using BenchmarkDotNet.Attributes;
using MimeKit;
using OfficeIMO.Benchmarks;

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

        EmailMimeObservation expected = EmailMimeComparisonCorpus.CreateExpectedObservation(scale);
        EmailMimeComparisonValidation.RecordLane("OfficeIMO", Scale, expected, ("OfficeIMO", EmailMimeComparisonValidation.ObserveOffice(_input)));
        EmailMimeComparisonValidation.RecordLane("MimeKit", Scale, expected, ("MimeKit", EmailMimeComparisonValidation.ObserveMimeKit(_input)));
#if NET10_0_OR_GREATER
        EmailMimeComparisonValidation.RecordLane("OneMail", Scale, expected, ("OneMail", EmailMimeComparisonValidation.ObserveOneMail(_input)));
#endif
    }

    /// <summary>Reads the complete EML through OfficeIMO and consumes all decoded payloads.</summary>
    [Benchmark(Baseline = true, Description = "OfficeIMO")]
    [BenchmarkCategory("library:OfficeIMO", "case:Default")]
    public int OfficeIMO() => EmailMimeComparisonValidation.ConsumeOffice(_input);

    /// <summary>Reads the same EML through MimeKit and consumes all decoded payloads.</summary>
    [Benchmark(Description = "MimeKit")]
    [BenchmarkCategory("library:MimeKit", "case:Default")]
    public int MimeKit() => EmailMimeComparisonValidation.ConsumeMimeKit(_input);
#if NET10_0_OR_GREATER

    /// <summary>Reads the same EML through OneMail's message view and consumes all decoded payloads.</summary>
    [Benchmark(Description = "OneMail")]
    [BenchmarkCategory("library:OneMail", "case:Default")]
    public int OneMail() => EmailMimeComparisonValidation.ConsumeOneMail(_input);
#endif
}

/// <summary>Compares complete in-memory EML serialization from equivalent prepared models.</summary>
[MemoryDiagnoser]
public class EmailMimeWriteBenchmarks {
    private EmailDocument _officeDocument = null!;
    private MimeMessage _mimeMessage = null!;
#if NET10_0_OR_GREATER
    private RecursiveAI.OneMail.Composition.OutgoingMessage _oneMailMessage = null!;
#endif

    /// <summary>Gets or sets the deterministic MIME corpus scale.</summary>
    [ParamsSource(nameof(Scales))]
    public string Scale { get; set; } = string.Empty;

    /// <summary>Returns the corpus scales shared by both implementations.</summary>
    public IEnumerable<string> Scales() => EmailMimeComparisonCorpus.ScaleNames;

    /// <summary>Creates equivalent prepared models and validates every lane's output before measurement.</summary>
    [GlobalSetup]
    public void Setup() {
        EmailMimeBenchmarkScale scale = EmailMimeComparisonCorpus.Get(Scale);
        _officeDocument = EmailMimeComparisonCorpus.CreateOfficeDocument(scale);
        _mimeMessage = EmailMimeComparisonCorpus.CreateMimeMessage(scale);
#if NET10_0_OR_GREATER
        _oneMailMessage = EmailMimeComparisonCorpus.CreateOneMailMessage(scale);
#endif
        EmailMimeComparisonValidation.Validate(Scale);

        // Each lane's own output, read back by both parsers.
        EmailMimeObservation expected = EmailMimeComparisonCorpus.CreateExpectedObservation(scale);
        byte[] officeOutput = OfficeIMO();
        byte[] mimeKitOutput = MimeKit();
        EmailMimeComparisonValidation.RecordLane(
            "OfficeIMO",
            Scale,
            expected,
            ("OfficeIMO", EmailMimeComparisonValidation.ObserveOffice(officeOutput)),
            ("MimeKit", EmailMimeComparisonValidation.ObserveMimeKit(officeOutput)));
        EmailMimeComparisonValidation.RecordLane(
            "MimeKit",
            Scale,
            expected,
            ("MimeKit", EmailMimeComparisonValidation.ObserveMimeKit(mimeKitOutput)),
            ("OfficeIMO", EmailMimeComparisonValidation.ObserveOffice(mimeKitOutput)));
#if NET10_0_OR_GREATER
        byte[] oneMailOutput = OneMail();
        EmailMimeComparisonValidation.RecordLane(
            "OneMail",
            Scale,
            expected,
            ("MimeKit", EmailMimeComparisonValidation.ObserveMimeKit(oneMailOutput)),
            ("OfficeIMO", EmailMimeComparisonValidation.ObserveOffice(oneMailOutput)));
#endif
    }

    /// <summary>Serializes the prepared model through OfficeIMO.</summary>
    [Benchmark(Baseline = true, Description = "OfficeIMO")]
    [BenchmarkCategory("library:OfficeIMO", "case:Default")]
    public byte[] OfficeIMO() => _officeDocument.ToBytes(EmailFileFormat.Eml);

    /// <summary>Serializes the equivalent prepared model through MimeKit.</summary>
    [Benchmark(Description = "MimeKit")]
    [BenchmarkCategory("library:MimeKit", "case:Default")]
    public byte[] MimeKit() => EmailMimeComparisonCorpus.WriteMimeKit(_mimeMessage);
#if NET10_0_OR_GREATER

    /// <summary>Builds the MIME tree from the equivalent prepared OneMail message and serializes it.</summary>
    [Benchmark(Description = "OneMail")]
    [BenchmarkCategory("library:OneMail", "case:Default")]
    public byte[] OneMail() => EmailMimeComparisonCorpus.WriteOneMail(_oneMailMessage);
#endif

    /// <summary>Releases the prepared MimeKit model.</summary>
    [GlobalCleanup]
    public void Cleanup() => _mimeMessage.Dispose();
}
