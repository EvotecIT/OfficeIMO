using System.Text;
using BenchmarkDotNet.Attributes;
using OfficeIMO.Benchmarks;
using OfficeIMO.Html;

namespace OfficeIMO.Rtf.Benchmarks.Comparisons;

/// <summary>Compares complete, output-validated RTF-to-HTML workflows.</summary>
[MemoryDiagnoser]
public class RtfToHtmlComparisonBenchmarks {
    private string _rtf = string.Empty;

    /// <summary>Gets or sets the deterministic corpus scale.</summary>
    [ParamsSource(nameof(CorpusScales))]
    public string Scale { get; set; } = string.Empty;

    /// <summary>Returns the corpus scales shared by both implementations.</summary>
    public IEnumerable<string> CorpusScales() => RtfHtmlComparisonCorpus.Scales;

    /// <summary>Prepares and validates both outputs before measurement begins.</summary>
    [GlobalSetup]
    public void Setup() {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        _rtf = RtfHtmlComparisonCorpus.Get(Scale).Rtf;
        RtfHtmlComparisonReport report;
        try {
            report = RtfHtmlComparisonValidation.Validate(Scale, _rtf);
        } catch (InvalidOperationException exception) {
            BenchmarkOutputValidationLog.Failed("OfficeIMO", $"Scale={Scale}; {exception.Message}");
            BenchmarkOutputValidationLog.Failed("RtfPipe", $"Scale={Scale}; {exception.Message}");
            throw;
        }

        // Validate checks the expected records and markers in each output's HTML text content, and that both outputs
        // carry the same text tokens and table cells.
        BenchmarkOutputValidationLog.Passed("OfficeIMO", Describe(report.OfficeIMO));
        BenchmarkOutputValidationLog.Passed("RtfPipe", Describe(report.RtfPipe));
    }

    private string Describe(RtfHtmlOutputEvidence evidence) =>
        $"Scale={Scale}; records={evidence.RecordCount}; tables={evidence.TableCount}; cells={evidence.CellCount}; " +
        $"text tokens={evidence.SemanticTokenCount}; token SHA-256={evidence.SemanticSha256}";

    /// <summary>Parses RTF and renders HTML through OfficeIMO's public format owners.</summary>
    [Benchmark(Baseline = true, Description = "OfficeIMO")]
    [BenchmarkCategory("library:OfficeIMO", "case:Default")]
    public string OfficeIMO() => RtfDocument.ParseResult(_rtf).Document.ToHtml();

    /// <summary>Parses the same RTF and renders HTML through RtfPipe.</summary>
    [Benchmark(Description = "RtfPipe")]
    [BenchmarkCategory("library:RtfPipe", "case:Default")]
    public string RtfPipe() => global::RtfPipe.Rtf.ToHtml(_rtf);
}
