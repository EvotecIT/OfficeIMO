using System.Text;
using BenchmarkDotNet.Attributes;
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
        RtfHtmlComparisonValidation.Validate(Scale, _rtf);
    }

    /// <summary>Parses RTF and renders HTML through OfficeIMO's public format owners.</summary>
    [Benchmark(Baseline = true, Description = "OfficeIMO")]
    public string OfficeIMO() => RtfDocument.Read(_rtf).Document.ToHtml();

    /// <summary>Parses the same RTF and renders HTML through RtfPipe.</summary>
    [Benchmark(Description = "RtfPipe")]
    public string RtfPipe() => global::RtfPipe.Rtf.ToHtml(_rtf);
}
