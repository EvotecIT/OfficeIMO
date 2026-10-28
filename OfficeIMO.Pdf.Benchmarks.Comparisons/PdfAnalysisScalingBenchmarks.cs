using BenchmarkDotNet.Attributes;
using OfficeIMO.Pdf;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

/// <summary>
/// Measures consolidated health, capability, optimization, signature, and append-only analysis
/// across page-count boundaries. Results are not comparable to former isolated-preflight timings.
/// </summary>
[MemoryDiagnoser]
public class PdfAnalysisScalingBenchmarks {
    private byte[] _source = null!;

    [Params(5, 6, 50, 500)]
    public int PageCount { get; set; }

    [Params(false, true)]
    public bool FeatureRich { get; set; }

    [GlobalSetup]
    public void Setup() {
        var options = new PdfCreateOptions {
            IncludePageLabels = FeatureRich,
            IncludeXmpMetadata = FeatureRich,
            ViewerPreferences = FeatureRich
                ? new PdfViewerPreferencesOptions { DisplayDocTitle = true }
                : null
        };
        using PdfDocument document = PdfDocument.Create(pdf => pdf.Content(content => {
            for (int page = 1; page <= PageCount; page++) {
                if (page > 1) {
                    content.PageBreak();
                }

                content.Paragraph(paragraph => paragraph.Text("Page " + page + " preflight content."));
            }
        }), options).Meta(title: "Preflight scaling");

        _source = document.ToBytes();
        using var input = new MemoryStream(_source, writable: false);
        using PdfDocument loaded = PdfDocument.Load(input);
        PdfDocumentPreflight result = loaded.Analyze().Preflight;
        if (!result.CanRead || !result.CanRewrite || result.DocumentInfo?.PageCount != PageCount ||
            result.Probe.HasPageLabels != FeatureRich ||
            result.Probe.HasXmpMetadata != FeatureRich ||
            result.Probe.HasViewerPreferences != FeatureRich) {
            throw new InvalidOperationException("Preflight scaling setup did not produce the expected readable, rewrite-safe document and catalog features.");
        }
    }

    [Benchmark]
    public PdfAnalysisReport Analyze() {
        using var input = new MemoryStream(_source, writable: false);
        using PdfDocument document = PdfDocument.Load(input);
        return document.Analyze();
    }
}
