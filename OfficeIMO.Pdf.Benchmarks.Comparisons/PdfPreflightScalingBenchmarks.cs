using BenchmarkDotNet.Attributes;
using OfficeIMO.Pdf;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

/// <summary>Measures preflight across page-count boundaries with and without catalog features.</summary>
[MemoryDiagnoser]
public class PdfPreflightScalingBenchmarks {
    private byte[] _source = null!;

    [Params(5, 6, 50, 500)]
    public int PageCount { get; set; }

    [Params(false, true)]
    public bool FeatureRich { get; set; }

    [GlobalSetup]
    public void Setup() {
        var options = new PdfOptions {
            IncludePageLabels = FeatureRich,
            IncludeXmpMetadata = FeatureRich,
            ViewerPreferences = FeatureRich
                ? new PdfViewerPreferencesOptions { DisplayDocTitle = true }
                : null
        };
        PdfDocument document = PdfDocument.Create(pdf => pdf.Content(content => {
            for (int page = 1; page <= PageCount; page++) {
                if (page > 1) {
                    content.PageBreak();
                }

                content.Paragraph(paragraph => paragraph.Text("Page " + page + " preflight content."));
            }
        }), options).Meta(title: "Preflight scaling");

        _source = document.ToBytes();
        PdfDocumentPreflight result = PdfDocument.Preflight(_source);
        if (!result.CanRead || !result.CanRewrite || result.DocumentInfo?.PageCount != PageCount ||
            result.Probe.HasPageLabels != FeatureRich ||
            result.Probe.HasXmpMetadata != FeatureRich ||
            result.Probe.HasViewerPreferences != FeatureRich) {
            throw new InvalidOperationException("Preflight scaling setup did not produce the expected readable, rewrite-safe document and catalog features.");
        }
    }

    [Benchmark]
    public PdfDocumentPreflight Preflight() => PdfDocument.Preflight(_source);
}
