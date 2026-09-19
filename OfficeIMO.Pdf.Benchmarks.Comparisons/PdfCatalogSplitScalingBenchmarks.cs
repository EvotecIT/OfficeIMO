using BenchmarkDotNet.Attributes;
using OfficeIMO.Pdf;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

public enum PdfCatalogFeatureMode {
    None,
    SingleRange,
    PerPageRange,
    PerPageBookmark
}

/// <summary>Measures every-page split with and without page labels across page-count boundaries.</summary>
[MemoryDiagnoser]
public class PdfCatalogSplitScalingBenchmarks {
    private byte[] _source = null!;

    [Params(5, 6, 50, 500)]
    public int PageCount { get; set; }

    [Params(PdfCatalogFeatureMode.None, PdfCatalogFeatureMode.SingleRange, PdfCatalogFeatureMode.PerPageRange, PdfCatalogFeatureMode.PerPageBookmark)]
    public PdfCatalogFeatureMode FeatureMode { get; set; }

    [GlobalSetup]
    public void Setup() {
        bool hasPageLabels = FeatureMode == PdfCatalogFeatureMode.SingleRange || FeatureMode == PdfCatalogFeatureMode.PerPageRange;
        var options = new PdfOptions { IncludePageLabels = FeatureMode == PdfCatalogFeatureMode.SingleRange };
        if (FeatureMode == PdfCatalogFeatureMode.PerPageRange) {
            for (int page = 1; page <= PageCount; page++) {
                options.AddPageLabelRange(page, PdfPageNumberStyle.Arabic, startNumber: page);
            }
        }

        PdfDocument document = PdfDocument.Create(pdf => pdf.Content(content => {
            for (int page = 1; page <= PageCount; page++) {
                if (page > 1) {
                    content.PageBreak();
                }

                if (FeatureMode == PdfCatalogFeatureMode.PerPageBookmark) {
                    content.Bookmark("Bookmark " + page.ToString("D4", System.Globalization.CultureInfo.InvariantCulture));
                }
                content.Paragraph(paragraph => paragraph.Text(Marker(page)));
            }
        }), options);

        _source = document.ToBytes();
        PdfDocumentPreflight preflight = PdfDocument.Preflight(_source);
        if (!preflight.CanRewrite || preflight.DocumentInfo?.PageCount != PageCount ||
            preflight.Probe.HasPageLabels != hasPageLabels ||
            preflight.Probe.HasNamedDestinations != (FeatureMode == PdfCatalogFeatureMode.PerPageBookmark)) {
            throw new InvalidOperationException("Catalog split setup did not produce the expected rewrite-safe source.");
        }

        byte[][] outputs = Split();
        if (outputs.Length != PageCount) {
            throw new InvalidOperationException("Catalog split output count differs from the source page count.");
        }

        for (int index = 0; index < outputs.Length; index++) {
            PdfReadDocument readback = PdfReadDocument.Open(outputs[index]);
            if (readback.Pages.Count != 1 ||
                !readback.ExtractText().Contains(Marker(index + 1), StringComparison.Ordinal) ||
                (readback.PageLabels.Count != 0) != hasPageLabels ||
                (hasPageLabels && readback.PageLabels[0].StartNumber != index + 1)) {
                throw new InvalidOperationException("Catalog split did not preserve page content and labels at page " + (index + 1) + ".");
            }
            if (FeatureMode == PdfCatalogFeatureMode.PerPageBookmark) {
                PdfDocumentInfo outputInfo = PdfDocument.Load(outputs[index]).Inspect();
                if (outputInfo.NamedDestinations.Count != 1 ||
                    outputInfo.NamedDestinations[0].Name != "Bookmark " + (index + 1).ToString("D4", System.Globalization.CultureInfo.InvariantCulture) ||
                    outputInfo.NamedDestinations[0].PageNumber != 1) {
                    throw new InvalidOperationException("Catalog split did not preserve the expected bookmark at page " + (index + 1) + ".");
                }
            }
        }
    }

    [Benchmark]
    public byte[][] Split() => PdfDocument.Load(_source).Pages.Split().Select(static page => page.ToBytes()).ToArray();

    private static string Marker(int page) => "Catalog split page " + page.ToString("D4", System.Globalization.CultureInfo.InvariantCulture);
}
