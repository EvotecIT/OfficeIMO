using BenchmarkDotNet.Attributes;
using OfficeIMO.Pdf;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

public enum PdfCatalogFeatureMode {
    None,
    SingleRange,
    PerPageRange,
    PerPageBookmark
}

/// <summary>Measures split and non-contiguous selection across page-count and catalog-feature boundaries.</summary>
[MemoryDiagnoser]
public class PdfCatalogPageOperationScalingBenchmarks {
    private byte[] _source = null!;
    private int[] _selectedPages = null!;

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

        int selectedCount = Math.Max(2, PageCount / 4);
        _selectedPages = new int[selectedCount];
        for (int index = 0; index < selectedCount; index++) {
            _selectedPages[index] = PageCount - (int)Math.Round(index * (PageCount - 1D) / (selectedCount - 1D));
        }

        PdfReadDocument selected = PdfReadDocument.Open(Select());
        if (selected.Pages.Count != selectedCount) {
            throw new InvalidOperationException("Catalog selection output count differs from the planned page count.");
        }
        for (int index = 0; index < selectedCount; index++) {
            if (!selected.Pages[index].ExtractText().Contains(Marker(_selectedPages[index]), StringComparison.Ordinal)) {
                throw new InvalidOperationException("Catalog selection changed page content or order at output page " + (index + 1) + ".");
            }
        }
        if (hasPageLabels && (selected.PageLabels.Count != selectedCount ||
            selected.PageLabels.Where((label, index) => label.StartPageIndex != index || label.StartNumber != _selectedPages[index]).Any())) {
            throw new InvalidOperationException("Catalog selection did not preserve the selected page labels.");
        }
        if (!hasPageLabels && selected.PageLabels.Count != 0) {
            throw new InvalidOperationException("Catalog selection added page labels to an unlabelled source.");
        }
        if (FeatureMode == PdfCatalogFeatureMode.PerPageBookmark) {
            IReadOnlyList<PdfNamedDestination> destinations = selected.NamedDestinations;
            if (destinations.Count != selectedCount ||
                destinations.Any(destination => destination.PageNumber is not int outputPage ||
                    outputPage < 1 || outputPage > selectedCount ||
                    destination.Name != "Bookmark " + _selectedPages[outputPage - 1].ToString("D4", System.Globalization.CultureInfo.InvariantCulture))) {
                throw new InvalidOperationException("Catalog selection did not preserve and remap selected bookmarks.");
            }
        }
    }

    [Benchmark]
    public byte[][] Split() => PdfDocument.Load(_source).Pages.Split().Select(static page => page.ToBytes()).ToArray();

    [Benchmark]
    public byte[] Select() => PdfDocument.Load(_source).Pages.Extract(_selectedPages).ToBytes();

    private static string Marker(int page) => "Catalog split page " + page.ToString("D4", System.Globalization.CultureInfo.InvariantCulture);
}
