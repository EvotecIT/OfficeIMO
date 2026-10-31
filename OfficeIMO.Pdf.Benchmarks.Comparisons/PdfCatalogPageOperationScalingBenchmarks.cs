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
        var options = new PdfCreateOptions { IncludePageLabels = FeatureMode == PdfCatalogFeatureMode.SingleRange };
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
        using var input = new MemoryStream(_source, writable: false);
        using PdfDocument loaded = PdfDocument.Load(input);
        PdfDocumentPreflight preflight = loaded.Analyze().Preflight;
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
            using var readbackStream = new MemoryStream(outputs[index], writable: false);
            using PdfDocument readback = PdfDocument.Load(readbackStream);
            PdfDocumentInfo readbackInfo = readback.Inspect();
            if (readback.Pages.Count != 1 ||
                !readback.ExtractText().Contains(Marker(index + 1), StringComparison.Ordinal) ||
                (readbackInfo.PageLabels.Count != 0) != hasPageLabels ||
                (hasPageLabels && readbackInfo.PageLabels[0].StartNumber != index + 1)) {
                throw new InvalidOperationException("Catalog split did not preserve page content and labels at page " + (index + 1) + ".");
            }
            if (FeatureMode == PdfCatalogFeatureMode.PerPageBookmark) {
                PdfDocumentInfo outputInfo = readbackInfo;
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

        using var selectedStream = new MemoryStream(Select(), writable: false);
        using PdfDocument selected = PdfDocument.Load(selectedStream);
        PdfDocumentInfo selectedInfo = selected.Inspect();
        if (selected.Pages.Count != selectedCount) {
            throw new InvalidOperationException("Catalog selection output count differs from the planned page count.");
        }
        for (int index = 0; index < selectedCount; index++) {
            if (!selected.Pages[index].ExtractText().Contains(Marker(_selectedPages[index]), StringComparison.Ordinal)) {
                throw new InvalidOperationException("Catalog selection changed page content or order at output page " + (index + 1) + ".");
            }
        }
        if (hasPageLabels && (selectedInfo.PageLabels.Count != selectedCount ||
            selectedInfo.PageLabels.Where((label, index) => label.StartPageIndex != index || label.StartNumber != _selectedPages[index]).Any())) {
            throw new InvalidOperationException("Catalog selection did not preserve the selected page labels.");
        }
        if (!hasPageLabels && selectedInfo.PageLabels.Count != 0) {
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
    public byte[][] Split() {
        using var input = new MemoryStream(_source, writable: false);
        using PdfDocument source = PdfDocument.Load(input);
        IReadOnlyList<PdfDocument> pages = source.Pages.Split();
        try {
            return pages.Select(static page => page.ToBytes()).ToArray();
        } finally {
            foreach (PdfDocument page in pages) page.Dispose();
        }
    }

    [Benchmark]
    public byte[] Select() {
        using var input = new MemoryStream(_source, writable: false);
        using PdfDocument source = PdfDocument.Load(input);
        using PdfDocument output = source.Pages.Extract(_selectedPages);
        return output.ToBytes();
    }

    private static string Marker(int page) => "Catalog split page " + page.ToString("D4", System.Globalization.CultureInfo.InvariantCulture);
}
