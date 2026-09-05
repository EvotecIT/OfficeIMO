using AngleSharp.Html.Dom;
using BenchmarkDotNet.Attributes;
using OfficeIMO.Drawing;
using OfficeIMO.Html.Pdf;

namespace OfficeIMO.Html.Benchmarks;

/// <summary>Measures parse, computed-style, and combined layout scaling over deterministic report markup.</summary>
[MemoryDiagnoser]
[BenchmarkCategory("HTML", "Stages")]
public class HtmlRenderingStageBenchmarks {
    private HtmlComputedStyleSet _computedStyles = null!;
    private IHtmlDocument _document = null!;
    private OfficeFontFaceCollection _fonts = null!;
    private string _html = string.Empty;
    private HtmlRenderOptions _options = null!;
    private HtmlCssPageRuleSet _pageRules = null!;
    private HtmlResourceSession _resources = null!;

    [Params(10, 100)]
    public int RowCount { get; set; }

    [GlobalSetup]
    public void Setup() {
        _html = HtmlBenchmarkCorpus.BuildReport(RowCount);
        _document = HtmlDocumentParser.ParseDocument(_html);
        _options = HtmlBenchmarkCorpus.CreateContinuousOptions();
        _computedStyles = HtmlComputedStyleEngine.ComputeForRendering(
            _document,
            _options,
            HtmlConversionLimits.CreateUntrustedProfile());
        _fonts = new OfficeFontFaceCollection();
        _pageRules = new HtmlCssPageRuleSet();
        _resources = new HtmlResourceSession();
    }

    [Benchmark]
    public IHtmlDocument Parse() => HtmlDocumentParser.ParseDocument(_html);

    [Benchmark]
    public IReadOnlyDictionary<AngleSharp.Dom.IElement, HtmlComputedStyle> ComputeStyles() =>
        HtmlComputedStyleEngine.Compute(_document, HtmlCssMediaContext.Screen);

    [Benchmark]
    public HtmlRenderDocument LayoutFromComputedStyles() => new HtmlRenderLayoutEngine(
        _document,
        _computedStyles,
        _options.Clone(),
        new HtmlDiagnosticReport(),
        _resources,
        _pageRules,
        _fonts).Render();

    [Benchmark]
    public HtmlRenderDocument ParseStyleAndLayout() => HtmlRenderEngine.Render(HtmlConversionDocument.Parse(_html), _options);
}

/// <summary>Measures shared-scene projection to Drawing, PNG, SVG, and rendered searchable PDF.</summary>
[MemoryDiagnoser]
[BenchmarkCategory("HTML", "Outputs")]
public class HtmlRenderingOutputBenchmarks {
    private OfficeDrawing _drawing = null!;
    private HtmlConversionDocument _document = null!;
    private HtmlRenderOptions _imageOptions = null!;
    private HtmlToPdfOptions _pdfOptions = null!;
    private HtmlRenderPage _renderedPage = null!;

    [Params(false, true)]
    public bool UnicodeText { get; set; }

    [GlobalSetup]
    public void Setup() {
        _document = HtmlConversionDocument.Parse(HtmlBenchmarkCorpus.BuildReport(40, UnicodeText));
        _imageOptions = HtmlBenchmarkCorpus.CreateContinuousOptions();
        HtmlRenderDocument rendered = HtmlRenderEngine.Render(_document, _imageOptions);
        _renderedPage = rendered.Pages[0];
        _drawing = _renderedPage.CreateDrawing();
        _pdfOptions = new HtmlToPdfOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(8.5D, 11D),
            Margins = HtmlRenderMargins.All(36D)
        };
    }

    [Benchmark]
    public OfficeDrawing PrepareDrawing() => _renderedPage.CreateDrawing();

    [Benchmark]
    public byte[] ExportPng() => OfficeDrawingRasterRenderer.ToPng(_drawing, 1D, OfficeColor.White);

    [Benchmark]
    public string ExportSvg() => OfficeDrawingSvgExporter.ToSvg(_drawing);

    [Benchmark]
    public byte[] ExportRenderedPdf() => _document.ToPdfBytes(_pdfOptions);
}

/// <summary>Measures realistic paged purchase-table scaling without retaining the final PDF artifact.</summary>
[MemoryDiagnoser]
[BenchmarkCategory("HTML", "PagedTables")]
public class HtmlPagedPurchaseTableBenchmarks {
    private HtmlConversionDocument _document = null!;
    private HtmlRenderOptions _renderOptions = null!;
    private HtmlToPdfOptions _pdfOptions = null!;

    [Params(250, 2500)]
    public int RowCount { get; set; }

    [GlobalSetup]
    public void Setup() {
        _document = HtmlConversionDocument.Parse(HtmlBenchmarkCorpus.BuildPurchaseTable(RowCount));
        _renderOptions = HtmlBenchmarkCorpus.CreatePagedOptions();
        _pdfOptions = new HtmlToPdfOptions(_renderOptions) {
            PdfOptions = new OfficeIMO.Pdf.PdfOptions {
                FileVersion = OfficeIMO.Pdf.PdfFileVersion.Pdf17,
                ObjectSerializationMode = OfficeIMO.Pdf.PdfObjectSerializationMode.ForwardOnly,
                TaggedStructureMode = OfficeIMO.Pdf.PdfTaggedStructureMode.CatalogMarkers
            }
        };
    }

    [Benchmark]
    public HtmlRenderDocument LayoutPaged() => HtmlRenderEngine.Render(_document, _renderOptions);

    [Benchmark]
    public long RenderPdfToForwardOnlyStream() => _document.SaveAsPdf(Stream.Null, _pdfOptions).RequireSuccess().BytesWritten;
}

/// <summary>Measures 100-page and 1,000-page legal-document style workloads.</summary>
[MemoryDiagnoser]
[BenchmarkCategory("HTML", "LongDocuments")]
public class HtmlLongDocumentBenchmarks {
    private HtmlComputedStyleSet _computedStyles = null!;
    private HtmlConversionDocument _document = null!;
    private OfficeFontFaceCollection _fonts = null!;
    private IHtmlDocument _htmlDocument = null!;
    private HtmlRenderOptions _renderOptions = null!;
    private HtmlToPdfOptions _pdfOptions = null!;
    private HtmlCssPageRuleSet _pageRules = null!;
    private HtmlResourceSession _resources = null!;

    [Params(100, 1000)]
    public int PageCount { get; set; }

    [GlobalSetup]
    public void Setup() {
        string html = HtmlBenchmarkCorpus.BuildLongDocument(PageCount);
        _document = HtmlConversionDocument.Parse(html);
        _htmlDocument = HtmlDocumentParser.ParseDocument(html);
        _renderOptions = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = OfficePageSizes.Letter,
            Margins = HtmlRenderMargins.All(48D),
            MaxPageCount = PageCount
        };
        _pdfOptions = new HtmlToPdfOptions(_renderOptions) {
            PdfOptions = new OfficeIMO.Pdf.PdfOptions {
                FileVersion = OfficeIMO.Pdf.PdfFileVersion.Pdf17,
                ObjectSerializationMode = OfficeIMO.Pdf.PdfObjectSerializationMode.ForwardOnly,
                TaggedStructureMode = OfficeIMO.Pdf.PdfTaggedStructureMode.CatalogMarkers
            }
        };
        _pageRules = HtmlCssPageSettingsResolver.Apply(_htmlDocument, _renderOptions, new HtmlDiagnosticReport());
        _computedStyles = HtmlComputedStyleEngine.ComputeForRendering(
            _htmlDocument,
            _renderOptions,
            HtmlConversionLimits.CreateUntrustedProfile());
        _fonts = new OfficeFontFaceCollection();
        _resources = new HtmlResourceSession();
        RequireLongDocumentContract(HtmlRenderEngine.Render(_document, _renderOptions));
    }

    [Benchmark]
    public IReadOnlyDictionary<AngleSharp.Dom.IElement, HtmlComputedStyle> ComputeStyles() =>
        HtmlComputedStyleEngine.Compute(_htmlDocument, HtmlCssMediaContext.Print);

    [Benchmark]
    public HtmlRenderDocument LayoutFromComputedStyles() => RequirePageCount(new HtmlRenderLayoutEngine(
            _htmlDocument,
            _computedStyles,
            _renderOptions.Clone(),
            new HtmlDiagnosticReport(),
            _resources,
            _pageRules,
            _fonts).Render());

    [Benchmark]
    public HtmlRenderDocument LayoutPaged() => RequirePageCount(HtmlRenderEngine.Render(_document, _renderOptions));

    [Benchmark]
    public long RenderPdfToForwardOnlyStream() {
        OfficeIMO.Pdf.PdfSaveResult saved = _document.SaveAsPdf(Stream.Null, _pdfOptions).RequireSuccess();
        if (saved.Serialization?.PageCount != PageCount) {
            throw new InvalidOperationException($"Expected {PageCount} rendered PDF pages but observed {saved.Serialization?.PageCount}.");
        }
        return saved.BytesWritten;
    }

    private HtmlRenderDocument RequirePageCount(HtmlRenderDocument rendered) {
        if (rendered.Pages.Count != PageCount) {
            throw new InvalidOperationException($"Expected {PageCount} rendered pages but observed {rendered.Pages.Count}.");
        }
        return rendered;
    }

    private HtmlRenderDocument RequireLongDocumentContract(HtmlRenderDocument rendered) {
        RequirePageCount(rendered);
        string text = rendered.Text;
        var markers = new HashSet<int>();
        int cursor = 0;
        while ((cursor = text.IndexOf("PAGE-", cursor, StringComparison.Ordinal)) >= 0) {
            int digits = cursor + 5;
            if (digits + 4 > text.Length
                || !int.TryParse(text.Substring(digits, 4), out int marker)
                || !markers.Add(marker)) {
                throw new InvalidOperationException($"Long-document marker at text offset {cursor} was malformed or duplicated.");
            }
            cursor = digits + 4;
        }
        if (markers.Count != PageCount || Enumerable.Range(0, PageCount).Any(marker => !markers.Contains(marker))) {
            throw new InvalidOperationException($"Expected {PageCount} unique long-document markers but observed {markers.Count}.");
        }
        return rendered;
    }
}

/// <summary>Measures the strict managed standards packet across layout and searchable PDF output.</summary>
[MemoryDiagnoser]
[BenchmarkCategory("HTML", "StaticStandards")]
public class HtmlStaticStandardsBenchmarks {
    private HtmlConversionDocument _document = null!;
    private HtmlRenderOptions _renderOptions = null!;
    private HtmlToPdfOptions _pdfOptions = null!;

    [GlobalSetup]
    public void Setup() {
        _document = HtmlConversionDocument.Parse(HtmlBenchmarkCorpus.BuildStaticStandardsShowcase());
        _renderOptions = HtmlBenchmarkCorpus.CreateStaticStandardsOptions();
        _pdfOptions = new HtmlToPdfOptions(_renderOptions) {
            PdfOptions = new OfficeIMO.Pdf.PdfOptions {
                FileVersion = OfficeIMO.Pdf.PdfFileVersion.Pdf17,
                ObjectSerializationMode = OfficeIMO.Pdf.PdfObjectSerializationMode.ForwardOnly,
                TaggedStructureMode = OfficeIMO.Pdf.PdfTaggedStructureMode.CatalogMarkers
            }
        };

        RequireStandardsContract(HtmlRenderEngine.Render(_document, _renderOptions));
        byte[] pdf = _document.ToPdfBytes(_pdfOptions);
        OfficeIMO.Pdf.PdfReadDocument readDocument = OfficeIMO.Pdf.PdfReadDocument.Open(pdf);
        string text = readDocument.ExtractText();
        if (readDocument.Pages.Count != 2
            || !text.Contains("Static standards packet", StringComparison.Ordinal)
            || !text.Contains("Second-page evidence", StringComparison.Ordinal)) {
            throw new InvalidOperationException("The static-standards PDF validation did not preserve the two-page searchable contract.");
        }
    }

    [Benchmark]
    public HtmlRenderDocument LayoutStrictPaged() => RequireStandardsContract(
        HtmlRenderEngine.Render(_document, _renderOptions));

    [Benchmark]
    public long RenderTaggedPdfToForwardOnlyStream() =>
        _document.SaveAsPdf(Stream.Null, _pdfOptions).RequireSuccess().RequireNoLoss().BytesWritten;

    private static HtmlRenderDocument RequireStandardsContract(HtmlRenderDocument rendered) {
        rendered.RequireNoLoss();
        if (rendered.Pages.Count != 2
            || !rendered.Text.Contains("Static standards packet", StringComparison.Ordinal)
            || !rendered.Text.Contains("Second-page evidence", StringComparison.Ordinal)) {
            throw new InvalidOperationException("The static-standards layout validation did not preserve the expected page and text contract.");
        }
        return rendered;
    }
}

internal static class HtmlBenchmarkCorpus {
    internal static HtmlRenderOptions CreateContinuousOptions() => new HtmlRenderOptions {
        Mode = HtmlRenderMode.Continuous,
        ViewportWidth = 816D,
        Margins = HtmlRenderMargins.All(36D),
        BackgroundColor = OfficeColor.White
    };

    internal static HtmlRenderOptions CreatePagedOptions() => new HtmlRenderOptions {
        Mode = HtmlRenderMode.Paged,
        PageSize = new OfficePageSize(5D, 4D),
        Margins = HtmlRenderMargins.All(32D),
        BackgroundColor = OfficeColor.White
    };

    internal static HtmlRenderOptions CreateStaticStandardsOptions() => new HtmlRenderOptions {
        Mode = HtmlRenderMode.Paged,
        PageSize = new OfficePageSize(5D, 4D),
        Margins = HtmlRenderMargins.All(32D),
        BackgroundColor = OfficeColor.White,
        FidelityPolicy = HtmlRenderFidelityPolicy.RequireNoLoss
    };

    internal static string BuildReport(int rowCount, bool includeUnicodeText = false) {
        var html = new System.Text.StringBuilder(rowCount * 90 + 1024);
        html.Append("<article><style>body{font-family:Arial}table{width:100%;border-collapse:collapse}th,td{border:1px solid #778;padding:4px}.summary{display:flex;gap:12px}.card{padding:10px;background:#eef4ff}</style>")
            .Append(includeUnicodeText ? "<h1>Benchmark Report Ω Ж שלום سلام</h1>" : "<h1>Benchmark Report</h1>")
            .Append("<div class='summary'><section class='card'><h2>Rows</h2><p>")
            .Append(rowCount)
            .Append("</p></section><section class='card'><h2>Status</h2><p>Ready</p></section></div>")
            .Append("<table><thead><tr><th>Id</th><th>Name</th><th>Amount</th></tr></thead><tbody>");
        for (int index = 0; index < rowCount; index++) {
            html.Append("<tr><td>").Append(index).Append("</td><td>Line ").Append(index)
                .Append("</td><td>").Append(index * 17).Append(".25</td></tr>");
        }
        return html.Append("</tbody></table></article>").ToString();
    }

    internal static string BuildPurchaseTable(int rowCount) {
        var html = new System.Text.StringBuilder(rowCount * 220 + 2048);
        html.Append("<style>@page{size:5in 4in;margin:.42in .35in;@top-center{content:'Purchase report';font:9px Arial}@bottom-right{content:'Page ' counter(page) ' of ' counter(pages);font:8px Arial}}")
            .Append("body{margin:0;font:9px/1.35 Arial;color:#172033}table{width:100%;border-collapse:collapse;table-layout:fixed}thead{display:table-header-group}")
            .Append("th,td{border:1px solid #94a3b8;padding:4px;vertical-align:top;overflow-wrap:anywhere}th{background:#dbeafe;text-align:left}th:first-child,td:first-child{width:62%}.number{text-align:right;white-space:nowrap}")
            .Append(".totals{break-inside:avoid;margin:10px 0 0 auto;width:180px;border-top:2px solid #2563eb;padding-top:5px}</style>")
            .Append("<main><h1>Purchase statement</h1><table><thead><tr><th>Item description</th><th>Qty</th><th class='number'>Rate</th><th class='number'>Amount</th></tr></thead><tbody>");
        decimal subtotal = 0m;
        for (int index = 0; index < rowCount; index++) {
            int quantity = index % 4 + 1;
            decimal amount = quantity * 19.95m;
            subtotal += amount;
            html.Append("<tr><td><strong>SKU-")
                .Append(index.ToString("D5"))
                .Append("</strong><br>Precision managed document service with audit evidence and regional reporting</td><td>")
                .Append(quantity)
                .Append("</td><td class='number'>$19.95</td><td class='number'>$")
                .Append(amount.ToString("N2", System.Globalization.CultureInfo.InvariantCulture))
                .Append("</td></tr>");
        }
        decimal tax = decimal.Round(subtotal * 0.08m, 2);
        return html.Append("</tbody></table><section class='totals'>Subtotal $")
            .Append(subtotal.ToString("N2", System.Globalization.CultureInfo.InvariantCulture))
            .Append("<br>Tax $")
            .Append(tax.ToString("N2", System.Globalization.CultureInfo.InvariantCulture))
            .Append("<br><strong>Total $")
            .Append((subtotal + tax).ToString("N2", System.Globalization.CultureInfo.InvariantCulture))
            .Append("</strong></section></main>")
            .ToString();
    }

    internal static string BuildLongDocument(int pageCount) {
        var html = new System.Text.StringBuilder(pageCount * 900 + 1024);
        html.Append("<style>@page{size:letter;margin:.65in;@top-center{content:'Managed legal packet';font:9px Arial}@bottom-right{content:'Page ' counter(page) ' of ' counter(pages);font:8px Arial}}")
            .Append("body{margin:0;font:11px/1.45 Arial;color:#172033}section{break-after:page}section:last-child{break-after:auto}h1{font-size:18px}p{margin:0 0 9px}</style><main>");
        const string paragraph = "This agreement records the service scope, delivery controls, evidence requirements, retention schedule, review obligations, and remedies for the parties. ";
        for (int pageIndex = 0; pageIndex < pageCount; pageIndex++) {
            html.Append("<section><h1>Article ")
                .Append(pageIndex + 1)
                .Append("</h1><p>Document marker PAGE-")
                .Append(pageIndex.ToString("D4"))
                .Append(".</p><p>")
                .Append(paragraph)
                .Append(paragraph)
                .Append(paragraph)
                .Append("</p><p>")
                .Append(paragraph)
                .Append(paragraph)
                .Append("</p></section>");
        }
        return html.Append("</main>").ToString();
    }

    internal static string BuildStaticStandardsShowcase() => """
        <style>
          @page{size:5in 4in;margin:44px 32px 38px;@top-center{content:element(doc-header,first)}@bottom-right{content:"Page " counter(page) " / " counter(pages)}}
          body{margin:0;font:12px/1.4 Arial,sans-serif;color:#172033}.running{position:running(doc-header);margin:0;padding:4px;border-bottom:1px solid #315b8a;color:#315b8a}
          section{break-after:page;-officeimo-pdf-tag-type:Sect}section:last-child{break-after:auto}h1{bookmark-level:1;bookmark-state:open}h2{bookmark-level:2}
          .matrix{display:grid;grid-template-columns:1fr 1fr;grid-template-rows:34px 52px;gap:8px}.sub{display:grid;grid-column:1;grid-row:1/3;grid-template-rows:subgrid;row-gap:inherit}
          .cell{padding:7px;background:linear-gradient(135deg,#edf4ff,#fff);border:1px solid #b9c9dc}.badge{clip-path:polygon(0 0,100% 0,88% 100%,0 100%);background:#315b8a;color:white;padding:7px}
          svg{width:150px;height:62px}path{fill:none;stroke:#2a7a58;stroke-width:4}
        </style>
        <header class="running">Managed static standards · vector output</header>
        <section><h1>Static standards packet</h1><p>Deterministic page one content remains searchable and tagged.</p>
          <div class="matrix"><div class="sub"><div class="cell">Inherited row A</div><div class="cell">Inherited row B</div></div><div class="badge">Clipped vector badge</div><div class="cell">Named page evidence</div></div>
          <svg viewBox="0 0 150 62" role="img" aria-label="Vector trend"><path d="M5 52 L42 31 L79 40 L113 15 L145 7"/></svg>
        </section>
        <section><h2>Second-page evidence</h2><p>Running elements, counters, outlines, and artifact tagging remain deterministic.</p></section>
        """;
}
