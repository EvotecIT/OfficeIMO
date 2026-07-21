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
    private HtmlPdfSaveOptions _pdfOptions = null!;
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
        _pdfOptions = new HtmlPdfSaveOptions {
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
    public byte[] ExportRenderedPdf() => _document.ToPdf(_pdfOptions);
}

internal static class HtmlBenchmarkCorpus {
    internal static HtmlRenderOptions CreateContinuousOptions() => new HtmlRenderOptions {
        Mode = HtmlRenderMode.Continuous,
        ViewportWidth = 816D,
        Margins = HtmlRenderMargins.All(36D),
        BackgroundColor = OfficeColor.White
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
}
