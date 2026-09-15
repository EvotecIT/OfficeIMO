using AngleSharp.Css.Parser;
using BenchmarkDotNet.Attributes;
using OfficeIMO.Html.Dom;
using NativeHtmlDocument = AngleSharp.Html.Dom.IHtmlDocument;

namespace OfficeIMO.Html.Benchmarks;

/// <summary>Measures the retained HTML provider and each OfficeIMO-owned document entry point.</summary>
[MemoryDiagnoser]
[BenchmarkCategory("HTML", "Providers", "Parsing")]
public class HtmlProviderParsingBenchmarks {
    private string _html = string.Empty;

    [Params(10, 100)]
    public int RowCount { get; set; }

    [GlobalSetup]
    public void Setup() => _html = HtmlBenchmarkCorpus.BuildReport(RowCount);

    [Benchmark(Baseline = true)]
    public NativeHtmlDocument AngleSharpNativeDocument() => HtmlDocumentParser.ParseDocument(_html);

    [Benchmark]
    public HtmlDocument OfficeIMOOwnedDocument() => HtmlDocumentEngine.Default.ParseDocument(_html);

    [Benchmark]
    public HtmlConversionDocument ConversionDocumentNativeGraph() => HtmlConversionDocument.Parse(_html);

    [Benchmark]
    public HtmlConversionDocument ConversionDocumentWithOwnedGraph() {
        HtmlConversionDocument conversion = HtmlConversionDocument.Parse(_html);
        _ = conversion.Document;
        return conversion;
    }
}

/// <summary>Separates raw AngleSharp.Css parsing from OfficeIMO's owned cascade result.</summary>
[MemoryDiagnoser]
[BenchmarkCategory("HTML", "Providers", "CSS")]
public class HtmlProviderCssBenchmarks {
    private CssParser _cssParser = null!;
    private string _css = string.Empty;
    private HtmlDocument _document = null!;

    [Params(25, 100)]
    public int RuleCount { get; set; }

    [GlobalSetup]
    public void Setup() {
        (_css, string html) = HtmlProviderBenchmarkCorpus.BuildStyledCards(RuleCount);
        _cssParser = new CssParser(new CssParserOptions { IsIncludingUnknownDeclarations = true });
        _document = HtmlDocumentEngine.Default.ParseDocument(html);
    }

    [Benchmark(Baseline = true)]
    public AngleSharp.Css.Dom.ICssStyleSheet AngleSharpCssSyntax() => _cssParser.ParseStyleSheet(_css);

    [Benchmark]
    public IReadOnlyDictionary<HtmlElement, HtmlComputedStyle> OfficeIMOOwnedCascade() =>
        HtmlComputedStyleEngine.Compute(_document, HtmlCssMediaContext.Screen);
}

internal static class HtmlProviderBenchmarkCorpus {
    internal static (string Css, string Html) BuildStyledCards(int ruleCount) {
        var css = new System.Text.StringBuilder(ruleCount * 190 + 1024);
        css.Append("@layer reset,theme,components;@layer reset{*{box-sizing:border-box}}")
            .Append("@layer theme{:root{--accent:#315b8a;--space:6px;color:#172033}}")
            .Append("@supports(display:grid){@media screen and (min-width:300px){.cards{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:var(--space)}}}")
            .Append("@font-face{font-family:Evidence;src:local('Arial');font-style:normal;font-weight:400}");
        for (int index = 0; index < ruleCount; index++) {
            css.Append("@layer components{.card-").Append(index)
                .Append("{font:400 12px/1.4 Evidence,Arial;color:var(--accent);padding:calc(var(--space) + 1px);background:linear-gradient(135deg,#eef4ff,#fff);border:1px solid color-mix(in srgb,var(--accent) 35%,white)}")
                .Append(".card-").Append(index).Append("::before{content:'Card ").Append(index).Append("';font-weight:700}} ");
        }

        var html = new System.Text.StringBuilder(css.Length + ruleCount * 80 + 128);
        html.Append("<style>").Append(css).Append("</style><main class='cards'>");
        for (int index = 0; index < ruleCount; index++) {
            html.Append("<article class='card-").Append(index).Append("'><h2>Item ").Append(index)
                .Append("</h2><p>Provider evidence row</p></article>");
        }
        html.Append("</main>");
        return (css.ToString(), html.ToString());
    }
}
