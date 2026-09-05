using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Pdf;

const string marker = "AotMarker";
const string html = "<style>body{margin:0}h1{color:#123456}</style><h1>AotMarker</h1><p><a href='https://example.test/'>Searchable PDF link</a></p>";
HtmlConversionDocument source = HtmlConversionDocument.Parse(html);
var imageOptions = new HtmlRenderOptions {
    ViewportWidth = 320D,
    Margins = HtmlRenderMargins.All(12D)
};

string svg = source.ToSvg(imageOptions);
byte[] png = source.ToPng(imageOptions);
byte[] pdf = source.ToPdfBytes(new HtmlToPdfOptions(imageOptions));
string extractedText = PdfReadDocument.Open(pdf).ExtractText();

if (!svg.Contains(marker, StringComparison.Ordinal)) throw new InvalidOperationException("The NativeAOT SVG output lost searchable text.");
if (png.Length < 8 || png[0] != 137 || png[1] != 80) throw new InvalidOperationException("The NativeAOT PNG output is invalid.");
if (!extractedText.Contains(marker, StringComparison.Ordinal)) throw new InvalidOperationException("The NativeAOT PDF output lost searchable text.");

const string standardsHtml = """
<style>
@page { size: 5in 4in; margin: 32px; @top-center { content: element(packet-header); } }
.running { position: running(packet-header); color: #315b8a; border-bottom: 1px solid #315b8a; }
.grid { display: grid; grid-template-columns: 1fr 1fr; grid-template-rows: 32px 32px; gap: 8px; }
.subgrid { display: grid; grid-row: 1 / 3; grid-template-rows: subgrid; row-gap: inherit; }
.badge { grid-column: 2; clip-path: polygon(0 0, 88% 0, 100% 50%, 88% 100%, 0 100%); background: #315b8a; color: white; }
.second { break-before: page; }
</style>
<header class="running">Managed static NativeAOT standards</header>
<main>
  <h1>Static standards page one</h1>
  <section class="grid"><div class="subgrid"><span>Row A</span><span>Row B</span></div><div class="badge">Clipped badge</div></section>
  <section class="second"><h2>Static standards page two</h2><p>Searchable tagged output</p></section>
</main>
""";
HtmlConversionDocument standardsSource = HtmlConversionDocument.Parse(standardsHtml);
var standardsOptions = new HtmlRenderOptions {
    Mode = HtmlRenderMode.Paged,
    PageSize = new OfficePageSize(5D, 4D),
    Margins = HtmlRenderMargins.All(32D),
    FidelityPolicy = HtmlRenderFidelityPolicy.RequireNoLoss
};
string secondPageSvg = standardsSource.ToSvg(standardsOptions, pageIndex: 1);
byte[] standardsPdf = standardsSource.ToPdfBytes(new HtmlToPdfOptions(standardsOptions));
PdfReadDocument standardsReadDocument = PdfReadDocument.Open(standardsPdf);
string standardsText = standardsReadDocument.ExtractText();
if (!secondPageSvg.Contains("Static standards page two", StringComparison.Ordinal)) throw new InvalidOperationException("The NativeAOT paged SVG output lost its second page.");
if (standardsReadDocument.Pages.Count != 2
    || !standardsText.Contains("Static standards page one", StringComparison.Ordinal)
    || !standardsText.Contains("Static standards page two", StringComparison.Ordinal)) {
    throw new InvalidOperationException("The NativeAOT strict static standards packet lost its two-page searchable contract.");
}

const string embeddedSvg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 40 18'><defs>"
    + "<filter id='s'><feDropShadow dx='2' dy='1' stdDeviation='1' flood-color='navy'/></filter></defs>"
    + "<rect x='1' y='1' width='12' height='12' fill='orange' filter='url(#s)'/>"
    + "<foreignObject x='17' y='1' width='20' height='14'><div xmlns='http://www.w3.org/1999/xhtml' "
    + "style='font:8px/12px Arial;color:navy;background:lime'>FxAot</div></foreignObject></svg>";
string advancedHtml = "<style>@page{size:4in 4in;margin:24px;bleed:4px;marks:crop}"
    + "body,p,ul,h2{margin:0}"
    + "h2{text-shadow:1px 1px 1px #99a;writing-mode:vertical-rl;height:80px}"
    + "li::marker,.note::footnote-marker{content:'[' url(\"data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNgYAAAAAMAASsJTYQAAAAASUVORK5CYII=\") ']';color:green}"
    + ".note{float:footnote;font-size:8px;line-height:10px}</style>"
    + "<h2>VerticalAot</h2><ul><li>MarkerAot</li></ul><p>Call<span class='note'>FootnoteAot</span></p>"
    + embeddedSvg.Replace("<svg ", "<svg aria-label='Advanced vector proof' style='width:160px;height:72px' ", StringComparison.Ordinal);
HtmlConversionDocument advancedSource = HtmlConversionDocument.Parse(advancedHtml);
var advancedOptions = new HtmlRenderOptions {
    Mode = HtmlRenderMode.Paged,
    PageSize = new OfficePageSize(4D, 4D),
    Margins = HtmlRenderMargins.All(24D),
    FidelityPolicy = HtmlRenderFidelityPolicy.AllowDiagnosedLoss
};
HtmlRenderDocument advancedRender = HtmlRenderEngine.Render(advancedSource, advancedOptions);
if (advancedRender.Diagnostics.Any(diagnostic =>
        diagnostic.Code is HtmlRenderDiagnosticCodes.SvgContentUnsupported
            or HtmlRenderDiagnosticCodes.SvgRasterFallback
            or HtmlRenderDiagnosticCodes.TextShadowValueUnsupported
            or HtmlRenderDiagnosticCodes.TextShadowLayerLimit)) {
    throw new InvalidOperationException("The NativeAOT advanced HTML packet lost an effects capability: "
        + string.Join(", ", advancedRender.Diagnostics.Select(diagnostic => diagnostic.Code)));
}
string advancedSvg = advancedSource.ToSvg(advancedOptions);
byte[] advancedPdf = advancedSource.ToPdfBytes(new HtmlToPdfOptions(advancedOptions));
string advancedText = PdfReadDocument.Open(advancedPdf).ExtractText();
string compactAdvancedText = string.Concat(advancedText.Where(character => !char.IsWhiteSpace(character)));
if (!compactAdvancedText.Contains("VerticalAot", StringComparison.Ordinal)
    || !compactAdvancedText.Contains("MarkerAot", StringComparison.Ordinal)
    || !compactAdvancedText.Contains("FootnoteAot", StringComparison.Ordinal)
    || !compactAdvancedText.Contains("FxAot", StringComparison.Ordinal)) {
    throw new InvalidOperationException("The NativeAOT advanced HTML/CSS/SVG packet lost searchable content: " + advancedText);
}
if (advancedRender.Pages.Count != 1
    || advancedRender.Pages[0].PrintProduction?.Marks != HtmlRenderPrintMarks.Crop
    || advancedRender.Pages[0].PrintProduction?.Bleed != 4D) {
    throw new InvalidOperationException(
        $"The NativeAOT advanced packet lost its one-page CSS production contract: pages={advancedRender.Pages.Count}, "
        + $"marks={advancedRender.Pages[0].PrintProduction?.Marks}, bleed={advancedRender.Pages[0].PrintProduction?.Bleed}.");
}
if (!advancedSvg.Contains("data:image/png;base64", StringComparison.Ordinal)) {
    throw new InvalidOperationException("The NativeAOT advanced packet lost its generated marker image.");
}

Console.WriteLine("OfficeIMO HTML NativeAOT smoke passed.");
