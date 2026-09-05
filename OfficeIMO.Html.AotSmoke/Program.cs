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

Console.WriteLine("OfficeIMO HTML NativeAOT smoke passed.");
