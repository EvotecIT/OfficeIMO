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
byte[] pdf = source.ToPdf(new HtmlPdfSaveOptions(imageOptions));
string extractedText = PdfReadDocument.Open(pdf).ExtractText();

if (!svg.Contains(marker, StringComparison.Ordinal)) throw new InvalidOperationException("The NativeAOT SVG output lost searchable text.");
if (png.Length < 8 || png[0] != 137 || png[1] != 80) throw new InvalidOperationException("The NativeAOT PNG output is invalid.");
if (!extractedText.Contains(marker, StringComparison.Ordinal)) throw new InvalidOperationException("The NativeAOT PDF output lost searchable text.");

Console.WriteLine("OfficeIMO HTML NativeAOT smoke passed.");
