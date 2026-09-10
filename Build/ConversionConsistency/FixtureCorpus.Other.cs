using System.IO.Compression;
using System.Text;
using HtmlTinkerX;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Visio;

namespace OfficeIMO.ConversionConsistency;

internal static partial class FixtureCorpus {
    private static async Task AddOtherFormatsAsync(string repository, string output, string family, List<ConsistencyCase> cases) {
        const int width = 480, height = 320;
        byte[] font = File.ReadAllBytes(Path.Combine(repository, "OfficeIMO.Drawing.Tests/TestAssets/OpenSans-Regular.woff2"));
        string commonCss = $"@font-face{{font-family:'{family}';src:url(data:font/woff2;base64,{Convert.ToBase64String(font)})}}" +
            $"html,body{{margin:0;padding:0;background:white;font-family:'{family}';font-size:16px;line-height:24px}}" +
            "p{margin:0 0 12px 0}";
        string css = commonCss + "section{height:320px;box-sizing:border-box;padding:24px;break-after:page}section:last-child{break-after:auto}";
        string body = string.Concat(Enumerable.Range(1, 2).Select(page =>
            $"<section><p>SERVICE REVIEW {page}</p><p>Completed actions: 9 of 12</p></section>"));
        string html = "<!doctype html><html><head><meta charset=\"utf-8\"/><style>" + css + "</style></head><body>" + body + "</body></html>";
        string emailHtml = "<!doctype html><html><head><style>" + commonCss + "</style></head><body><p>SERVICE REVIEW 1</p><p>Completed actions: 9 of 12</p></body></html>";
        File.WriteAllText(Path.Combine(output, "service-review.html"), html);
        File.WriteAllText(Path.Combine(output, "service-review.eml"),
            "From: Reports <reports@example.test>\r\nTo: Reader <reader@example.test>\r\n" +
            "Subject: Service review\r\nMIME-Version: 1.0\r\nContent-Type: text/html; charset=utf-8\r\n" +
            "Content-Transfer-Encoding: base64\r\n\r\n" + Convert.ToBase64String(Encoding.UTF8.GetBytes(emailHtml)) + "\r\n");
        using (var archive = ZipFile.Open(Path.Combine(output, "service-review.epub"), ZipArchiveMode.Create)) {
            WriteEntry("mimetype", "application/epub+zip", CompressionLevel.NoCompression);
            WriteEntry("META-INF/container.xml", "<container version=\"1.0\" xmlns=\"urn:oasis:names:tc:opendocument:xmlns:container\"><rootfiles><rootfile full-path=\"content.opf\" media-type=\"application/oebps-package+xml\"/></rootfiles></container>");
            WriteEntry("content.opf", "<package xmlns=\"http://www.idpf.org/2007/opf\" version=\"3.0\" unique-identifier=\"book-id\"><metadata xmlns:dc=\"http://purl.org/dc/elements/1.1/\"><dc:identifier id=\"book-id\">urn:officeimo:consistency</dc:identifier><dc:title>Service review</dc:title><dc:language>en</dc:language><meta property=\"dcterms:modified\">2026-01-01T00:00:00Z</meta></metadata><manifest><item id=\"chapter\" href=\"chapter.xhtml\" media-type=\"application/xhtml+xml\"/><item id=\"nav\" href=\"nav.xhtml\" media-type=\"application/xhtml+xml\" properties=\"nav\"/></manifest><spine><itemref idref=\"chapter\"/></spine></package>");
            WriteEntry("chapter.xhtml", html.Replace("<!doctype html><html>", "<html xmlns=\"http://www.w3.org/1999/xhtml\">"));
            WriteEntry("nav.xhtml", "<html xmlns=\"http://www.w3.org/1999/xhtml\" xmlns:epub=\"http://www.idpf.org/2007/ops\"><head><title>Contents</title></head><body><nav epub:type=\"toc\"><ol><li><a href=\"chapter.xhtml\">Service review</a></li></ol></nav></body></html>");
            void WriteEntry(string name, string value, CompressionLevel compression = CompressionLevel.Optimal) {
                using var writer = new StreamWriter(archive.CreateEntry(name, compression).Open(), new UTF8Encoding(false));
                writer.Write(value);
            }
        }
        await using (var browser = new HtmlBrowserPdfRenderer(new HtmlBrowserPdfRendererOptions(networkPolicy: HtmlBrowserNetworkPolicy.Offline))) {
            var reference = await browser.CaptureAsync(new HtmlBrowserPdfRequest(HtmlBrowserPdfSource.FromHtml(html),
                new HtmlBrowserPdfOptions(width: "480px", height: "320px", marginTop: "0", marginRight: "0", marginBottom: "0", marginLeft: "0", printBackground: true),
                readiness: new HtmlBrowserPdfReadiness(loadState: HtmlBrowserLoadState.Load, stable: true, stableMilliseconds: 250, timeout: 15000)));
            if (reference.Diagnostics.BlockedRequestCount != 0 || reference.Diagnostics.Warnings.Count != 0)
                throw new InvalidDataException("Reference browser capture reported warnings or blocked resources.");
            File.WriteAllBytes(Path.Combine(output, "browser-reference.pdf"), reference.PdfBytes);
            var emailReference = await browser.CaptureAsync(new HtmlBrowserPdfRequest(HtmlBrowserPdfSource.FromHtml(emailHtml),
                new HtmlBrowserPdfOptions(width: "480px", height: "320px", marginTop: "0", marginRight: "0", marginBottom: "0", marginLeft: "0", printBackground: true),
                readiness: new HtmlBrowserPdfReadiness(loadState: HtmlBrowserLoadState.Load, stable: true, stableMilliseconds: 250, timeout: 15000)));
            if (emailReference.Diagnostics.BlockedRequestCount != 0 || emailReference.Diagnostics.Warnings.Count != 0)
                throw new InvalidDataException("Email reference capture reported warnings or blocked resources.");
            File.WriteAllBytes(Path.Combine(output, "email-reference.pdf"), emailReference.PdfBytes);
        }
        var profile = new OfficeRenderingProfile("conversion-consistency",
            new OfficeFontFaceCollection().Add(family, font).AddFallbackFamily(family), OfficeManagedTextShapingProvider.Instance);
        var rendering = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged, TargetDpi = 96, PageSize = new OfficePageSize(5, 320D / 96),
            DefaultFontFamily = family, Margins = HtmlRenderMargins.All(0), BackgroundColor = OfficeColor.White
        };
        rendering.UseRenderingProfile(profile);
        File.WriteAllBytes(Path.Combine(output, "service-review.pdf"),
            HtmlConversionDocument.Parse(html).ToPdfDocumentResult(new HtmlToPdfOptions(rendering)).ToBytes());
        var visio = VisioDocument.Create(Path.Combine(output, "service-review.vsdx"));
        for (int number = 1; number <= 2; number++) {
            var page = visio.AddPage("Review " + number).Size(5, 320D / 96);
            var shape = page.AddRectangle(2.5, 2, 4, 1, "SERVICE REVIEW " + number);
            shape.TextStyle = new VisioTextStyle { FontFamily = family, Size = 12 };
        }
        visio.Save();
        foreach (string format in new[] { "html", "pdf", "eml", "epub", "vsdx" }) {
            bool external = format is "eml" or "epub";
            cases.Add(new ConsistencyCase {
                Id = "native-" + format, Format = format,
                Source = Path.GetRelativePath(repository, Path.Combine(output, "service-review." + format)).Replace('\\', '/'),
                ReferencePdf = format == "eml" ? "email-reference.pdf" : external ? "browser-reference.pdf" : null,
                ComparePdfPixels = format is not ("eml" or "vsdx"),
                AllowedDiagnostics = format == "vsdx" ? new() { "pdf-projection-visio-semantic-fallback" } : new(),
                Limitations = format == "eml"
                    ? new() { "Email images add message presentation chrome. The external PDF verifies body text and dimensions; its pixels are not compared with the decorated email layout." }
                    : format == "vsdx" ? new() { "Visio PDF currently uses a semantic projection onto paper pages. Diagram image formats are compared with each other; PDF verifies page count, text and paper dimensions." } : new(),
                Evidence = external
                    ? "Authored HTML packaged as MIME or EPUB. Reference PDF is captured independently by Chromium from the same HTML and embedded font during prepare. Source and reference hashes are recorded in the bundle."
                    : "Two-page authored source generated by FixtureCorpus.Other.cs; the PDF fixture uses the native HTML PDF writer, and Visio uses two native diagram pages.",
                Pages = Enumerable.Range(1, format == "eml" ? 1 : 2).Select(number => new PageExpectation {
                    Width = width, Height = height, Text = new() { "SERVICE REVIEW " + number },
                    PdfWidth = format == "vsdx" ? 816 : null, PdfHeight = format == "vsdx" ? 1056 : null
                }).ToList()
            });
        }
    }
}
