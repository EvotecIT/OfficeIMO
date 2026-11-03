using System;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using OfficeIMO.Html;
using OfficeIMO.Mhtml;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public async Task MhtmlPdf_AppliesStylesheetWithCidContentLocationAndNoContentId() {
        const string location = "cid:css-site@mhtml.blink";
        const string archive = "MIME-Version: 1.0\r\n"
            + "Content-Type: multipart/related; boundary=archive; type=\"text/html\"\r\n\r\n"
            + "--archive\r\nContent-Type: text/html; charset=utf-8\r\n"
            + "Content-Location: https://snapshot.example.test/page.html\r\n\r\n"
            + "<html><head><link rel='stylesheet' href='" + location + "'></head>"
            + "<body><p class='hidden'>HiddenByArchiveCss</p><p>VisibleArchiveText</p></body></html>\r\n"
            + "--archive\r\nContent-Type: text/css; charset=utf-8\r\n"
            + "Content-Location: " + location + "\r\n\r\n"
            + ".hidden { display: none; }\r\n"
            + "--archive--\r\n";
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(archive));
        MhtmlDocument document = MhtmlDocument.Load(stream);

        Assert.Null(Assert.Single(document.Resources).ContentId);
        var result = await document.ToPdfAsync();
        string text = PdfCore.PdfReadDocument.Open(result.Value.ToBytes()).ExtractText();

        Assert.Contains("VisibleArchiveText", text, StringComparison.Ordinal);
        Assert.DoesNotContain("HiddenByArchiveCss", text, StringComparison.Ordinal);
        Assert.DoesNotContain(result.Report.Diagnostics, warning => warning.Code == HtmlRenderDiagnosticCodes.ResourceUnavailable
            && warning.Location == location);
    }
}
