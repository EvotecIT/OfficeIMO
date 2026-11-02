using System;
using System.IO;
using System.Threading.Tasks;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlPdfResourceAdmissionTests {
    [Theory]
    [InlineData(false, "disable")]
    [InlineData(true, "disable")]
    [InlineData(false, "resolver")]
    [InlineData(true, "resolver")]
    [InlineData(false, "unchanged")]
    [InlineData(true, "unchanged")]
    public async Task RetainedRequestsRecheckLocalFileAdmission(bool asynchronous, string policy) {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-html-retained-admission-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            File.WriteAllBytes(Path.Combine(root, "pixel.png"), Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII="));
            var source = HtmlConversionDocument.Parse("<html><body><img src=\"pixel.png\" alt=\"pixel\"></body></html>",
                new HtmlConversionDocumentOptions {
                    BaseUri = new Uri(Path.Combine(root, "source.html")),
                    ResourceUrlPolicy = new HtmlUrlPolicy { DisallowFileUrls = false }
                });
            var options = new HtmlToPdfOptions();
            options.ResourcePolicy.AllowLocalFileAccess = true;
            var initial = source.RenderToPdfResult(HtmlRenderRequest.Create(HtmlRenderIntentProfile.PrintPaged, HtmlRenderEncoder.Pdf, options));
            using var initialModel = initial.Output.Value;
            Assert.Equal(1, CountImages(initial.Output.ToBytes()));
            var retained = Assert.IsType<HtmlToPdfOptions>(initial.RenderResult.Request.Options);
            if (policy == "resolver") retained.ResourceResolver = (request, token) => Task.FromResult<HtmlResolvedResource?>(null);
            else if (policy == "disable") retained.ResourcePolicy.AllowLocalFileAccess = false;
            HtmlRenderRequest next = initial.RenderResult.Request.WithOptions(retained);
            var result = asynchronous ? await source.RenderToPdfResultAsync(next) : source.RenderToPdfResult(next);
            using var resultModel = result.Output.Value;
            Assert.Equal(policy == "unchanged" ? 1 : 0, CountImages(result.Output.ToBytes()));
        } finally { Directory.Delete(root, true); }
    }

    private static int CountImages(byte[] bytes) {
        using var input = new MemoryStream(bytes);
        using var document = PdfDocument.Load(input);
        return document.Images.Extract().Count;
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public async Task AConfiguredResolverIsNeverBypassedByLocalFileLoading(bool asynchronous, bool admit) {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-html-admission-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            byte[] pixel = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=");
            File.WriteAllBytes(Path.Combine(root, "pixel.png"), pixel);
            var document = HtmlConversionDocument.Parse("<html><body><img src=\"pixel.png\" alt=\"pixel\"></body></html>",
                new HtmlConversionDocumentOptions {
                    BaseUri = new Uri(Path.Combine(root, "source.html")),
                    ResourceUrlPolicy = new HtmlUrlPolicy { DisallowFileUrls = false }
                });
            int requests = 0;
            var options = new HtmlToPdfOptions {
                ResourceResolver = (request, token) => {
                    requests++;
                    return Task.FromResult<HtmlResolvedResource?>(admit ? new HtmlResolvedResource(pixel, "image/png") : null);
                }
            };
            options.ResourcePolicy.AllowLocalFileAccess = true;
            using var output = new MemoryStream();
            var report = asynchronous
                ? await document.SaveAsPdfAsync(output, options)
                : document.SaveAsPdf(output, options);
            // Application resolvers are an asynchronous boundary. Synchronous rendering
            // omits the resource rather than loading a file behind the host's policy.
            if (asynchronous) Assert.True(requests > 0);
            else Assert.Equal(0, requests);
            output.Position = 0;
            using var pdf = PdfDocument.Load(output);
            Assert.Equal(asynchronous && admit ? 1 : 0, pdf.Images.Extract().Count);
            if (asynchronous && !admit) Assert.Contains(report.Diagnostics, item => item.Code == "HtmlRenderResourceUnavailable");
        } finally { Directory.Delete(root, true); }
    }
}
