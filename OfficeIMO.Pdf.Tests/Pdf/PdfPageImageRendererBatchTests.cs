using System.Text;
using System.Threading;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfPageImageRendererBatchTests {
    [Fact]
    public void CapabilityManifest_IsStableAndSharesCodesWithPageDiagnostics() {
        PdfRenderCapabilityManifest manifest = PdfPageImageRenderer.GetCapabilityManifest();
        string json = manifest.ToJson();

        Assert.Equal("officeimo.pdf.render-capabilities.v1", manifest.Schema);
        Assert.Equal(manifest.Entries.OrderBy(static entry => entry.Id, StringComparer.Ordinal).Select(static entry => entry.Id), manifest.Entries.Select(static entry => entry.Id));
        Assert.Equal(manifest.Entries.Count, manifest.Entries.Select(static entry => entry.Id).Distinct(StringComparer.Ordinal).Count());
        Assert.Contains("\"schema\":\"officeimo.pdf.render-capabilities.v1\"", json, StringComparison.Ordinal);
        Assert.Contains(manifest.Entries, static entry => entry.SupportLevel == PdfRenderSupportLevel.Supported);
        Assert.Contains(manifest.Entries, static entry => entry.SupportLevel == PdfRenderSupportLevel.Simplified);
        Assert.Contains(manifest.Entries, static entry => entry.SupportLevel == PdfRenderSupportLevel.Unsupported);

        byte[] pdf = BuildSinglePagePdf("1 M /RelativeColorimetric ri 0.5 i /Tag MP 1 2 FuturePaint");
        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(
            pdf,
            options: new PdfPageRenderOptions { Format = PdfPageRenderFormat.Svg }));

        Assert.True(result.Succeeded, string.Join(" ", result.Diagnostics));
        Assert.Contains(result.CapabilityDiagnostics, static diagnostic => diagnostic.Code == "render.operator.miter-limit-simplified" && diagnostic.Subject == "M");
        Assert.Contains(result.CapabilityDiagnostics, static diagnostic => diagnostic.Code == "render.operator.rendering-intent-simplified" && diagnostic.Subject == "ri");
        Assert.Contains(result.CapabilityDiagnostics, static diagnostic => diagnostic.Code == "render.operator.flatness-simplified" && diagnostic.Subject == "i");
        Assert.Contains(result.CapabilityDiagnostics, static diagnostic => diagnostic.Code == "render.operator.marked-point-simplified" && diagnostic.Subject == "MP");
        Assert.Contains(result.CapabilityDiagnostics, static diagnostic => diagnostic.Code == "render.operator.unsupported" && diagnostic.Subject == "FuturePaint");
        Assert.All(result.CapabilityDiagnostics, diagnostic => Assert.Contains(manifest.Entries, entry => entry.Id == diagnostic.Code));
    }

    [Fact]
    public void RenderPages_RendersCallerOrderedSvgRangeWithPerPageReports() {
        byte[] pdf = BuildTwoPagePdf();

        IReadOnlyList<PdfPageRenderResult> results = PdfDocument.Open(pdf).Read.RenderPages(
            PdfPageSelection.From(2, 1),
            new PdfPageRenderOptions { Format = PdfPageRenderFormat.Svg, Scale = 2D });

        Assert.Equal(new[] { 2, 1 }, results.Select(static result => result.PageNumber));
        Assert.All(results, result => {
            Assert.True(result.Succeeded, string.Join(" ", result.Diagnostics));
            Assert.StartsWith("<svg", Encoding.UTF8.GetString(result.Bytes!), StringComparison.Ordinal);
            Assert.True(result.Width > 1000);
            Assert.True(result.Height > 1500);
            Assert.True(result.Elapsed >= TimeSpan.Zero);
        });
    }

    [Fact]
    public void RenderPages_SupportsThumbnailsCancellationAndExplicitLimits() {
        byte[] pdf = BuildTwoPagePdf();
        IReadOnlyList<PdfPageRenderResult> thumbnails = PdfPageImageRenderer.RenderPages(
            pdf,
            "1-2",
            new PdfPageRenderOptions { ThumbnailMaxDimension = 96 });

        Assert.All(thumbnails, result => {
            Assert.True(result.Succeeded, string.Join(" ", result.Diagnostics));
            Assert.True(Math.Max(result.Width, result.Height) <= 96);
            Assert.Equal(new byte[] { 137, 80, 78, 71 }, result.Bytes!.Take(4));
        });

        PdfReadLimitException pageLimit = Assert.Throws<PdfReadLimitException>(() =>
            PdfPageImageRenderer.RenderPages(pdf, options: new PdfPageRenderOptions { MaxPages = 1 }));
        Assert.Equal(PdfReadLimitKind.RenderPages, pageLimit.Kind);

        PdfPageRenderResult pixelFailure = Assert.Single(PdfPageImageRenderer.RenderPages(
            pdf,
            PdfPageSelection.From(1),
            new PdfPageRenderOptions { MaxPixelsPerPage = 10 }));
        Assert.False(pixelFailure.Succeeded);
        Assert.Contains(pixelFailure.Diagnostics, diagnostic => diagnostic.Contains("render pixel count", StringComparison.OrdinalIgnoreCase));

        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() =>
            PdfPageImageRenderer.RenderPages(pdf, cancellationToken: cancellation.Token));
    }

    [Fact]
    public void RenderPages_EnforcesPerPageAndAggregateEncodedOutputBudgets() {
        byte[] pdf = BuildTwoPagePdf();

        PdfReadLimitException perPage = Assert.Throws<PdfReadLimitException>(() =>
            PdfPageImageRenderer.RenderPages(
                pdf,
                PdfPageSelection.From(1),
                new PdfPageRenderOptions {
                    Format = PdfPageRenderFormat.Svg,
                    MaxOutputBytesPerPage = 1,
                    ContinueOnError = false
                }));
        Assert.Equal(PdfReadLimitKind.RenderBytes, perPage.Kind);

        PdfReadLimitException aggregate = Assert.Throws<PdfReadLimitException>(() =>
            PdfPageImageRenderer.RenderPages(
                pdf,
                options: new PdfPageRenderOptions {
                    Format = PdfPageRenderFormat.Svg,
                    MaxTotalOutputBytes = 1
                }));
        Assert.Equal(PdfReadLimitKind.RenderBytes, aggregate.Kind);
    }

    private static byte[] BuildTwoPagePdf() => PdfDocument.Create()
        .Paragraph(paragraph => paragraph.Text("Batch page one"))
        .PageBreak()
        .Paragraph(paragraph => paragraph.Text("Batch page two"))
        .ToBytes();

    private static byte[] BuildSinglePagePdf(string content) {
        int length = Encoding.ASCII.GetByteCount(content);
        return Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj", "<< /Type /Catalog /Pages 2 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Resources << >> /Contents 4 0 R >>", "endobj",
            "4 0 obj", "<< /Length " + length.ToString(System.Globalization.CultureInfo.InvariantCulture) + " >>", "stream",
            content, "endstream", "endobj",
            "trailer", "<< /Root 1 0 R /Size 5 >>", "%%EOF"
        }));
    }
}
