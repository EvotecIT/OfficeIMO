using OfficeIMO.Pdf;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

/// <summary>Adapts the published comparison package without changing the shared HTML or its fonts.</summary>
internal static class PdfPeckerGenerator {
    internal static byte[] Generate(string html, PdfEmbeddedFontFamily? additionalFont = null) {
        var options = new global::PdfPecker.PdfCreateOptions {
            PageSize = global::PdfPecker.PageSize.A4,
            Accessibility = global::PdfPecker.PdfAccessibility.Tagged
        };
        if (additionalFont != null) {
            foreach (byte[] face in new[] { additionalFont.Regular, additionalFont.Bold, additionalFont.Italic, additionalFont.BoldItalic }.OfType<byte[]>()) {
                options.Fonts.Add(face, additionalFont.FamilyName);
            }
        }
        return global::PdfPecker.HtmlToPdf.Convert(html, options, CancellationToken.None);
    }

    /// <summary>Observes cancellation after the public converter has begun resolving an input resource.</summary>
    internal static async Task<HtmlPdfCancellationEvidence> ProbeCancellationAsync(string html) {
        var started = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
        using var cancellation = new CancellationTokenSource();
        var release = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
        var options = new global::PdfPecker.PdfCreateOptions {
            ResourceResolver = _ => {
                started.TrySetResult(true);
                // The resolver is synchronous and has no token. Release it after cancellation;
                // the converter itself must then observe the caller's cancelled token.
                release.Task.Wait(TimeSpan.FromSeconds(10));
                return null;
            }
        };
        const string image = "<img src=\"https://officeimo.invalid/cancellation-probe.png\" alt=\"cancellation probe\">";
        int bodyEnd = html.LastIndexOf("</body>", StringComparison.OrdinalIgnoreCase);
        string probeHtml = bodyEnd >= 0 ? html.Insert(bodyEnd, image) : html + image;
        Task<byte[]> conversion = Task.Run(() => global::PdfPecker.HtmlToPdf.Convert(probeHtml, options, cancellation.Token));
        try {
            await started.Task.WaitAsync(TimeSpan.FromSeconds(10)).ConfigureAwait(false);
            var stopwatch = System.Diagnostics.Stopwatch.StartNew();
            cancellation.Cancel();
            release.TrySetResult(true);
            try {
                _ = await conversion.WaitAsync(TimeSpan.FromSeconds(3)).ConfigureAwait(false);
                return new(true, "Failed", "An in-flight PdfPecker HTML-to-PDF request completed instead of cancelling.");
            } catch (OperationCanceledException) when (cancellation.IsCancellationRequested) {
                return new(true, "Passed", $"An in-flight PdfPecker HTML-to-PDF request cancelled in {stopwatch.Elapsed.TotalMilliseconds:0} ms through HtmlToPdf.Convert.");
            }
        } catch (Exception exception) {
            return new(true, "Failed", exception.GetType().Name + ": " + exception.Message);
        } finally {
            cancellation.Cancel();
            release.TrySetResult(true);
            // Bound cleanup too: a broken cancellation implementation must fail the CLI
            // rather than leave the coordinator waiting indefinitely.
            try { _ = await conversion.WaitAsync(TimeSpan.FromSeconds(10)).ConfigureAwait(false); } catch (Exception) { }
        }
    }
}
