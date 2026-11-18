#if NET10_0_OR_GREATER
using OfficeIMO.Pdf;
using OneRender.Image;
using OneRender.Pdf;
using OneRender.Text;
using OneRenderPageSetup = OneRender.Pdf.PdfPageSetup;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

/// <summary>
/// Adapts OneRender's one-call HTML-to-tagged-PDF API to the shared comparison documents.
/// The font provider is created once and reused, as OneRender recommends: creating one
/// reloads the packaged faces, and each document clears the previous document's
/// @font-face faces before adding its own.
/// </summary>
internal static class OneRenderHtmlGenerator {
    /// <summary>Creates the reusable text service, optionally registering the shared fixture font.</summary>
    internal static TextFontProvider CreateFontProvider(PdfEmbeddedFontFamily? additionalFont = null) {
        TextFontProvider fonts = FontProviders.CreateDefault(null, null);
        if (additionalFont != null) {
            foreach (byte[] face in new[] { additionalFont.Regular, additionalFont.Bold, additionalFont.Italic, additionalFont.BoldItalic }.OfType<byte[]>()) {
                if (!fonts.Catalog.AddFont(face, additionalFont.FamilyName, null, null, null)) {
                    throw new InvalidOperationException($"OneRender did not accept a '{additionalFont.FamilyName}' face.");
                }
            }
        }
        return fonts;
    }

    internal static byte[] Generate(string html, TextFontProvider fonts) {
        // @page rules in the shared documents set A4 and their margins; the A4 setup only
        // supplies the default for pages no rule covers. Output is tagged by default.
        // OneRender never loads resources itself; the host supplies image bytes. The other
        // engines decode the report's data: URL image natively, so this lane decodes it
        // inside the measured call through the documented Resources hook.
        var options = new HtmlPdfOptions {
            Page = OneRenderPageSetup.A4,
            Fonts = fonts,
            Resources = ResolveDataUrl
        };
        return HtmlPdf.ToPdf(html, options, CancellationToken.None);
    }

    private static byte[]? ResolveDataUrl(string address) {
        if (!address.StartsWith("data:", StringComparison.OrdinalIgnoreCase)) {
            return null;
        }

        int comma = address.IndexOf(',', 5);
        if (comma < 0) {
            return null;
        }

        string payload = address[(comma + 1)..];
        bool base64 = address.AsSpan(5, comma - 5).EndsWith(";base64", StringComparison.OrdinalIgnoreCase);
        return base64
            ? Convert.FromBase64String(payload)
            : System.Text.Encoding.UTF8.GetBytes(Uri.UnescapeDataString(payload));
    }
}
#endif
