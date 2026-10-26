using OfficeIMO.Html;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Tests;

/// <summary>
/// Keeps renderer-focused tests concise while the public renderer contract remains source-model-only.
/// Raw test fixtures are parsed through the same native source lifecycle used by consumers.
/// Fixtures use explicit portable font policy, so host installations cannot change their geometry.
/// Installed-font integration and image-export defaults use the public engine APIs directly.
/// </summary>
internal static class HtmlRenderTestDriver {
    internal static HtmlRenderDocument Render(string html, HtmlRenderOptions? options = null) =>
        HtmlRenderEngine.Render(HtmlConversionDocument.Parse(html), PortableOptions(options));

    internal static HtmlRenderDocument Render(HtmlConversionDocument document, HtmlRenderOptions? options = null) =>
        HtmlRenderEngine.Render(document, PortableOptions(options));

    internal static Task<HtmlRenderDocument> RenderAsync(string html, HtmlRenderOptions? options = null, CancellationToken cancellationToken = default) =>
        HtmlRenderEngine.RenderAsync(HtmlConversionDocument.Parse(html), PortableOptions(options), cancellationToken);

    internal static Task<HtmlRenderDocument> RenderAsync(HtmlConversionDocument document, HtmlRenderOptions? options = null, CancellationToken cancellationToken = default) =>
        HtmlRenderEngine.RenderAsync(document, PortableOptions(options), cancellationToken);

    internal static HtmlRenderDocument RenderHtml(this string html, HtmlRenderOptions? options = null) => Render(html, options);

    private static HtmlRenderOptions PortableOptions(HtmlRenderOptions? options) {
        HtmlRenderOptions result = options?.Clone() ?? new() { Mode = HtmlRenderMode.Continuous };
        result.AllowSystemFontFallback = false;
        return result;
    }

    // Geometric fixtures may deliberately omit font bytes. Their new typed estimate
    // remains visible; every other diagnostic still fails the original clean-path check.
    internal static void AssertOnlyUnselectedFontEstimates(HtmlRenderDocument rendered) =>
        Xunit.Assert.All(rendered.Diagnostics, diagnostic => {
            Xunit.Assert.Equal(HtmlRenderDiagnosticCodes.FontMetricsEstimated, diagnostic.Code);
            Xunit.Assert.Equal(OfficeConversionLossKind.Approximation, diagnostic.LossKind);
        });
}
