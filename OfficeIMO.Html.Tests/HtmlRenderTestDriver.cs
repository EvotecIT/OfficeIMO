using OfficeIMO.Html;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Tests;

/// <summary>
/// Keeps renderer-focused tests concise while the public renderer contract remains source-model-only.
/// Raw test fixtures are parsed through the same native source lifecycle used by consumers.
/// Omitted options retain a continuous layout; image-export defaults are tested through the public image APIs.
/// </summary>
internal static class HtmlRenderTestDriver {
    internal static HtmlRenderDocument Render(string html, HtmlRenderOptions? options = null) =>
        HtmlRenderEngine.Render(HtmlConversionDocument.Parse(html), options ?? ContinuousOptions());

    internal static HtmlRenderDocument Render(HtmlConversionDocument document, HtmlRenderOptions? options = null) =>
        HtmlRenderEngine.Render(document, options ?? ContinuousOptions());

    internal static Task<HtmlRenderDocument> RenderAsync(string html, HtmlRenderOptions? options = null, CancellationToken cancellationToken = default) =>
        HtmlRenderEngine.RenderAsync(HtmlConversionDocument.Parse(html), options ?? ContinuousOptions(), cancellationToken);

    internal static Task<HtmlRenderDocument> RenderAsync(HtmlConversionDocument document, HtmlRenderOptions? options = null, CancellationToken cancellationToken = default) =>
        HtmlRenderEngine.RenderAsync(document, options ?? ContinuousOptions(), cancellationToken);

    internal static HtmlRenderDocument RenderHtml(this string html, HtmlRenderOptions? options = null) => Render(html, options);

    private static HtmlRenderOptions ContinuousOptions() => new() { Mode = HtmlRenderMode.Continuous };
}
