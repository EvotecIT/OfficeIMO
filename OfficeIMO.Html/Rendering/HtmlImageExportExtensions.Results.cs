using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

public static partial class HtmlImageExportExtensions {
    /// <summary>Renders one selected HTML surface to PNG plus dimensions and diagnostics.</summary>
    /// <example><code>OfficeImageExportResult result = html.ToPngResult();</code></example>
    public static OfficeImageExportResult ToPngResult(this string html, HtmlRenderOptions? options = null, int pageIndex = 0) =>
        html.ExportImage(OfficeImageExportFormat.Png, options, pageIndex);

    /// <summary>Renders one selected HTML surface to SVG plus dimensions and diagnostics.</summary>
    /// <example><code>OfficeImageExportResult result = html.ToSvgResult();</code></example>
    public static OfficeImageExportResult ToSvgResult(this string html, HtmlRenderOptions? options = null, int pageIndex = 0) =>
        html.ExportImage(OfficeImageExportFormat.Svg, options, pageIndex);

    /// <summary>Renders every paged HTML surface, or the single continuous surface, to PNG results.</summary>
    public static IReadOnlyList<OfficeImageExportResult> ToPngResults(this string html, HtmlRenderOptions? options = null) =>
        html.ExportImages(OfficeImageExportFormat.Png, options);

    /// <summary>Renders every paged HTML surface, or the single continuous surface, to SVG results.</summary>
    public static IReadOnlyList<OfficeImageExportResult> ToSvgResults(this string html, HtmlRenderOptions? options = null) =>
        html.ExportImages(OfficeImageExportFormat.Svg, options);

    /// <summary>Asynchronously renders one selected HTML surface to PNG plus dimensions and diagnostics.</summary>
    public static Task<OfficeImageExportResult> ToPngResultAsync(this string html, HtmlRenderOptions? options = null, int pageIndex = 0, CancellationToken cancellationToken = default) =>
        html.ExportImageAsync(OfficeImageExportFormat.Png, options, pageIndex, cancellationToken);

    /// <summary>Asynchronously renders one selected HTML surface to SVG plus dimensions and diagnostics.</summary>
    public static Task<OfficeImageExportResult> ToSvgResultAsync(this string html, HtmlRenderOptions? options = null, int pageIndex = 0, CancellationToken cancellationToken = default) =>
        html.ExportImageAsync(OfficeImageExportFormat.Svg, options, pageIndex, cancellationToken);

    /// <summary>Asynchronously renders every paged HTML surface, or the single continuous surface, to PNG results.</summary>
    public static Task<IReadOnlyList<OfficeImageExportResult>> ToPngResultsAsync(this string html, HtmlRenderOptions? options = null, CancellationToken cancellationToken = default) =>
        html.ExportImagesAsync(OfficeImageExportFormat.Png, options, cancellationToken);

    /// <summary>Asynchronously renders every paged HTML surface, or the single continuous surface, to SVG results.</summary>
    public static Task<IReadOnlyList<OfficeImageExportResult>> ToSvgResultsAsync(this string html, HtmlRenderOptions? options = null, CancellationToken cancellationToken = default) =>
        html.ExportImagesAsync(OfficeImageExportFormat.Svg, options, cancellationToken);
}
