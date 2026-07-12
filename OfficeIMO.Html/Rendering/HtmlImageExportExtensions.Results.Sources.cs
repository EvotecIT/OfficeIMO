using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

public static partial class HtmlImageExportExtensions {
    /// <summary>Renders one selected surface from a shared HTML conversion document to PNG plus dimensions and diagnostics.</summary>
    public static OfficeImageExportResult ToPngResult(this HtmlConversionDocument document, HtmlRenderOptions? options = null, int pageIndex = 0) =>
        GetDocument(document, options).ExportImage(OfficeImageExportFormat.Png, options, pageIndex);

    /// <summary>Renders one selected surface from a shared HTML conversion document to SVG plus dimensions and diagnostics.</summary>
    public static OfficeImageExportResult ToSvgResult(this HtmlConversionDocument document, HtmlRenderOptions? options = null, int pageIndex = 0) =>
        GetDocument(document, options).ExportImage(OfficeImageExportFormat.Svg, options, pageIndex);

    /// <summary>Renders all surfaces from a shared HTML conversion document to PNG results.</summary>
    public static IReadOnlyList<OfficeImageExportResult> ToPngResults(this HtmlConversionDocument document, HtmlRenderOptions? options = null) =>
        GetDocument(document, options).ExportImages(OfficeImageExportFormat.Png, options);

    /// <summary>Renders all surfaces from a shared HTML conversion document to SVG results.</summary>
    public static IReadOnlyList<OfficeImageExportResult> ToSvgResults(this HtmlConversionDocument document, HtmlRenderOptions? options = null) =>
        GetDocument(document, options).ExportImages(OfficeImageExportFormat.Svg, options);

    /// <summary>Asynchronously renders one selected surface from a shared HTML conversion document to PNG plus dimensions and diagnostics.</summary>
    public static Task<OfficeImageExportResult> ToPngResultAsync(this HtmlConversionDocument document, HtmlRenderOptions? options = null, int pageIndex = 0, CancellationToken cancellationToken = default) =>
        GetDocument(document, options).ExportImageAsync(OfficeImageExportFormat.Png, options, pageIndex, cancellationToken);

    /// <summary>Asynchronously renders one selected surface from a shared HTML conversion document to SVG plus dimensions and diagnostics.</summary>
    public static Task<OfficeImageExportResult> ToSvgResultAsync(this HtmlConversionDocument document, HtmlRenderOptions? options = null, int pageIndex = 0, CancellationToken cancellationToken = default) =>
        GetDocument(document, options).ExportImageAsync(OfficeImageExportFormat.Svg, options, pageIndex, cancellationToken);

    /// <summary>Asynchronously renders all surfaces from a shared HTML conversion document to PNG results.</summary>
    public static Task<IReadOnlyList<OfficeImageExportResult>> ToPngResultsAsync(this HtmlConversionDocument document, HtmlRenderOptions? options = null, CancellationToken cancellationToken = default) =>
        GetDocument(document, options).ExportImagesAsync(OfficeImageExportFormat.Png, options, cancellationToken);

    /// <summary>Asynchronously renders all surfaces from a shared HTML conversion document to SVG results.</summary>
    public static Task<IReadOnlyList<OfficeImageExportResult>> ToSvgResultsAsync(this HtmlConversionDocument document, HtmlRenderOptions? options = null, CancellationToken cancellationToken = default) =>
        GetDocument(document, options).ExportImagesAsync(OfficeImageExportFormat.Svg, options, cancellationToken);

    /// <summary>Reads UTF-8 HTML and renders one selected surface to PNG plus dimensions and diagnostics.</summary>
    public static OfficeImageExportResult ToPngResult(this Stream htmlStream, HtmlRenderOptions? options = null, int pageIndex = 0) =>
        ReadHtml(htmlStream).ToPngResult(options, pageIndex);

    /// <summary>Reads UTF-8 HTML and renders one selected surface to SVG plus dimensions and diagnostics.</summary>
    public static OfficeImageExportResult ToSvgResult(this Stream htmlStream, HtmlRenderOptions? options = null, int pageIndex = 0) =>
        ReadHtml(htmlStream).ToSvgResult(options, pageIndex);

    /// <summary>Reads UTF-8 HTML and renders all surfaces to PNG results.</summary>
    public static IReadOnlyList<OfficeImageExportResult> ToPngResults(this Stream htmlStream, HtmlRenderOptions? options = null) =>
        ReadHtml(htmlStream).ToPngResults(options);

    /// <summary>Reads UTF-8 HTML and renders all surfaces to SVG results.</summary>
    public static IReadOnlyList<OfficeImageExportResult> ToSvgResults(this Stream htmlStream, HtmlRenderOptions? options = null) =>
        ReadHtml(htmlStream).ToSvgResults(options);

    /// <summary>Asynchronously reads UTF-8 HTML and renders one selected surface to PNG plus dimensions and diagnostics.</summary>
    public static async Task<OfficeImageExportResult> ToPngResultAsync(this Stream htmlStream, HtmlRenderOptions? options = null, int pageIndex = 0, CancellationToken cancellationToken = default) =>
        await (await ReadHtmlAsync(htmlStream, cancellationToken).ConfigureAwait(false)).ToPngResultAsync(options, pageIndex, cancellationToken).ConfigureAwait(false);

    /// <summary>Asynchronously reads UTF-8 HTML and renders one selected surface to SVG plus dimensions and diagnostics.</summary>
    public static async Task<OfficeImageExportResult> ToSvgResultAsync(this Stream htmlStream, HtmlRenderOptions? options = null, int pageIndex = 0, CancellationToken cancellationToken = default) =>
        await (await ReadHtmlAsync(htmlStream, cancellationToken).ConfigureAwait(false)).ToSvgResultAsync(options, pageIndex, cancellationToken).ConfigureAwait(false);

    /// <summary>Asynchronously reads UTF-8 HTML and renders all surfaces to PNG results.</summary>
    public static async Task<IReadOnlyList<OfficeImageExportResult>> ToPngResultsAsync(this Stream htmlStream, HtmlRenderOptions? options = null, CancellationToken cancellationToken = default) =>
        await (await ReadHtmlAsync(htmlStream, cancellationToken).ConfigureAwait(false)).ToPngResultsAsync(options, cancellationToken).ConfigureAwait(false);

    /// <summary>Asynchronously reads UTF-8 HTML and renders all surfaces to SVG results.</summary>
    public static async Task<IReadOnlyList<OfficeImageExportResult>> ToSvgResultsAsync(this Stream htmlStream, HtmlRenderOptions? options = null, CancellationToken cancellationToken = default) =>
        await (await ReadHtmlAsync(htmlStream, cancellationToken).ConfigureAwait(false)).ToSvgResultsAsync(options, cancellationToken).ConfigureAwait(false);
}
