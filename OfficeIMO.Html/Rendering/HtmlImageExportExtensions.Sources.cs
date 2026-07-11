using System.Text;

namespace OfficeIMO.Html;

public static partial class HtmlImageExportExtensions {
    /// <summary>Renders a shared HTML conversion document to PNG bytes.</summary>
    public static byte[] ToPng(this HtmlConversionDocument document, HtmlRenderOptions? options = null, int pageIndex = 0) =>
        GetHtml(document).ToPng(options, pageIndex);

    /// <summary>Renders a shared HTML conversion document to SVG text.</summary>
    public static string ToSvg(this HtmlConversionDocument document, HtmlRenderOptions? options = null, int pageIndex = 0) =>
        GetHtml(document).ToSvg(options, pageIndex);

    /// <summary>Reads UTF-8 HTML from a stream and renders it to PNG bytes.</summary>
    public static byte[] ToPng(this Stream htmlStream, HtmlRenderOptions? options = null, int pageIndex = 0) =>
        ReadHtml(htmlStream).ToPng(options, pageIndex);

    /// <summary>Reads UTF-8 HTML from a stream and renders it to SVG text.</summary>
    public static string ToSvg(this Stream htmlStream, HtmlRenderOptions? options = null, int pageIndex = 0) =>
        ReadHtml(htmlStream).ToSvg(options, pageIndex);

    /// <summary>Asynchronously renders a shared HTML conversion document to PNG bytes.</summary>
    public static Task<byte[]> ToPngAsync(this HtmlConversionDocument document, HtmlRenderOptions? options = null, int pageIndex = 0, CancellationToken cancellationToken = default) =>
        GetHtml(document).ToPngAsync(options, pageIndex, cancellationToken);

    /// <summary>Asynchronously renders a shared HTML conversion document to SVG text.</summary>
    public static Task<string> ToSvgAsync(this HtmlConversionDocument document, HtmlRenderOptions? options = null, int pageIndex = 0, CancellationToken cancellationToken = default) =>
        GetHtml(document).ToSvgAsync(options, pageIndex, cancellationToken);

    /// <summary>Asynchronously reads UTF-8 HTML from a stream and renders it to PNG bytes.</summary>
    public static async Task<byte[]> ToPngAsync(this Stream htmlStream, HtmlRenderOptions? options = null, int pageIndex = 0, CancellationToken cancellationToken = default) =>
        await (await ReadHtmlAsync(htmlStream, cancellationToken).ConfigureAwait(false)).ToPngAsync(options, pageIndex, cancellationToken).ConfigureAwait(false);

    /// <summary>Asynchronously reads UTF-8 HTML from a stream and renders it to SVG text.</summary>
    public static async Task<string> ToSvgAsync(this Stream htmlStream, HtmlRenderOptions? options = null, int pageIndex = 0, CancellationToken cancellationToken = default) =>
        await (await ReadHtmlAsync(htmlStream, cancellationToken).ConfigureAwait(false)).ToSvgAsync(options, pageIndex, cancellationToken).ConfigureAwait(false);

    /// <summary>Saves a shared HTML conversion document as a PNG file.</summary>
    public static void SaveAsPng(this HtmlConversionDocument document, string path, HtmlRenderOptions? options = null, int pageIndex = 0) =>
        GetHtml(document).SaveAsPng(path, options, pageIndex);

    /// <summary>Saves a shared HTML conversion document as an SVG file.</summary>
    public static void SaveAsSvg(this HtmlConversionDocument document, string path, HtmlRenderOptions? options = null, int pageIndex = 0) =>
        GetHtml(document).SaveAsSvg(path, options, pageIndex);

    /// <summary>Writes a shared HTML conversion document as PNG to a stream.</summary>
    public static void SaveAsPng(this HtmlConversionDocument document, Stream stream, HtmlRenderOptions? options = null, int pageIndex = 0) =>
        GetHtml(document).SaveAsPng(stream, options, pageIndex);

    /// <summary>Writes a shared HTML conversion document as SVG to a stream.</summary>
    public static void SaveAsSvg(this HtmlConversionDocument document, Stream stream, HtmlRenderOptions? options = null, int pageIndex = 0) =>
        GetHtml(document).SaveAsSvg(stream, options, pageIndex);

    /// <summary>Reads UTF-8 HTML from a stream and saves it as a PNG file.</summary>
    public static void SaveAsPng(this Stream htmlStream, string path, HtmlRenderOptions? options = null, int pageIndex = 0) =>
        ReadHtml(htmlStream).SaveAsPng(path, options, pageIndex);

    /// <summary>Reads UTF-8 HTML from a stream and saves it as an SVG file.</summary>
    public static void SaveAsSvg(this Stream htmlStream, string path, HtmlRenderOptions? options = null, int pageIndex = 0) =>
        ReadHtml(htmlStream).SaveAsSvg(path, options, pageIndex);

    /// <summary>Reads UTF-8 HTML and writes PNG to another stream.</summary>
    public static void SaveAsPng(this Stream htmlStream, Stream pngStream, HtmlRenderOptions? options = null, int pageIndex = 0) =>
        ReadHtml(htmlStream).SaveAsPng(pngStream, options, pageIndex);

    /// <summary>Reads UTF-8 HTML and writes SVG to another stream.</summary>
    public static void SaveAsSvg(this Stream htmlStream, Stream svgStream, HtmlRenderOptions? options = null, int pageIndex = 0) =>
        ReadHtml(htmlStream).SaveAsSvg(svgStream, options, pageIndex);

    /// <summary>Asynchronously saves a shared HTML conversion document as a PNG file.</summary>
    public static Task SaveAsPngAsync(this HtmlConversionDocument document, string path, HtmlRenderOptions? options = null, int pageIndex = 0, CancellationToken cancellationToken = default) =>
        GetHtml(document).SaveAsPngAsync(path, options, pageIndex, cancellationToken);

    /// <summary>Asynchronously saves a shared HTML conversion document as an SVG file.</summary>
    public static Task SaveAsSvgAsync(this HtmlConversionDocument document, string path, HtmlRenderOptions? options = null, int pageIndex = 0, CancellationToken cancellationToken = default) =>
        GetHtml(document).SaveAsSvgAsync(path, options, pageIndex, cancellationToken);

    /// <summary>Asynchronously writes a shared HTML conversion document as PNG to a stream.</summary>
    public static Task SaveAsPngAsync(this HtmlConversionDocument document, Stream stream, HtmlRenderOptions? options = null, int pageIndex = 0, CancellationToken cancellationToken = default) =>
        GetHtml(document).SaveAsPngAsync(stream, options, pageIndex, cancellationToken);

    /// <summary>Asynchronously writes a shared HTML conversion document as SVG to a stream.</summary>
    public static Task SaveAsSvgAsync(this HtmlConversionDocument document, Stream stream, HtmlRenderOptions? options = null, int pageIndex = 0, CancellationToken cancellationToken = default) =>
        GetHtml(document).SaveAsSvgAsync(stream, options, pageIndex, cancellationToken);

    /// <summary>Asynchronously reads UTF-8 HTML and saves it as a PNG file.</summary>
    public static async Task SaveAsPngAsync(this Stream htmlStream, string path, HtmlRenderOptions? options = null, int pageIndex = 0, CancellationToken cancellationToken = default) =>
        await (await ReadHtmlAsync(htmlStream, cancellationToken).ConfigureAwait(false)).SaveAsPngAsync(path, options, pageIndex, cancellationToken).ConfigureAwait(false);

    /// <summary>Asynchronously reads UTF-8 HTML and saves it as an SVG file.</summary>
    public static async Task SaveAsSvgAsync(this Stream htmlStream, string path, HtmlRenderOptions? options = null, int pageIndex = 0, CancellationToken cancellationToken = default) =>
        await (await ReadHtmlAsync(htmlStream, cancellationToken).ConfigureAwait(false)).SaveAsSvgAsync(path, options, pageIndex, cancellationToken).ConfigureAwait(false);

    /// <summary>Asynchronously reads UTF-8 HTML and writes PNG to another stream.</summary>
    public static async Task SaveAsPngAsync(this Stream htmlStream, Stream pngStream, HtmlRenderOptions? options = null, int pageIndex = 0, CancellationToken cancellationToken = default) =>
        await (await ReadHtmlAsync(htmlStream, cancellationToken).ConfigureAwait(false)).SaveAsPngAsync(pngStream, options, pageIndex, cancellationToken).ConfigureAwait(false);

    /// <summary>Asynchronously reads UTF-8 HTML and writes SVG to another stream.</summary>
    public static async Task SaveAsSvgAsync(this Stream htmlStream, Stream svgStream, HtmlRenderOptions? options = null, int pageIndex = 0, CancellationToken cancellationToken = default) =>
        await (await ReadHtmlAsync(htmlStream, cancellationToken).ConfigureAwait(false)).SaveAsSvgAsync(svgStream, options, pageIndex, cancellationToken).ConfigureAwait(false);

    private static string GetHtml(HtmlConversionDocument document) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        return document.HtmlForConversion;
    }

    private static string ReadHtml(Stream htmlStream) {
        if (htmlStream == null) throw new ArgumentNullException(nameof(htmlStream));
        using var reader = new StreamReader(htmlStream, Encoding.UTF8, true, 4096, true);
        return reader.ReadToEnd();
    }

    private static async Task<string> ReadHtmlAsync(Stream htmlStream, CancellationToken cancellationToken) {
        if (htmlStream == null) throw new ArgumentNullException(nameof(htmlStream));
        using var reader = new StreamReader(htmlStream, Encoding.UTF8, true, 4096, true);
#if NET8_0_OR_GREATER
        return await reader.ReadToEndAsync(cancellationToken).ConfigureAwait(false);
#else
        string html = await reader.ReadToEndAsync().ConfigureAwait(false);
        cancellationToken.ThrowIfCancellationRequested();
        return html;
#endif
    }
}
