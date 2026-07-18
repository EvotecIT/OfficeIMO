using System.Text;

namespace OfficeIMO.Html;

public static partial class HtmlImageExportExtensions {
    /// <summary>Renders a shared HTML conversion document to PNG bytes.</summary>
    public static byte[] ToPng(this HtmlConversionDocument document, HtmlRenderOptions? options = null, int pageIndex = 0) =>
        document.ExportImage(OfficeIMO.Drawing.OfficeImageExportFormat.Png, options, pageIndex).Bytes;

    /// <summary>Renders a shared HTML conversion document to SVG text.</summary>
    public static string ToSvg(this HtmlConversionDocument document, HtmlRenderOptions? options = null, int pageIndex = 0) =>
        Encoding.UTF8.GetString(document.ExportImage(OfficeIMO.Drawing.OfficeImageExportFormat.Svg, options, pageIndex).Bytes);

    /// <summary>Asynchronously resolves resources and renders a shared HTML conversion document to PNG bytes.</summary>
    public static async Task<byte[]> ToPngAsync(this HtmlConversionDocument document, HtmlRenderOptions? options = null, int pageIndex = 0, CancellationToken cancellationToken = default) =>
        (await document.ExportImageAsync(OfficeIMO.Drawing.OfficeImageExportFormat.Png, options, pageIndex, cancellationToken).ConfigureAwait(false)).Bytes;

    /// <summary>Asynchronously resolves resources and renders a shared HTML conversion document to SVG text.</summary>
    public static async Task<string> ToSvgAsync(this HtmlConversionDocument document, HtmlRenderOptions? options = null, int pageIndex = 0, CancellationToken cancellationToken = default) =>
        Encoding.UTF8.GetString((await document.ExportImageAsync(OfficeIMO.Drawing.OfficeImageExportFormat.Svg, options, pageIndex, cancellationToken).ConfigureAwait(false)).Bytes);

    /// <summary>Saves a shared HTML conversion document as a PNG file.</summary>
    public static OfficeIMO.Drawing.OfficeImageExportResult SaveAsPng(this HtmlConversionDocument document, string path, HtmlRenderOptions? options = null, int pageIndex = 0) =>
        new HtmlPageImageExportBuilder(document, options).Page(pageIndex).AsPng().Save(path);

    /// <summary>Saves a shared HTML conversion document as an SVG file.</summary>
    public static OfficeIMO.Drawing.OfficeImageExportResult SaveAsSvg(this HtmlConversionDocument document, string path, HtmlRenderOptions? options = null, int pageIndex = 0) =>
        new HtmlPageImageExportBuilder(document, options).Page(pageIndex).AsSvg().Save(path);

    /// <summary>Writes a shared HTML conversion document as PNG to a caller-owned stream.</summary>
    public static OfficeIMO.Drawing.OfficeImageExportResult SaveAsPng(this HtmlConversionDocument document, Stream stream, HtmlRenderOptions? options = null, int pageIndex = 0) =>
        new HtmlPageImageExportBuilder(document, options).Page(pageIndex).AsPng().Save(stream);

    /// <summary>Writes a shared HTML conversion document as SVG to a caller-owned stream.</summary>
    public static OfficeIMO.Drawing.OfficeImageExportResult SaveAsSvg(this HtmlConversionDocument document, Stream stream, HtmlRenderOptions? options = null, int pageIndex = 0) =>
        new HtmlPageImageExportBuilder(document, options).Page(pageIndex).AsSvg().Save(stream);

    /// <summary>Asynchronously resolves resources and saves a shared HTML conversion document as a PNG file.</summary>
    public static Task<OfficeIMO.Drawing.OfficeImageExportResult> SaveAsPngAsync(this HtmlConversionDocument document, string path, HtmlRenderOptions? options = null, int pageIndex = 0, CancellationToken cancellationToken = default) =>
        new HtmlPageImageExportBuilder(document, options).Page(pageIndex).AsPng().SaveAsync(path, cancellationToken);

    /// <summary>Asynchronously resolves resources and saves a shared HTML conversion document as an SVG file.</summary>
    public static Task<OfficeIMO.Drawing.OfficeImageExportResult> SaveAsSvgAsync(this HtmlConversionDocument document, string path, HtmlRenderOptions? options = null, int pageIndex = 0, CancellationToken cancellationToken = default) =>
        new HtmlPageImageExportBuilder(document, options).Page(pageIndex).AsSvg().SaveAsync(path, cancellationToken);

    /// <summary>Asynchronously resolves resources and writes PNG to a caller-owned stream.</summary>
    public static Task<OfficeIMO.Drawing.OfficeImageExportResult> SaveAsPngAsync(this HtmlConversionDocument document, Stream stream, HtmlRenderOptions? options = null, int pageIndex = 0, CancellationToken cancellationToken = default) =>
        new HtmlPageImageExportBuilder(document, options).Page(pageIndex).AsPng().SaveAsync(stream, cancellationToken);

    /// <summary>Asynchronously resolves resources and writes SVG to a caller-owned stream.</summary>
    public static Task<OfficeIMO.Drawing.OfficeImageExportResult> SaveAsSvgAsync(this HtmlConversionDocument document, Stream stream, HtmlRenderOptions? options = null, int pageIndex = 0, CancellationToken cancellationToken = default) =>
        new HtmlPageImageExportBuilder(document, options).Page(pageIndex).AsSvg().SaveAsync(stream, cancellationToken);
}
