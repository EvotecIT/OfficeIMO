using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

public static partial class HtmlImageExportExtensions {
    /// <summary>Renders a shared HTML conversion document to JPEG bytes.</summary>
    public static byte[] ToJpeg(this HtmlConversionDocument document, HtmlRenderOptions? options = null, int pageIndex = 0) =>
        document.ExportImage(OfficeImageExportFormat.Jpeg, options, pageIndex).Bytes;

    /// <summary>Renders a shared HTML conversion document to TIFF bytes.</summary>
    public static byte[] ToTiff(this HtmlConversionDocument document, HtmlRenderOptions? options = null, int pageIndex = 0) =>
        document.ExportImage(OfficeImageExportFormat.Tiff, options, pageIndex).Bytes;

    /// <summary>Renders a shared HTML conversion document to lossless WebP bytes.</summary>
    public static byte[] ToWebp(this HtmlConversionDocument document, HtmlRenderOptions? options = null, int pageIndex = 0) =>
        document.ExportImage(OfficeImageExportFormat.Webp, options, pageIndex).Bytes;

    /// <summary>Asynchronously resolves resources and renders a shared HTML conversion document to JPEG bytes.</summary>
    public static async Task<byte[]> ToJpegAsync(this HtmlConversionDocument document, HtmlRenderOptions? options = null, int pageIndex = 0, CancellationToken cancellationToken = default) =>
        (await document.ExportImageAsync(OfficeImageExportFormat.Jpeg, options, pageIndex, cancellationToken).ConfigureAwait(false)).Bytes;

    /// <summary>Asynchronously resolves resources and renders a shared HTML conversion document to TIFF bytes.</summary>
    public static async Task<byte[]> ToTiffAsync(this HtmlConversionDocument document, HtmlRenderOptions? options = null, int pageIndex = 0, CancellationToken cancellationToken = default) =>
        (await document.ExportImageAsync(OfficeImageExportFormat.Tiff, options, pageIndex, cancellationToken).ConfigureAwait(false)).Bytes;

    /// <summary>Asynchronously resolves resources and renders a shared HTML conversion document to lossless WebP bytes.</summary>
    public static async Task<byte[]> ToWebpAsync(this HtmlConversionDocument document, HtmlRenderOptions? options = null, int pageIndex = 0, CancellationToken cancellationToken = default) =>
        (await document.ExportImageAsync(OfficeImageExportFormat.Webp, options, pageIndex, cancellationToken).ConfigureAwait(false)).Bytes;

    /// <summary>Saves a shared HTML conversion document as a JPEG file.</summary>
    public static OfficeImageExportResult SaveAsJpeg(this HtmlConversionDocument document, string path, HtmlRenderOptions? options = null, int pageIndex = 0) =>
        SaveRaster(document, path, OfficeImageExportFormat.Jpeg, options, pageIndex);

    /// <summary>Saves a shared HTML conversion document as a TIFF file.</summary>
    public static OfficeImageExportResult SaveAsTiff(this HtmlConversionDocument document, string path, HtmlRenderOptions? options = null, int pageIndex = 0) =>
        SaveRaster(document, path, OfficeImageExportFormat.Tiff, options, pageIndex);

    /// <summary>Saves a shared HTML conversion document as a lossless WebP file.</summary>
    public static OfficeImageExportResult SaveAsWebp(this HtmlConversionDocument document, string path, HtmlRenderOptions? options = null, int pageIndex = 0) =>
        SaveRaster(document, path, OfficeImageExportFormat.Webp, options, pageIndex);

    /// <summary>Writes a shared HTML conversion document as JPEG to a caller-owned stream.</summary>
    public static OfficeImageExportResult SaveAsJpeg(this HtmlConversionDocument document, Stream stream, HtmlRenderOptions? options = null, int pageIndex = 0) =>
        SaveRaster(document, stream, OfficeImageExportFormat.Jpeg, options, pageIndex);

    /// <summary>Writes a shared HTML conversion document as TIFF to a caller-owned stream.</summary>
    public static OfficeImageExportResult SaveAsTiff(this HtmlConversionDocument document, Stream stream, HtmlRenderOptions? options = null, int pageIndex = 0) =>
        SaveRaster(document, stream, OfficeImageExportFormat.Tiff, options, pageIndex);

    /// <summary>Writes a shared HTML conversion document as lossless WebP to a caller-owned stream.</summary>
    public static OfficeImageExportResult SaveAsWebp(this HtmlConversionDocument document, Stream stream, HtmlRenderOptions? options = null, int pageIndex = 0) =>
        SaveRaster(document, stream, OfficeImageExportFormat.Webp, options, pageIndex);

    /// <summary>Asynchronously resolves resources and saves a shared HTML conversion document as a JPEG file.</summary>
    public static Task<OfficeImageExportResult> SaveAsJpegAsync(this HtmlConversionDocument document, string path, HtmlRenderOptions? options = null, int pageIndex = 0, CancellationToken cancellationToken = default) =>
        SaveRasterAsync(document, path, OfficeImageExportFormat.Jpeg, options, pageIndex, cancellationToken);

    /// <summary>Asynchronously resolves resources and saves a shared HTML conversion document as a TIFF file.</summary>
    public static Task<OfficeImageExportResult> SaveAsTiffAsync(this HtmlConversionDocument document, string path, HtmlRenderOptions? options = null, int pageIndex = 0, CancellationToken cancellationToken = default) =>
        SaveRasterAsync(document, path, OfficeImageExportFormat.Tiff, options, pageIndex, cancellationToken);

    /// <summary>Asynchronously resolves resources and saves a shared HTML conversion document as a lossless WebP file.</summary>
    public static Task<OfficeImageExportResult> SaveAsWebpAsync(this HtmlConversionDocument document, string path, HtmlRenderOptions? options = null, int pageIndex = 0, CancellationToken cancellationToken = default) =>
        SaveRasterAsync(document, path, OfficeImageExportFormat.Webp, options, pageIndex, cancellationToken);

    /// <summary>Asynchronously resolves resources and writes JPEG to a caller-owned stream.</summary>
    public static Task<OfficeImageExportResult> SaveAsJpegAsync(this HtmlConversionDocument document, Stream stream, HtmlRenderOptions? options = null, int pageIndex = 0, CancellationToken cancellationToken = default) =>
        SaveRasterAsync(document, stream, OfficeImageExportFormat.Jpeg, options, pageIndex, cancellationToken);

    /// <summary>Asynchronously resolves resources and writes TIFF to a caller-owned stream.</summary>
    public static Task<OfficeImageExportResult> SaveAsTiffAsync(this HtmlConversionDocument document, Stream stream, HtmlRenderOptions? options = null, int pageIndex = 0, CancellationToken cancellationToken = default) =>
        SaveRasterAsync(document, stream, OfficeImageExportFormat.Tiff, options, pageIndex, cancellationToken);

    /// <summary>Asynchronously resolves resources and writes lossless WebP to a caller-owned stream.</summary>
    public static Task<OfficeImageExportResult> SaveAsWebpAsync(this HtmlConversionDocument document, Stream stream, HtmlRenderOptions? options = null, int pageIndex = 0, CancellationToken cancellationToken = default) =>
        SaveRasterAsync(document, stream, OfficeImageExportFormat.Webp, options, pageIndex, cancellationToken);

    private static OfficeImageExportResult SaveRaster(
        HtmlConversionDocument document,
        string path,
        OfficeImageExportFormat format,
        HtmlRenderOptions? options,
        int pageIndex) {
        return new HtmlPageImageExportBuilder(document, options)
            .Page(pageIndex)
            .As(format)
            .Save(path);
    }

    private static OfficeImageExportResult SaveRaster(
        HtmlConversionDocument document,
        Stream stream,
        OfficeImageExportFormat format,
        HtmlRenderOptions? options,
        int pageIndex) {
        return new HtmlPageImageExportBuilder(document, options)
            .Page(pageIndex)
            .As(format)
            .Save(stream);
    }

    private static async Task<OfficeImageExportResult> SaveRasterAsync(
        HtmlConversionDocument document,
        string path,
        OfficeImageExportFormat format,
        HtmlRenderOptions? options,
        int pageIndex,
        CancellationToken cancellationToken) {
        return await new HtmlPageImageExportBuilder(document, options)
            .Page(pageIndex)
            .As(format)
            .SaveAsync(path, cancellationToken)
            .ConfigureAwait(false);
    }

    private static async Task<OfficeImageExportResult> SaveRasterAsync(
        HtmlConversionDocument document,
        Stream stream,
        OfficeImageExportFormat format,
        HtmlRenderOptions? options,
        int pageIndex,
        CancellationToken cancellationToken) {
        return await new HtmlPageImageExportBuilder(document, options)
            .Page(pageIndex)
            .As(format)
            .SaveAsync(stream, cancellationToken)
            .ConfigureAwait(false);
    }
}
