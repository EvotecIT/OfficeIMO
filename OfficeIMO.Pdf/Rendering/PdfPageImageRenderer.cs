using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

/// <summary>
/// Managed PDF page image renderer for the OfficeIMO-generated visual subset.
/// </summary>
public static class PdfPageImageRenderer {
    /// <summary>
    /// Projects a one-based PDF page into the shared OfficeIMO drawing scene.
    /// </summary>
    public static OfficeDrawing RenderPage(byte[] pdf, int pageNumber = 1) {
        Guard.NotNull(pdf, nameof(pdf));
        return RenderPage(PdfReadDocument.Load(pdf), pageNumber);
    }

    /// <summary>
    /// Projects a one-based PDF page from the current stream position into the shared OfficeIMO drawing scene.
    /// </summary>
    public static OfficeDrawing RenderPage(Stream stream, int pageNumber = 1) {
        Guard.NotNull(stream, nameof(stream));
        if (!stream.CanRead) {
            throw new ArgumentException("Stream must be readable.", nameof(stream));
        }

        using var buffer = new MemoryStream();
        stream.CopyTo(buffer);
        return RenderPage(buffer.ToArray(), pageNumber);
    }

    /// <summary>
    /// Projects a one-based PDF page from a file into the shared OfficeIMO drawing scene.
    /// </summary>
    public static OfficeDrawing RenderPage(string path, int pageNumber = 1) {
        Guard.NotNullOrWhiteSpace(path, nameof(path));
        return RenderPage(File.ReadAllBytes(path), pageNumber);
    }

    /// <summary>
    /// Projects a one-based page from an already loaded PDF document into the shared OfficeIMO drawing scene.
    /// </summary>
    public static OfficeDrawing RenderPage(PdfReadDocument document, int pageNumber = 1) {
        Guard.NotNull(document, nameof(document));
        ValidatePageNumber(document, pageNumber);
        return document.Pages[pageNumber - 1].ToDrawing();
    }

    /// <summary>
    /// Renders a one-based PDF page to dependency-free PNG bytes through the shared OfficeIMO drawing rasterizer.
    /// </summary>
    public static byte[] RenderPageAsPng(byte[] pdf, int pageNumber = 1, double scale = 1D, OfficeColor? background = null) =>
        OfficeDrawingRasterRenderer.ToPng(RenderPage(pdf, pageNumber), scale, background ?? OfficeColor.White);

    /// <summary>
    /// Renders a one-based PDF page to dependency-free PNG bytes from the current stream position.
    /// </summary>
    public static byte[] RenderPageAsPng(Stream stream, int pageNumber = 1, double scale = 1D, OfficeColor? background = null) =>
        OfficeDrawingRasterRenderer.ToPng(RenderPage(stream, pageNumber), scale, background ?? OfficeColor.White);

    /// <summary>
    /// Renders a one-based PDF page to dependency-free PNG bytes from a file path.
    /// </summary>
    public static byte[] RenderPageAsPng(string path, int pageNumber = 1, double scale = 1D, OfficeColor? background = null) =>
        OfficeDrawingRasterRenderer.ToPng(RenderPage(path, pageNumber), scale, background ?? OfficeColor.White);

    /// <summary>
    /// Renders a one-based PDF page to UTF-8 SVG bytes through the shared OfficeIMO drawing SVG exporter.
    /// </summary>
    public static byte[] RenderPageAsSvg(byte[] pdf, int pageNumber = 1, double scale = 1D) =>
        OfficeDrawingSvgExporter.ToSvgBytes(RenderPage(pdf, pageNumber), scale);

    /// <summary>
    /// Renders a one-based PDF page to UTF-8 SVG bytes from the current stream position.
    /// </summary>
    public static byte[] RenderPageAsSvg(Stream stream, int pageNumber = 1, double scale = 1D) =>
        OfficeDrawingSvgExporter.ToSvgBytes(RenderPage(stream, pageNumber), scale);

    /// <summary>
    /// Renders a one-based PDF page to UTF-8 SVG bytes from a file path.
    /// </summary>
    public static byte[] RenderPageAsSvg(string path, int pageNumber = 1, double scale = 1D) =>
        OfficeDrawingSvgExporter.ToSvgBytes(RenderPage(path, pageNumber), scale);

    private static void ValidatePageNumber(PdfReadDocument document, int pageNumber) {
        if (pageNumber <= 0 || pageNumber > document.Pages.Count) {
            throw new ArgumentOutOfRangeException(nameof(pageNumber), pageNumber, "Page number must refer to an existing one-based PDF page.");
        }
    }
}
