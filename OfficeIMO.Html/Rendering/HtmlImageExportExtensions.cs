using System.Text;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

/// <summary>
/// Direct HTML to PNG and SVG helpers backed by the shared HTML render model.
/// </summary>
public static class HtmlImageExportExtensions {
    /// <summary>Exports the selected HTML surface as PNG or SVG.</summary>
    public static OfficeImageExportResult ExportImage(this string html, OfficeImageExportFormat format, HtmlImageExportOptions? options = null) {
        HtmlImageExportOptions resolved = Normalize(options);
        HtmlRenderDocument rendered = HtmlRenderEngine.Render(html, resolved);
        if (resolved.PageIndex >= rendered.Pages.Count) {
            throw new ArgumentOutOfRangeException(nameof(options), "The selected HTML render page does not exist.");
        }

        return RenderPage(rendered.Pages[resolved.PageIndex], format, resolved, rendered.Diagnostics, CancellationToken.None);
    }

    /// <summary>Exports every paged HTML surface, or the single continuous surface, as PNG or SVG.</summary>
    public static IReadOnlyList<OfficeImageExportResult> ExportImages(this string html, OfficeImageExportFormat format, HtmlImageExportOptions? options = null) {
        HtmlImageExportOptions resolved = Normalize(options);
        HtmlRenderDocument rendered = HtmlRenderEngine.Render(html, resolved);
        var results = new List<OfficeImageExportResult>(rendered.Pages.Count);
        foreach (HtmlRenderPage page in rendered.Pages) {
            results.Add(RenderPage(page, format, resolved, rendered.Diagnostics, CancellationToken.None));
        }

        return results.AsReadOnly();
    }

    /// <summary>Asynchronously resolves resources and exports the selected HTML surface as PNG or SVG.</summary>
    public static async Task<OfficeImageExportResult> ExportImageAsync(this string html, OfficeImageExportFormat format, HtmlImageExportOptions? options = null, CancellationToken cancellationToken = default) {
        HtmlImageExportOptions resolved = Normalize(options);
        HtmlRenderDocument rendered = await HtmlRenderEngine.RenderAsync(html, resolved, cancellationToken).ConfigureAwait(false);
        cancellationToken.ThrowIfCancellationRequested();
        if (resolved.PageIndex >= rendered.Pages.Count) throw new ArgumentOutOfRangeException(nameof(options), "The selected HTML render page does not exist.");
        OfficeImageExportResult result = RenderPage(rendered.Pages[resolved.PageIndex], format, resolved, rendered.Diagnostics, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        return result;
    }

    /// <summary>Asynchronously resolves resources and exports every paged HTML surface, or the single continuous surface.</summary>
    public static async Task<IReadOnlyList<OfficeImageExportResult>> ExportImagesAsync(this string html, OfficeImageExportFormat format, HtmlImageExportOptions? options = null, CancellationToken cancellationToken = default) {
        HtmlImageExportOptions resolved = Normalize(options);
        HtmlRenderDocument rendered = await HtmlRenderEngine.RenderAsync(html, resolved, cancellationToken).ConfigureAwait(false);
        var results = new List<OfficeImageExportResult>(rendered.Pages.Count);
        foreach (HtmlRenderPage page in rendered.Pages) {
            cancellationToken.ThrowIfCancellationRequested();
            results.Add(RenderPage(page, format, resolved, rendered.Diagnostics, cancellationToken));
            cancellationToken.ThrowIfCancellationRequested();
        }

        return results.AsReadOnly();
    }

    /// <summary>Renders the selected HTML surface to dependency-free PNG bytes.</summary>
    public static byte[] ToPng(this string html, HtmlImageExportOptions? options = null) =>
        ExportImage(html, OfficeImageExportFormat.Png, options).Bytes;

    /// <summary>Renders the selected HTML surface to SVG text.</summary>
    public static string ToSvg(this string html, HtmlImageExportOptions? options = null) =>
        Encoding.UTF8.GetString(ExportImage(html, OfficeImageExportFormat.Svg, options).Bytes);

    /// <summary>Asynchronously resolves resources and renders the selected HTML surface to PNG bytes.</summary>
    public static async Task<byte[]> ToPngAsync(this string html, HtmlImageExportOptions? options = null, CancellationToken cancellationToken = default) =>
        (await ExportImageAsync(html, OfficeImageExportFormat.Png, options, cancellationToken).ConfigureAwait(false)).Bytes;

    /// <summary>Asynchronously resolves resources and renders the selected HTML surface to SVG text.</summary>
    public static async Task<string> ToSvgAsync(this string html, HtmlImageExportOptions? options = null, CancellationToken cancellationToken = default) =>
        Encoding.UTF8.GetString((await ExportImageAsync(html, OfficeImageExportFormat.Svg, options, cancellationToken).ConfigureAwait(false)).Bytes);

    /// <summary>Saves the selected HTML surface as a PNG file.</summary>
    public static void SaveAsPng(this string html, string path, HtmlImageExportOptions? options = null) => WriteFile(path, html.ToPng(options));

    /// <summary>Saves the selected HTML surface as an SVG file.</summary>
    public static void SaveAsSvg(this string html, string path, HtmlImageExportOptions? options = null) => WriteFile(path, Encoding.UTF8.GetBytes(html.ToSvg(options)));

    /// <summary>Writes the selected HTML surface as PNG to a stream.</summary>
    public static void SaveAsPng(this string html, Stream stream, HtmlImageExportOptions? options = null) => WriteStream(stream, html.ToPng(options));

    /// <summary>Writes the selected HTML surface as SVG to a stream.</summary>
    public static void SaveAsSvg(this string html, Stream stream, HtmlImageExportOptions? options = null) => WriteStream(stream, Encoding.UTF8.GetBytes(html.ToSvg(options)));

    /// <summary>Asynchronously resolves resources and saves the selected HTML surface as a PNG file.</summary>
    public static async Task SaveAsPngAsync(this string html, string path, HtmlImageExportOptions? options = null, CancellationToken cancellationToken = default) =>
        await WriteFileAsync(path, await html.ToPngAsync(options, cancellationToken).ConfigureAwait(false), cancellationToken).ConfigureAwait(false);

    /// <summary>Asynchronously resolves resources and saves the selected HTML surface as an SVG file.</summary>
    public static async Task SaveAsSvgAsync(this string html, string path, HtmlImageExportOptions? options = null, CancellationToken cancellationToken = default) =>
        await WriteFileAsync(path, Encoding.UTF8.GetBytes(await html.ToSvgAsync(options, cancellationToken).ConfigureAwait(false)), cancellationToken).ConfigureAwait(false);

    /// <summary>Asynchronously resolves resources and writes the selected HTML surface as PNG to a stream.</summary>
    public static async Task SaveAsPngAsync(this string html, Stream stream, HtmlImageExportOptions? options = null, CancellationToken cancellationToken = default) =>
        await WriteStreamAsync(stream, await html.ToPngAsync(options, cancellationToken).ConfigureAwait(false), cancellationToken).ConfigureAwait(false);

    /// <summary>Asynchronously resolves resources and writes the selected HTML surface as SVG to a stream.</summary>
    public static async Task SaveAsSvgAsync(this string html, Stream stream, HtmlImageExportOptions? options = null, CancellationToken cancellationToken = default) =>
        await WriteStreamAsync(stream, Encoding.UTF8.GetBytes(await html.ToSvgAsync(options, cancellationToken).ConfigureAwait(false)), cancellationToken).ConfigureAwait(false);

    private static OfficeImageExportResult RenderPage(HtmlRenderPage page, OfficeImageExportFormat format, HtmlImageExportOptions options, HtmlDiagnosticReport diagnostics, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        OfficeDrawing drawing = page.CreateDrawing(cancellationToken);
        byte[] bytes = format == OfficeImageExportFormat.Svg
            ? OfficeDrawingSvgExporter.ToSvgBytes(drawing, options.Scale)
            : OfficeDrawingRasterRenderer.ToPng(drawing, options.Scale, options.BackgroundColor);
        cancellationToken.ThrowIfCancellationRequested();
        int width = Math.Max(1, (int)Math.Ceiling(page.Width * options.Scale));
        int height = Math.Max(1, (int)Math.Ceiling(page.Height * options.Scale));
        return new OfficeImageExportResult(
            format,
            width,
            height,
            bytes,
            "Page " + page.PageNumber.ToString(System.Globalization.CultureInfo.InvariantCulture),
            "HTML render page " + page.PageNumber.ToString(System.Globalization.CultureInfo.InvariantCulture),
            MapDiagnostics(diagnostics));
    }

    private static IReadOnlyList<OfficeImageExportDiagnostic> MapDiagnostics(HtmlDiagnosticReport report) {
        var diagnostics = new List<OfficeImageExportDiagnostic>(report.Count);
        foreach (HtmlDiagnostic diagnostic in report.Diagnostics) {
            OfficeImageExportDiagnosticSeverity severity = diagnostic.Severity == HtmlDiagnosticSeverity.Error
                ? OfficeImageExportDiagnosticSeverity.Error
                : diagnostic.Severity == HtmlDiagnosticSeverity.Warning
                    ? OfficeImageExportDiagnosticSeverity.Warning
                    : OfficeImageExportDiagnosticSeverity.Info;
            diagnostics.Add(new OfficeImageExportDiagnostic(severity, diagnostic.Code, diagnostic.Message, diagnostic.Source));
        }

        return diagnostics.AsReadOnly();
    }

    private static HtmlImageExportOptions Normalize(HtmlImageExportOptions? options) {
        HtmlImageExportOptions resolved = options?.CloneImage() ?? new HtmlImageExportOptions();
        resolved.Validate();
        if (resolved.PageIndex < 0) {
            throw new ArgumentOutOfRangeException(nameof(options), "HTML render page index cannot be negative.");
        }

        return resolved;
    }

    private static void WriteFile(string path, byte[] bytes) {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("An output path is required.", nameof(path));
        File.WriteAllBytes(path, bytes);
    }

    private static void WriteStream(Stream stream, byte[] bytes) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanWrite) throw new ArgumentException("The output stream must be writable.", nameof(stream));
        stream.Write(bytes, 0, bytes.Length);
    }

    private static async Task WriteFileAsync(string path, byte[] bytes, CancellationToken cancellationToken) {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("An output path is required.", nameof(path));
        using var stream = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None, 81920, useAsync: true);
        await stream.WriteAsync(bytes, 0, bytes.Length, cancellationToken).ConfigureAwait(false);
    }

    private static async Task WriteStreamAsync(Stream stream, byte[] bytes, CancellationToken cancellationToken) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanWrite) throw new ArgumentException("The output stream must be writable.", nameof(stream));
        await stream.WriteAsync(bytes, 0, bytes.Length, cancellationToken).ConfigureAwait(false);
    }
}
