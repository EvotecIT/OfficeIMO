using OfficeIMO.Drawing.Internal;
using System.Text;
using AngleSharp.Html.Dom;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

/// <summary>Direct HTML-to-PNG and HTML-to-SVG helpers backed by the shared HTML render scene.</summary>
public static partial class HtmlImageExportExtensions {
    /// <summary>Exports one selected surface from a prepared HTML DOM as PNG or SVG.</summary>
    public static OfficeImageExportResult ExportImage(this IHtmlDocument document, OfficeImageExportFormat format, HtmlRenderOptions? options = null, int pageIndex = 0) {
        HtmlRenderOptions resolved = Normalize(options, pageIndex);
        HtmlRenderDocument rendered = HtmlRenderEngine.Render(document, resolved);
        if (pageIndex >= rendered.Pages.Count) throw new ArgumentOutOfRangeException(nameof(pageIndex), "The selected HTML render page does not exist.");
        return RenderPage(rendered.Pages[pageIndex], format, resolved, rendered.Diagnostics, CancellationToken.None);
    }

    /// <summary>Exports all surfaces from a prepared HTML DOM as PNG or SVG.</summary>
    public static IReadOnlyList<OfficeImageExportResult> ExportImages(this IHtmlDocument document, OfficeImageExportFormat format, HtmlRenderOptions? options = null) {
        HtmlRenderOptions resolved = Normalize(options, 0);
        HtmlRenderDocument rendered = HtmlRenderEngine.Render(document, resolved);
        var results = new List<OfficeImageExportResult>(rendered.Pages.Count);
        foreach (HtmlRenderPage page in rendered.Pages) results.Add(RenderPage(page, format, resolved, rendered.Diagnostics, CancellationToken.None));
        return results.AsReadOnly();
    }

    /// <summary>Asynchronously exports one selected surface from a prepared HTML DOM.</summary>
    public static async Task<OfficeImageExportResult> ExportImageAsync(this IHtmlDocument document, OfficeImageExportFormat format, HtmlRenderOptions? options = null, int pageIndex = 0, CancellationToken cancellationToken = default) {
        HtmlRenderOptions resolved = Normalize(options, pageIndex);
        HtmlRenderDocument rendered = await HtmlRenderEngine.RenderAsync(document, resolved, cancellationToken).ConfigureAwait(false);
        cancellationToken.ThrowIfCancellationRequested();
        if (pageIndex >= rendered.Pages.Count) throw new ArgumentOutOfRangeException(nameof(pageIndex), "The selected HTML render page does not exist.");
        return RenderPage(rendered.Pages[pageIndex], format, resolved, rendered.Diagnostics, cancellationToken);
    }

    /// <summary>Asynchronously exports all surfaces from a prepared HTML DOM.</summary>
    public static async Task<IReadOnlyList<OfficeImageExportResult>> ExportImagesAsync(this IHtmlDocument document, OfficeImageExportFormat format, HtmlRenderOptions? options = null, CancellationToken cancellationToken = default) {
        HtmlRenderOptions resolved = Normalize(options, 0);
        HtmlRenderDocument rendered = await HtmlRenderEngine.RenderAsync(document, resolved, cancellationToken).ConfigureAwait(false);
        var results = new List<OfficeImageExportResult>(rendered.Pages.Count);
        foreach (HtmlRenderPage page in rendered.Pages) {
            cancellationToken.ThrowIfCancellationRequested();
            results.Add(RenderPage(page, format, resolved, rendered.Diagnostics, cancellationToken));
        }
        return results.AsReadOnly();
    }

    /// <summary>Exports one selected HTML surface as PNG or SVG.</summary>
    public static OfficeImageExportResult ExportImage(this string html, OfficeImageExportFormat format, HtmlRenderOptions? options = null, int pageIndex = 0) {
        HtmlRenderOptions resolved = Normalize(options, pageIndex);
        HtmlRenderDocument rendered = HtmlRenderEngine.Render(html, resolved);
        if (pageIndex >= rendered.Pages.Count) throw new ArgumentOutOfRangeException(nameof(pageIndex), "The selected HTML render page does not exist.");
        return RenderPage(rendered.Pages[pageIndex], format, resolved, rendered.Diagnostics, CancellationToken.None);
    }

    /// <summary>Exports every paged HTML surface, or the single continuous surface, as PNG or SVG.</summary>
    public static IReadOnlyList<OfficeImageExportResult> ExportImages(this string html, OfficeImageExportFormat format, HtmlRenderOptions? options = null) {
        HtmlRenderOptions resolved = Normalize(options, 0);
        HtmlRenderDocument rendered = HtmlRenderEngine.Render(html, resolved);
        var results = new List<OfficeImageExportResult>(rendered.Pages.Count);
        foreach (HtmlRenderPage page in rendered.Pages) results.Add(RenderPage(page, format, resolved, rendered.Diagnostics, CancellationToken.None));
        return results.AsReadOnly();
    }

    /// <summary>Asynchronously resolves resources and exports one selected HTML surface as PNG or SVG.</summary>
    public static async Task<OfficeImageExportResult> ExportImageAsync(this string html, OfficeImageExportFormat format, HtmlRenderOptions? options = null, int pageIndex = 0, CancellationToken cancellationToken = default) {
        HtmlRenderOptions resolved = Normalize(options, pageIndex);
        HtmlRenderDocument rendered = await HtmlRenderEngine.RenderAsync(html, resolved, cancellationToken).ConfigureAwait(false);
        cancellationToken.ThrowIfCancellationRequested();
        if (pageIndex >= rendered.Pages.Count) throw new ArgumentOutOfRangeException(nameof(pageIndex), "The selected HTML render page does not exist.");
        return RenderPage(rendered.Pages[pageIndex], format, resolved, rendered.Diagnostics, cancellationToken);
    }

    /// <summary>Asynchronously resolves resources and exports every HTML surface.</summary>
    public static async Task<IReadOnlyList<OfficeImageExportResult>> ExportImagesAsync(this string html, OfficeImageExportFormat format, HtmlRenderOptions? options = null, CancellationToken cancellationToken = default) {
        HtmlRenderOptions resolved = Normalize(options, 0);
        HtmlRenderDocument rendered = await HtmlRenderEngine.RenderAsync(html, resolved, cancellationToken).ConfigureAwait(false);
        var results = new List<OfficeImageExportResult>(rendered.Pages.Count);
        foreach (HtmlRenderPage page in rendered.Pages) {
            cancellationToken.ThrowIfCancellationRequested();
            results.Add(RenderPage(page, format, resolved, rendered.Diagnostics, cancellationToken));
        }
        return results.AsReadOnly();
    }

    /// <summary>Renders one HTML surface to dependency-free PNG bytes.</summary>
    /// <example><code>byte[] png = html.ToPng();</code></example>
    public static byte[] ToPng(this string html, HtmlRenderOptions? options = null, int pageIndex = 0) =>
        html.ToPngResult(options, pageIndex).Bytes;

    /// <summary>Renders one HTML surface to SVG text.</summary>
    /// <example><code>string svg = html.ToSvg();</code></example>
    public static string ToSvg(this string html, HtmlRenderOptions? options = null, int pageIndex = 0) =>
        Encoding.UTF8.GetString(html.ToSvgResult(options, pageIndex).Bytes);

    /// <summary>Asynchronously resolves resources and renders one HTML surface to PNG bytes.</summary>
    public static async Task<byte[]> ToPngAsync(this string html, HtmlRenderOptions? options = null, int pageIndex = 0, CancellationToken cancellationToken = default) =>
        (await html.ToPngResultAsync(options, pageIndex, cancellationToken).ConfigureAwait(false)).Bytes;

    /// <summary>Asynchronously resolves resources and renders one HTML surface to SVG text.</summary>
    public static async Task<string> ToSvgAsync(this string html, HtmlRenderOptions? options = null, int pageIndex = 0, CancellationToken cancellationToken = default) =>
        Encoding.UTF8.GetString((await html.ToSvgResultAsync(options, pageIndex, cancellationToken).ConfigureAwait(false)).Bytes);

    /// <summary>Saves one HTML surface as a PNG file.</summary>
    public static void SaveAsPng(this string html, string path, HtmlRenderOptions? options = null, int pageIndex = 0) => WriteFile(path, html.ToPng(options, pageIndex));

    /// <summary>Saves one HTML surface as an SVG file.</summary>
    public static void SaveAsSvg(this string html, string path, HtmlRenderOptions? options = null, int pageIndex = 0) => WriteFile(path, Encoding.UTF8.GetBytes(html.ToSvg(options, pageIndex)));

    /// <summary>Writes one HTML surface as PNG to a stream.</summary>
    public static void SaveAsPng(this string html, Stream stream, HtmlRenderOptions? options = null, int pageIndex = 0) => WriteStream(stream, html.ToPng(options, pageIndex));

    /// <summary>Writes one HTML surface as SVG to a stream.</summary>
    public static void SaveAsSvg(this string html, Stream stream, HtmlRenderOptions? options = null, int pageIndex = 0) => WriteStream(stream, Encoding.UTF8.GetBytes(html.ToSvg(options, pageIndex)));

    /// <summary>Asynchronously resolves resources and saves one HTML surface as a PNG file.</summary>
    public static async Task SaveAsPngAsync(this string html, string path, HtmlRenderOptions? options = null, int pageIndex = 0, CancellationToken cancellationToken = default) =>
        await WriteFileAsync(path, await html.ToPngAsync(options, pageIndex, cancellationToken).ConfigureAwait(false), cancellationToken).ConfigureAwait(false);

    /// <summary>Asynchronously resolves resources and saves one HTML surface as an SVG file.</summary>
    public static async Task SaveAsSvgAsync(this string html, string path, HtmlRenderOptions? options = null, int pageIndex = 0, CancellationToken cancellationToken = default) =>
        await WriteFileAsync(path, Encoding.UTF8.GetBytes(await html.ToSvgAsync(options, pageIndex, cancellationToken).ConfigureAwait(false)), cancellationToken).ConfigureAwait(false);

    /// <summary>Asynchronously resolves resources and writes one HTML surface as PNG to a stream.</summary>
    public static async Task SaveAsPngAsync(this string html, Stream stream, HtmlRenderOptions? options = null, int pageIndex = 0, CancellationToken cancellationToken = default) =>
        await WriteStreamAsync(stream, await html.ToPngAsync(options, pageIndex, cancellationToken).ConfigureAwait(false), cancellationToken).ConfigureAwait(false);

    /// <summary>Asynchronously resolves resources and writes one HTML surface as SVG to a stream.</summary>
    public static async Task SaveAsSvgAsync(this string html, Stream stream, HtmlRenderOptions? options = null, int pageIndex = 0, CancellationToken cancellationToken = default) =>
        await WriteStreamAsync(stream, Encoding.UTF8.GetBytes(await html.ToSvgAsync(options, pageIndex, cancellationToken).ConfigureAwait(false)), cancellationToken).ConfigureAwait(false);

    private static OfficeImageExportResult RenderPage(HtmlRenderPage page, OfficeImageExportFormat format, HtmlRenderOptions options, HtmlDiagnosticReport diagnostics, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        OfficeDrawing drawing = page.CreateDrawing(cancellationToken);
        byte[] bytes = format == OfficeImageExportFormat.Svg
            ? OfficeDrawingSvgExporter.ToSvgBytes(drawing, options.Scale, OfficeSvgSizeUnit.Pixel)
            : OfficeDrawingRasterRenderer.ToPng(drawing, options.Scale, options.BackgroundColor);
        cancellationToken.ThrowIfCancellationRequested();
        return new OfficeImageExportResult(
            format,
            Math.Max(1, (int)Math.Ceiling(page.Width * options.Scale)),
            Math.Max(1, (int)Math.Ceiling(page.Height * options.Scale)),
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
                : diagnostic.Severity == HtmlDiagnosticSeverity.Warning ? OfficeImageExportDiagnosticSeverity.Warning : OfficeImageExportDiagnosticSeverity.Info;
            diagnostics.Add(new OfficeImageExportDiagnostic(severity, diagnostic.Code, diagnostic.Message, diagnostic.Source));
        }
        return diagnostics.AsReadOnly();
    }

    private static HtmlRenderOptions Normalize(HtmlRenderOptions? options, int pageIndex) {
        HtmlRenderOptions resolved = options?.Clone() ?? new HtmlRenderOptions();
        resolved.Validate();
        if (pageIndex < 0) throw new ArgumentOutOfRangeException(nameof(pageIndex), "HTML render page index cannot be negative.");
        return resolved;
    }

    private static void WriteFile(string path, byte[] bytes) {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("An output path is required.", nameof(path));
        OfficeFileCommit.WriteAllBytes(path, bytes);
    }

    private static void WriteStream(Stream stream, byte[] bytes) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanWrite) throw new ArgumentException("The output stream must be writable.", nameof(stream));
        stream.Write(bytes, 0, bytes.Length);
    }

    private static async Task WriteFileAsync(string path, byte[] bytes, CancellationToken cancellationToken) {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("An output path is required.", nameof(path));
        using var stream = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None, 81920, true);
        await stream.WriteAsync(bytes, 0, bytes.Length, cancellationToken).ConfigureAwait(false);
    }

    private static async Task WriteStreamAsync(Stream stream, byte[] bytes, CancellationToken cancellationToken) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanWrite) throw new ArgumentException("The output stream must be writable.", nameof(stream));
        await stream.WriteAsync(bytes, 0, bytes.Length, cancellationToken).ConfigureAwait(false);
    }
}
