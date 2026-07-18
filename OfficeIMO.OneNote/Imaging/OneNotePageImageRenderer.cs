using OfficeIMO.Drawing;

namespace OfficeIMO.OneNote;

internal static class OneNotePageImageRenderer {
    internal static OfficeImageExportResult Render(
        OneNotePage page,
        OfficeImageExportFormat format,
        OneNotePageRenderingOptions options,
        string? name = null,
        string? source = null) {
        if (page == null) throw new ArgumentNullException(nameof(page));
        if (options == null) throw new ArgumentNullException(nameof(options));
        OneNotePageVisualSnapshot snapshot = OneNotePageRenderer.CreateSnapshot(page, options);
        if (format == OfficeImageExportFormat.Svg) {
            var diagnostics = new List<OfficeImageExportDiagnostic>(snapshot.Diagnostics);
            var fallbackCodec = new OfficeRasterImageFallbackCodec(options.ImageCodec, diagnostics, source ?? "OneNote page");
            byte[] bytes = OfficeDrawingSvgExporter.ToSvgBytes(
                snapshot.Drawing,
                options.Scale,
                OfficeSvgSizeUnit.Pixel,
                fallbackCodec);
            return new OfficeImageExportResult(
                format,
                Scaled(snapshot.Drawing.Width, options.Scale),
                Scaled(snapshot.Drawing.Height, options.Scale),
                bytes,
                name ?? page.Title,
                source ?? "OneNote page",
                diagnostics);
        }
        if (format == OfficeImageExportFormat.Png || format == OfficeImageExportFormat.Jpeg ||
            format == OfficeImageExportFormat.Tiff || format == OfficeImageExportFormat.Webp) {
            var diagnostics = new List<OfficeImageExportDiagnostic>(snapshot.Diagnostics);
            OfficeRasterScaleLimit limit = ResolveRasterScaleLimit(snapshot.Drawing.Width, snapshot.Drawing.Height, format, options);
            if (limit.WasLimited) {
                diagnostics.Add(new OfficeImageExportDiagnostic(
                    OfficeImageExportDiagnosticSeverity.Warning,
                    "ONENOTE_IMAGE_RASTER_SCALE_LIMITED",
                    "The raster scale was reduced from " + Format(options.Scale) + " to " + Format(limit.Scale) +
                    " to respect the decoded-raster limits.",
                    source ?? "OneNote page"));
            }
            var fallbackCodec = new OfficeRasterImageFallbackCodec(options.ImageCodec, diagnostics, source ?? "OneNote page");
            OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(snapshot.Drawing, new OfficeDrawingRasterRenderOptions {
                Scale = limit.Scale,
                Background = options.BackgroundColor,
                ImageCodec = fallbackCodec
            });
            byte[] bytes = OfficeRasterImageEncoder.Encode(raster, format, options.RasterEncoding);
            return new OfficeImageExportResult(
                format,
                raster.Width,
                raster.Height,
                bytes,
                name ?? page.Title,
                source ?? "OneNote page",
                diagnostics);
        }
        throw new ArgumentOutOfRangeException(nameof(format));
    }

    internal static OfficeRasterScaleLimit ResolveRasterScaleLimit(
        double width,
        double height,
        OfficeImageExportFormat format,
        OneNotePageRenderingOptions options) {
        if (options == null) throw new ArgumentNullException(nameof(options));
        long maximumPixels = Math.Min(options.MaximumRasterPixels, OfficeRasterImageEncoder.GetMaximumPixelCount(format));
        return OfficeRasterScaleLimiter.Resolve(
            width,
            height,
            options.Scale,
            maximumPixels,
            OfficeRasterImageEncoder.GetMaximumDimension(format));
    }

    private static int Scaled(double value, double scale) => Math.Max(1, checked((int)Math.Ceiling(value * scale)));
    private static string Format(double value) => value.ToString("0.########", System.Globalization.CultureInfo.InvariantCulture);
}
