using OfficeIMO.Drawing;
using System.Threading;

namespace OfficeIMO.OneNote;

internal static class OneNotePageImageRenderer {
    internal static OfficeImageExportResult Render(
        OneNotePage page,
        OfficeImageExportFormat format,
        OneNotePageRenderingOptions options,
        string? name = null,
        string? source = null,
        CancellationToken cancellationToken = default) {
        if (page == null) throw new ArgumentNullException(nameof(page));
        if (options == null) throw new ArgumentNullException(nameof(options));
        cancellationToken.ThrowIfCancellationRequested();
        OneNotePageVisualSnapshot snapshot = OneNotePageRenderer.CreateSnapshot(page, options);
        cancellationToken.ThrowIfCancellationRequested();
        if (format == OfficeImageExportFormat.Svg) {
            var diagnostics = new List<OfficeImageExportDiagnostic>(snapshot.Diagnostics);
            var fallbackCodec = new OfficeRasterImageFallbackCodec(options.ImageCodec, diagnostics, source ?? "OneNote page");
            byte[] bytes = OfficeDrawingSvgExporter.ToSvgBytes(
                snapshot.Drawing,
                options.Scale,
                OfficeSvgSizeUnit.Pixel,
                fallbackCodec);
            return options.EnsureAccepted(new OfficeImageExportResult(
                format,
                Scaled(snapshot.Drawing.Width, options.Scale),
                Scaled(snapshot.Drawing.Height, options.Scale),
                bytes,
                name ?? page.Title,
                source ?? "OneNote page",
                diagnostics));
        }
        if (format == OfficeImageExportFormat.Png || format == OfficeImageExportFormat.Jpeg ||
            format == OfficeImageExportFormat.Tiff || format == OfficeImageExportFormat.Webp) {
            var diagnostics = new List<OfficeImageExportDiagnostic>(snapshot.Diagnostics);
            OfficeRasterExportPlan plan = OfficeRasterExportPlanner.Resolve(
                snapshot.Drawing.Width,
                snapshot.Drawing.Height,
                format,
                options,
                source ?? "OneNote page");
            if (plan.Diagnostic != null) diagnostics.Add(plan.Diagnostic);
            var fallbackCodec = new OfficeRasterImageFallbackCodec(options.ImageCodec, diagnostics, source ?? "OneNote page");
            OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(snapshot.Drawing, new OfficeDrawingRasterRenderOptions {
                Scale = plan.Limit.Scale,
                Background = options.BackgroundColor,
                ImageCodec = fallbackCodec,
                CancellationToken = cancellationToken
            });
            byte[] bytes = OfficeRasterImageEncoder.Encode(raster, format, options.RasterEncoding);
            cancellationToken.ThrowIfCancellationRequested();
            return options.EnsureAccepted(new OfficeImageExportResult(
                format,
                raster.Width,
                raster.Height,
                bytes,
                name ?? page.Title,
                source ?? "OneNote page",
                diagnostics));
        }
        throw new ArgumentOutOfRangeException(nameof(format));
    }

    internal static OfficeRasterScaleLimit ResolveRasterScaleLimit(
        double width,
        double height,
        OfficeImageExportFormat format,
        OneNotePageRenderingOptions options) {
        if (options == null) throw new ArgumentNullException(nameof(options));
        return OfficeRasterExportPlanner.Resolve(width, height, format, options).Limit;
    }

    private static int Scaled(double value, double scale) => Math.Max(1, checked((int)Math.Ceiling(value * scale)));
}
