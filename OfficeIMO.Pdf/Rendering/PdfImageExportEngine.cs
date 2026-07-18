using OfficeIMO.Drawing;
using System.Threading;

namespace OfficeIMO.Pdf;

internal static class PdfImageExportEngine {
    internal static OfficeImageExportResult Export(
        PdfReadPage page,
        OfficeImageExportFormat format,
        PdfImageExportOptions options,
        int? pageNumber = null,
        IReadOnlyList<OfficeImageExportDiagnostic>? initialDiagnostics = null,
        CancellationToken cancellationToken = default) {
        Guard.NotNull(page, nameof(page));
        Guard.NotNull(options, nameof(options));
        options.Validate();
        cancellationToken.ThrowIfCancellationRequested();

        OfficeDrawing drawing = page.ToDrawing();
        drawing.ApplyImageExportOptions(options);
        PdfImageExportOptions effective = options.Clone();
        double requestedScale = options.Scale;
        effective.Scale = options.ResolveScale(drawing);
        if (options.TargetDpi.HasValue && effective.Scale < requestedScale) {
            double effectiveDpi = effective.Scale * effective.LogicalUnitsPerInch;
            effective.RasterEncoding.DpiX = effectiveDpi;
            effective.RasterEncoding.DpiY = effectiveDpi;
        }
        // The target DPI has already been resolved into Scale. Keeping it on the clone would let
        // the shared validation step overwrite a stricter thumbnail scale.
        effective.TargetDpi = null;
        IReadOnlyList<PdfRenderCapabilityDiagnostic> capabilityDiagnostics =
            page.GetRenderCapabilityDiagnostics();
        var diagnostics = new List<OfficeImageExportDiagnostic>(
            (initialDiagnostics?.Count ?? 0) + capabilityDiagnostics.Count);
        if (initialDiagnostics != null) diagnostics.AddRange(initialDiagnostics);
        diagnostics.AddRange(MapDiagnostics(capabilityDiagnostics, pageNumber));
        string name = pageNumber.HasValue ? "Page " + pageNumber.Value : "Page";
        string source = pageNumber.HasValue ? "PDF page " + pageNumber.Value : "PDF page";
        drawing.AppendFontDiagnostics(diagnostics, source);
        var fallbackCodec = new OfficeRasterImageFallbackCodec(effective.ImageCodec, diagnostics, source);

        cancellationToken.ThrowIfCancellationRequested();
        if (format == OfficeImageExportFormat.Svg) {
            drawing = AddBackground(drawing, effective.BackgroundColor);
            byte[] svg = OfficeDrawingSvgExporter.ToSvgBytes(
                drawing,
                effective.Scale,
                OfficeSvgSizeUnit.Pixel,
                fallbackCodec);
            return options.EnsureAccepted(new OfficeImageExportResult(
                format,
                Scaled(drawing.Width, effective.Scale),
                Scaled(drawing.Height, effective.Scale),
                svg,
                name,
                source,
                diagnostics));
        }
        if (!format.IsRaster()) {
            throw new ArgumentOutOfRangeException(nameof(format), format, "Unsupported image export format.");
        }

        OfficeRasterExportPlan plan = OfficeRasterExportPlanner.Resolve(
            drawing.Width,
            drawing.Height,
            format,
            effective,
            source);
        if (plan.Diagnostic != null) diagnostics.Add(plan.Diagnostic);
        cancellationToken.ThrowIfCancellationRequested();
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(drawing, new OfficeDrawingRasterRenderOptions {
            Scale = plan.Limit.Scale,
            Background = effective.BackgroundColor,
            ImageCodec = fallbackCodec,
            CancellationToken = cancellationToken
        });
        byte[] bytes = OfficeRasterImageEncoder.Encode(raster, format, effective.RasterEncoding);
        cancellationToken.ThrowIfCancellationRequested();
        return options.EnsureAccepted(new OfficeImageExportResult(
            format,
            raster.Width,
            raster.Height,
            bytes,
            name,
            source,
            diagnostics));
    }

    internal static IReadOnlyList<OfficeImageExportResult> Export(
        PdfReadDocument document,
        OfficeImageExportFormat format,
        PdfImageExportOptions options,
        PdfPageSelection? selection,
        IReadOnlyList<OfficeImageExportDiagnostic>? initialDiagnostics = null,
        CancellationToken cancellationToken = default) {
        var results = new List<OfficeImageExportResult>();
        ExportEach(document, format, options, selection, results.Add, initialDiagnostics, cancellationToken);
        return results.AsReadOnly();
    }

    internal static void ExportEach(
        PdfReadDocument document,
        OfficeImageExportFormat format,
        PdfImageExportOptions options,
        PdfPageSelection? selection,
        OfficeImageExportConsumer consumer,
        IReadOnlyList<OfficeImageExportDiagnostic>? initialDiagnostics = null,
        CancellationToken cancellationToken = default) {
        Guard.NotNull(document, nameof(document));
        Guard.NotNull(options, nameof(options));
        Guard.NotNull(consumer, nameof(consumer));
        options.Validate();
        int[] pages = selection?.ToPageNumbers(document.Pages.Count, nameof(selection))
            ?? Enumerable.Range(1, document.Pages.Count).ToArray();

        OfficeImageExportBatchProcessor.ForEachOrdered(
            pages,
            options.MaximumDegreeOfParallelism,
            (pageNumber, _, token) => Export(
                document.Pages[pageNumber - 1],
                format,
                options,
                pageNumber,
                initialDiagnostics,
                token),
            consumer,
            cancellationToken,
            options);
    }

    private static List<OfficeImageExportDiagnostic> MapDiagnostics(
        IReadOnlyList<PdfRenderCapabilityDiagnostic> source,
        int? pageNumber) {
        var diagnostics = new List<OfficeImageExportDiagnostic>(source.Count);
        string diagnosticSource = pageNumber.HasValue ? "PDF page " + pageNumber.Value : "PDF page";
        for (int index = 0; index < source.Count; index++) {
            PdfRenderCapabilityDiagnostic diagnostic = source[index];
            OfficeImageExportDiagnosticSeverity severity =
                diagnostic.SupportLevel == PdfRenderSupportLevel.Supported
                    ? OfficeImageExportDiagnosticSeverity.Info
                    : OfficeImageExportDiagnosticSeverity.Warning;
            diagnostics.Add(new OfficeImageExportDiagnostic(
                severity,
                diagnostic.Code,
                diagnostic.Message,
                diagnosticSource,
                diagnostic.SupportLevel switch {
                    PdfRenderSupportLevel.Simplified => OfficeImageExportLossKind.Approximation,
                    PdfRenderSupportLevel.Unsupported => OfficeImageExportLossKind.Omission,
                    _ => OfficeImageExportLossKind.None
                }));
        }
        return diagnostics;
    }

    internal static IReadOnlyList<OfficeImageExportDiagnostic> MapConversionDiagnostics(
        PdfDocumentConversionResult conversion) {
        Guard.NotNull(conversion, nameof(conversion));
        var diagnostics = new List<OfficeImageExportDiagnostic>(conversion.Warnings.Count);
        for (int index = 0; index < conversion.Warnings.Count; index++) {
            PdfConversionWarning warning = conversion.Warnings[index];
            diagnostics.Add(new OfficeImageExportDiagnostic(
                warning.Severity switch {
                    PdfConversionWarningSeverity.Error => OfficeImageExportDiagnosticSeverity.Error,
                    PdfConversionWarningSeverity.Warning => OfficeImageExportDiagnosticSeverity.Warning,
                    _ => OfficeImageExportDiagnosticSeverity.Info
                },
                warning.Code,
                warning.Message,
                string.IsNullOrWhiteSpace(warning.Source) ? warning.Converter : warning.Source));
        }
        return diagnostics.AsReadOnly();
    }

    private static int Scaled(double value, double scale) =>
        Math.Max(1, checked((int)Math.Ceiling(value * scale)));

    private static OfficeDrawing AddBackground(OfficeDrawing drawing, OfficeColor color) {
        var composed = new OfficeDrawing(drawing.Width, drawing.Height);
        composed.Fonts.AddRange(drawing.Fonts);
        OfficeShape background = OfficeShape.Rectangle(drawing.Width, drawing.Height);
        background.FillColor = color;
        background.StrokeWidth = 0D;
        composed.AddShape(background, 0D, 0D);
        composed.AddClippedDrawing(
            drawing,
            0D,
            0D,
            OfficeClipPath.Rectangle(drawing.Width, drawing.Height));
        return composed;
    }
}
