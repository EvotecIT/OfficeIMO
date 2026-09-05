using OfficeIMO.Drawing;
using System.Threading;

namespace OfficeIMO.Pdf;

internal static class PdfImageExportEngine {
    internal static IReadOnlyList<OfficeImageExportResult> Export(
        Func<CancellationToken, PdfReadDocument> documentFactory,
        OfficeImageExportFormat format,
        PdfImageExportOptions options,
        Func<PdfReadDocument, PdfPageSelection?> selectionFactory,
        IReadOnlyList<OfficeImageExportDiagnostic>? initialDiagnostics = null,
        Func<IReadOnlyList<OfficeImageExportDiagnostic>>? diagnosticsFactory = null,
        CancellationToken cancellationToken = default) {
        var results = new List<OfficeImageExportResult>();
        ExportEach(
            documentFactory,
            format,
            options,
            selectionFactory,
            results.Add,
            initialDiagnostics,
            diagnosticsFactory,
            cancellationToken);
        return results.AsReadOnly();
    }

    internal static void ExportEach(
        Func<CancellationToken, PdfReadDocument> documentFactory,
        OfficeImageExportFormat format,
        PdfImageExportOptions options,
        Func<PdfReadDocument, PdfPageSelection?> selectionFactory,
        OfficeImageExportConsumer consumer,
        IReadOnlyList<OfficeImageExportDiagnostic>? initialDiagnostics = null,
        Func<IReadOnlyList<OfficeImageExportDiagnostic>>? diagnosticsFactory = null,
        CancellationToken cancellationToken = default) {
        Guard.NotNull(documentFactory, nameof(documentFactory));
        Guard.NotNull(options, nameof(options));
        Guard.NotNull(selectionFactory, nameof(selectionFactory));
        Guard.NotNull(consumer, nameof(consumer));
        options.Validate();
        using OfficeImageExportExecutionScope execution = OfficeImageExportExecutionScope.Start(
            options.RenderTimeout,
            cancellationToken);
        try {
            PdfReadDocument document = documentFactory(execution.Token);
            execution.ThrowIfCancellationRequested();
            IReadOnlyList<OfficeImageExportDiagnostic>? diagnostics = initialDiagnostics;
            if (diagnosticsFactory != null) {
                var combined = new List<OfficeImageExportDiagnostic>();
                if (initialDiagnostics != null) combined.AddRange(initialDiagnostics);
                combined.AddRange(diagnosticsFactory());
                diagnostics = combined.AsReadOnly();
            }
            ExportEach(
                document,
                format,
                options,
                selectionFactory(document),
                consumer,
                diagnostics,
                execution.Token);
            execution.ThrowIfCancellationRequested();
        } catch (OperationCanceledException exception) when (execution.IsTimeoutCancellation(exception)) {
            throw execution.CreateTimeoutException(exception);
        }
    }

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
        using OfficeImageExportExecutionScope execution = OfficeImageExportExecutionScope.Start(
            options.RenderTimeout,
            cancellationToken);
        try {
            var encodingBudget = new OfficeImageExportEncodingBudget(options.MaximumTotalEncodedBytes);
            OfficeImageExportResult result = ExportCore(
                page,
                format,
                options,
                pageNumber,
                initialDiagnostics,
                encodingBudget,
                execution.Token);
            execution.ThrowIfCancellationRequested();
            return result;
        } catch (OperationCanceledException exception) when (execution.IsTimeoutCancellation(exception)) {
            throw execution.CreateTimeoutException(exception);
        }
    }

    private static OfficeImageExportResult ExportCore(
        PdfReadPage page,
        OfficeImageExportFormat format,
        PdfImageExportOptions options,
        int? pageNumber,
        IReadOnlyList<OfficeImageExportDiagnostic>? initialDiagnostics,
        OfficeImageExportEncodingBudget encodingBudget,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();

        OfficeDrawing drawing = page.ToDrawing(cancellationToken);
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
            effective.Scale = effective.GetEffectiveScale(drawing.Width, drawing.Height);
            drawing = AddBackground(drawing, effective.BackgroundColor);
            byte[] svg = encodingBudget.EncodeWithinRemainingBudget(
                maximumBytes => OfficeDrawingSvgExporter.ToSvgBytes(
                    drawing,
                    effective.Scale,
                    OfficeSvgSizeUnit.Pixel,
                    fallbackCodec,
                    resourceIdPrefix: null,
                    maximumUtf8Bytes: maximumBytes,
                    cancellationToken: cancellationToken),
                cancellationToken);
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
            TextShapingProvider = effective.TextShapingProvider,
            TextShapingLanguage = effective.TextShapingLanguage,
            DiagnosticSink = diagnostics,
            DiagnosticSource = source,
            MaximumRasterPixels = effective.MaximumRasterPixels,
            CancellationToken = cancellationToken
        });
        byte[] bytes = OfficeRasterImageEncoder.Encode(
            raster,
            format,
            plan.CreateEncodingOptions(),
            encodingBudget,
            cancellationToken);
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
        using OfficeImageExportExecutionScope execution = OfficeImageExportExecutionScope.Start(
            options.RenderTimeout,
            cancellationToken);
        try {
            int[] pages = selection?.ToPageNumbers(document.Pages.Count, nameof(selection))
                ?? Enumerable.Range(1, document.Pages.Count).ToArray();
            var encodingBudget = new OfficeImageExportEncodingBudget(options.MaximumTotalEncodedBytes);

            OfficeImageExportBatchProcessor.ForEachOrdered(
                pages,
                options.MaximumDegreeOfParallelism,
                (pageNumber, _, token) => ExportCore(
                    document.Pages[pageNumber - 1],
                    format,
                    options,
                    pageNumber,
                    initialDiagnostics,
                    encodingBudget,
                    token),
                consumer,
                execution.Token,
                options);
            execution.ThrowIfCancellationRequested();
        } catch (OperationCanceledException exception) when (execution.IsTimeoutCancellation(exception)) {
            throw execution.CreateTimeoutException(exception);
        }
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
                    PdfRenderSupportLevel.Simplified => OfficeConversionLossKind.Approximation,
                    PdfRenderSupportLevel.Unsupported => OfficeConversionLossKind.Omission,
                    _ => OfficeConversionLossKind.None
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
                string.IsNullOrWhiteSpace(warning.Source) ? warning.Converter : warning.Source,
                warning.LossKind));
        }
        for (int index = 0; index < conversion.SourceConversionReports.Count; index++) {
            if (conversion.SourceConversionReports[index].HasLoss) {
                diagnostics.Add(new OfficeImageExportDiagnostic(
                    OfficeImageExportDiagnosticSeverity.Warning,
                    "SourceConversionLoss",
                    "An upstream conversion stage reported content loss. Inspect the source conversion report for details.",
                    "source-stage:" + (index + 1),
                    OfficeConversionLossKind.Approximation));
            }
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
