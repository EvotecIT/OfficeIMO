using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint;

internal static partial class PowerPointSlideImageRenderer {
    internal static OfficeImageExportResult Render(
        PowerPointSlide slide,
        OfficeImageExportFormat format,
        PowerPointImageExportOptions options,
        CancellationToken cancellationToken = default,
        OfficeImageExportEncodingBudget? encodingBudget = null, PowerPointSlideVisualInfo? preparedSnapshot = null) {
        if (slide == null) {
            throw new ArgumentNullException(nameof(slide));
        }

        if (options == null) {
            throw new ArgumentNullException(nameof(options));
        }

        return OfficeImageExportExecutionScope.Run(
            options,
            cancellationToken,
            token => RenderCore(slide, format, options, token, encodingBudget, preparedSnapshot));
    }

    private static OfficeImageExportResult RenderCore(
        PowerPointSlide slide,
        OfficeImageExportFormat format,
        PowerPointImageExportOptions options,
        CancellationToken cancellationToken,
        OfficeImageExportEncodingBudget? encodingBudget, PowerPointSlideVisualInfo? preparedSnapshot) {
        cancellationToken.ThrowIfCancellationRequested();
        PowerPointSlideVisualInfo snapshot = preparedSnapshot ?? CreateSnapshot(slide, options, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        OfficeDrawing drawing = snapshot.Drawing;

        if (format == OfficeImageExportFormat.Svg) {
            List<OfficeImageExportDiagnostic> diagnostics = new List<OfficeImageExportDiagnostic>(snapshot.Diagnostics);
            var fallbackCodec = new OfficeRasterImageFallbackCodec(options.ImageCodec, diagnostics, "PowerPoint slide");
            double scale = options.GetEffectiveScale(drawing.Width, drawing.Height);
            byte[] svg = encodingBudget == null
                ? OfficeDrawingSvgExporter.ToSvgBytes(
                    drawing,
                    scale,
                    OfficeSvgSizeUnit.Pixel,
                    fallbackCodec,
                    resourceIdPrefix: null,
                    maximumUtf8Bytes: options.MaximumTotalEncodedBytes,
                    cancellationToken)
                : encodingBudget.EncodeWithinRemainingBudget(
                    remaining => OfficeDrawingSvgExporter.ToSvgBytes(
                        drawing,
                        scale,
                        OfficeSvgSizeUnit.Pixel,
                        fallbackCodec,
                        resourceIdPrefix: null,
                        maximumUtf8Bytes: remaining,
                        cancellationToken),
                    cancellationToken);
            return options.EnsureAccepted(new OfficeImageExportResult(format, ScaledWidth(drawing, scale), ScaledHeight(drawing, scale), svg, "Slide", "PowerPoint slide", diagnostics));
        }

        if (format.IsRaster()) {
            List<OfficeImageExportDiagnostic> diagnostics = new List<OfficeImageExportDiagnostic>(snapshot.Diagnostics);
            const string source = "PowerPoint slide";
            OfficeRasterExportPlan plan = OfficeRasterExportPlanner.Resolve(
                drawing.Width,
                drawing.Height,
                format,
                options,
                source);
            if (plan.Diagnostic != null) diagnostics.Add(plan.Diagnostic);
            var fallbackCodec = new OfficeRasterImageFallbackCodec(options.ImageCodec, diagnostics, source);
            OfficeRasterImage image = OfficeDrawingRasterRenderer.Render(drawing, new OfficeDrawingRasterRenderOptions {
                // Retain the pixel snapping and sampling of the previous page-image route.
                PdfRasterConventions = true,
                Scale = plan.Limit.Scale,
                Background = options.BackgroundColor,
                ImageCodec = fallbackCodec,
                TextShapingProvider = options.TextShapingProvider,
                TextShapingLanguage = options.TextShapingLanguage,
                DiagnosticSink = diagnostics,
                DiagnosticSource = source,
                CancellationToken = cancellationToken
            });
            byte[] bytes = encodingBudget == null
                ? OfficeRasterImageEncoder.Encode(
                    image,
                    format,
                    plan.CreateEncodingOptions(),
                    options.MaximumTotalEncodedBytes,
                    cancellationToken)
                : OfficeRasterImageEncoder.Encode(
                    image,
                    format,
                    plan.CreateEncodingOptions(),
                    encodingBudget,
                    cancellationToken);
            cancellationToken.ThrowIfCancellationRequested();
            return options.EnsureAccepted(new OfficeImageExportResult(format, image.Width, image.Height, bytes, "Slide", source, diagnostics));
        }

        throw new ArgumentOutOfRangeException(nameof(format), format, "Unsupported image export format.");
    }

    internal static PowerPointSlideVisualInfo CreateSnapshot(PowerPointSlide slide, PowerPointImageExportOptions options,
        CancellationToken cancellationToken = default, PowerPointEmbeddedImageReadBudget? imageBudget = null) {
        if (slide == null) throw new ArgumentNullException(nameof(slide));
        if (options == null) throw new ArgumentNullException(nameof(options));
        return OfficeImageExportExecutionScope.Run(options, cancellationToken, token => {
            token.ThrowIfCancellationRequested();
            options.Validate();
            var diagnostics = new List<OfficeImageExportDiagnostic>();
            OfficeDrawing drawing = CreateDrawing(slide, options, diagnostics, token, imageBudget);
            drawing.AppendFontDiagnostics(diagnostics, "PowerPoint slide");
            token.ThrowIfCancellationRequested();
            return new PowerPointSlideVisualInfo(drawing, diagnostics.AsReadOnly());
        });
    }

    private static int ScaledWidth(OfficeDrawing drawing, double scale) =>
        Math.Max(1, (int)Math.Ceiling(drawing.Width * scale));

    private static int ScaledHeight(OfficeDrawing drawing, double scale) =>
        Math.Max(1, (int)Math.Ceiling(drawing.Height * scale));

    private static int UnscaledWidth(OfficeDrawing drawing) =>
        Math.Max(1, (int)Math.Ceiling(drawing.Width));

    private static int UnscaledHeight(OfficeDrawing drawing) =>
        Math.Max(1, (int)Math.Ceiling(drawing.Height));
}
