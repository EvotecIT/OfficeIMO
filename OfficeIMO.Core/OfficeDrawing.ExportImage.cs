using System;
using System.Collections.Generic;
using System.Threading;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeDrawing {
    /// <summary>Exports a detached drawing through the shared density, encoding, font, diagnostic and safety contracts. Options define the drawing's logical units per inch.</summary>
    public OfficeImageExportResult ExportImage(OfficeImageExportFormat format, OfficeImageExportOptions? options = null, CancellationToken cancellationToken = default) {
        var effective = (options ?? new OfficeImageExportOptions()).CreateEffectiveImageExportOptions<OfficeImageExportOptions>();
        return OfficeImageExportExecutionScope.Run(effective, cancellationToken, token => {
            token.ThrowIfCancellationRequested();
            var drawing = Clone().ApplyImageExportOptions(effective);
            var diagnostics = new List<OfficeImageExportDiagnostic>();
            const string source = "Drawing";
            var codec = new OfficeRasterImageFallbackCodec(effective.ImageCodec, diagnostics, source);
            if (format == OfficeImageExportFormat.Svg) {
                double scale = effective.GetEffectiveScale(Width, Height);
                byte[] svg = OfficeDrawingSvgExporter.ToSvgBytes(drawing, scale, OfficeSvgSizeUnit.Pixel, codec,
                    resourceIdPrefix: null, maximumUtf8Bytes: effective.MaximumTotalEncodedBytes, cancellationToken: token);
                return effective.EnsureAccepted(new OfficeImageExportResult(format, (int)Math.Ceiling(Width * scale),
                    (int)Math.Ceiling(Height * scale), svg, source: source, diagnostics: diagnostics));
            }
            if (!format.IsRaster()) throw new ArgumentOutOfRangeException(nameof(format));
            var plan = OfficeRasterExportPlanner.Resolve(Width, Height, format, effective, source);
            if (plan.Diagnostic != null) diagnostics.Add(plan.Diagnostic);
            var image = OfficeDrawingRasterRenderer.Render(drawing, new OfficeDrawingRasterRenderOptions {
                Scale = plan.Limit.Scale, Background = effective.BackgroundColor, ImageCodec = codec,
                MaximumRasterPixels = effective.MaximumRasterPixels, TextShapingProvider = effective.TextShapingProvider,
                TextShapingLanguage = effective.TextShapingLanguage, DiagnosticSink = diagnostics,
                DiagnosticSource = source, CancellationToken = token
            });
            byte[] bytes = OfficeRasterImageEncoder.Encode(image, format, plan.CreateEncodingOptions());
            token.ThrowIfCancellationRequested();
            return effective.EnsureAccepted(new OfficeImageExportResult(format, image.Width, image.Height, bytes, source: source, diagnostics: diagnostics));
        });
    }
}
