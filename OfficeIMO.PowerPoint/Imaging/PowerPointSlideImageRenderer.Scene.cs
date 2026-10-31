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
    private static OfficeDrawing CreateDrawing(PowerPointSlide slide, PowerPointImageExportOptions options, List<OfficeImageExportDiagnostic> diagnostics, CancellationToken token = default, PowerPointEmbeddedImageReadBudget? sharedImageBudget = null) {
        token.ThrowIfCancellationRequested();
        (double width, double height) = GetSlideSizePoints(slide);
        OfficeDrawing drawing = new OfficeDrawing(width, height).ApplyImageExportOptions(options);
        var imageBudget = sharedImageBudget ?? new PowerPointEmbeddedImageReadBudget(
            options.MaximumEmbeddedImageBytes,
            options.MaximumTotalEmbeddedImageBytes);
        AddBackgroundRectangle(drawing, options.BackgroundColor, fillGradient: null);

        if (options.IncludeSlideBackground) {
            AddResolvedBackground(slide, drawing, diagnostics, imageBudget);
        }

        if (options.IncludeSlideContent) {
            A.ColorScheme? colorScheme = GetSlideColorScheme(slide);
            AddSlideContent(slide.GetInheritedShapesForExport(), drawing, diagnostics,
                PowerPointShapeBoundsMapping.Identity, colorScheme, options, imageBudget, 0, token);
            AddSlideContent(slide.Shapes, drawing, diagnostics,
                PowerPointShapeBoundsMapping.Identity, colorScheme, options, imageBudget, 0, token);
        }

        return drawing;
    }

    private static void AddSlideContent(IEnumerable<PowerPointShape> shapes, OfficeDrawing drawing,
        List<OfficeImageExportDiagnostic> diagnostics, PowerPointShapeBoundsMapping mapping,
        A.ColorScheme? colorScheme, PowerPointImageExportOptions options,
        PowerPointEmbeddedImageReadBudget imageBudget, int groupDepth, CancellationToken token = default) {
        foreach (PowerPointShape shape in shapes) {
            token.ThrowIfCancellationRequested();
            if (shape.Hidden && !options.IncludeHiddenShapes) {
                continue;
            }
            if (!ShouldIncludeShape(shape, options)) {
                continue;
            }

            // Retain finite bleed content until the root slide viewport clips it. Each
            // component still receives a positive local frame through its existing owner.
            if (shape.TryGetExportBoundsPoints(out double rawX, out double rawY, out double rawWidth, out double rawHeight)) {
                double x = mapping.MapX(rawX), y = mapping.MapY(rawY);
                double width = mapping.MapWidth(rawWidth), height = mapping.MapHeight(rawHeight);
                if (width > 0 && height > 0 && (x < -0.000001D || y < -0.000001D
                    || x + width > drawing.Width + 0.000001D || y + height > drawing.Height + 0.000001D)) {
                    if (!mapping.HasTransformedAncestor && shape is not PowerPointGroupShape &&
                        !HasUnsupportedTransform(shape) &&
                        (x >= drawing.Width || y >= drawing.Height || x + width <= 0D || y + height <= 0D)) {
                        diagnostics.Add(new OfficeImageExportDiagnostic(OfficeImageExportDiagnosticSeverity.Warning,
                            OfficeDiagnosticCodes.InvalidShapeBounds,
                            "PowerPoint shape content lies entirely outside the visible slide bounds.",
                            DescribeShape(shape), OfficeConversionLossKind.Omission));
                    }
                    var local = new OfficeDrawing(width + 0.000001D, height + 0.000001D).ApplyImageExportOptions(options);
                    local.Fonts.AddRange(drawing.Fonts);
                    var placement = new PowerPointShapeBoundsMapping(mapping.MapX(0) - x, mapping.MapY(0) - y,
                        mapping.MapWidth(1), mapping.MapHeight(1), mapping.HasTransformedAncestor);
                    AddSlideContent(new[] { shape }, local, diagnostics, placement, colorScheme, options, imageBudget, groupDepth, token);
                    drawing.AddDrawingForClippedRendering(local, x, y, null);
                    continue;
                }
            }

            if (shape is PowerPointGroupShape groupShape) {
                AddGroupShape(drawing, groupShape, diagnostics, mapping, colorScheme, options, imageBudget, groupDepth, token);
            } else if (shape is PowerPointPicture picture) {
                AddImage(drawing, picture, diagnostics, mapping, imageBudget);
            } else if (shape is PowerPointTable table) {
                AddTable(drawing, table, diagnostics, mapping, colorScheme);
            } else if (shape is PowerPointChart chart) {
                AddChart(drawing, chart, diagnostics, mapping, colorScheme);
            } else if (shape is PowerPointSmartArt smartArt) {
                // Office caches the laid-out SmartArt shapes for exactly this purpose.
                if (smartArt.TryCreateCachedDrawing(out PowerPointGroupShape cachedDrawing, out string? cachedDrawingFailure)) {
                    AddGroupShape(drawing, cachedDrawing, diagnostics, mapping, colorScheme, options, imageBudget, groupDepth, token);
                } else {
                    AddSmartArt(drawing, smartArt, diagnostics, mapping, cachedDrawingFailure);
                }
            } else if (shape is PowerPointTextBox textBox) {
                AddTextBox(drawing, textBox, diagnostics, mapping, colorScheme, cancellationToken: token);
            } else if (shape is PowerPointAutoShape autoShape) {
                AddAutoShape(drawing, autoShape, diagnostics, mapping, colorScheme);
            } else if (shape is PowerPointConnectionShape connectionShape) {
                AddAutoShape(drawing, connectionShape, diagnostics, mapping, colorScheme);
            } else if (shape is PowerPointOleObject) {
                AddUnsupportedShapeDiagnostic(diagnostics, shape,
                    "Skipped a PowerPoint OLE object because embedded-object previews are not yet projected through OfficeIMO.Drawing.");
            } else if (HasUnsupportedTransform(shape)) {
                AddUnsupportedShapeDiagnostic(diagnostics, shape, "Skipped a PowerPoint shape because rotated or flipped slide content is not yet projected through OfficeIMO.Drawing.");
                continue;
            } else {
                AddUnsupportedShapeDiagnostic(diagnostics, shape, "Skipped a PowerPoint shape type that is not yet projected through OfficeIMO.Drawing.");
            }
        }
    }
}
