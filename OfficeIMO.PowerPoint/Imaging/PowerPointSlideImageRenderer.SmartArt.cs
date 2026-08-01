using System;
using System.Collections.Generic;
using OfficeIMO.Drawing;

namespace OfficeIMO.PowerPoint {
    internal static partial class PowerPointSlideImageRenderer {
        private static void AddSmartArt(OfficeDrawing drawing,
            PowerPointSmartArt smartArt,
            List<OfficeImageExportDiagnostic> diagnostics,
            PowerPointShapeBoundsMapping mapping) {
            if (!TryGetBounds(smartArt, drawing, diagnostics, mapping,
                    out double left, out double top, out double width,
                    out double height)) {
                return;
            }
            if (!smartArt.TryGetOfficeDiagramSnapshot(
                    out OfficeDiagramSnapshot source)) {
                AddUnsupportedShapeDiagnostic(diagnostics, smartArt,
                    "Skipped a PowerPoint SmartArt diagram because its semantic node data could not be read safely.");
                return;
            }

            try {
                var snapshot = new OfficeDiagramSnapshot(
                    source.Name, source.Kind, source.Nodes, width, height);
                OfficeDrawing smartArtDrawing =
                    OfficeDiagramDrawingRenderer.Render(snapshot,
                        includeBackground: false);
                var transform = new OfficeImageFrameTransform(
                    smartArt.Rotation ?? 0D,
                    left + width / 2D,
                    top + height / 2D,
                    smartArt.HorizontalFlip == true,
                    smartArt.VerticalFlip == true);
                if (transform.HasTransform) {
                    drawing.AddDrawing(smartArtDrawing, left, top, transform);
                } else {
                    drawing.AddDrawing(smartArtDrawing, left, top);
                }
            } catch (ArgumentException) {
                AddUnsupportedShapeDiagnostic(diagnostics, smartArt,
                    "Skipped a PowerPoint SmartArt diagram because its frame is too small for safe rendering.");
            } catch (InvalidOperationException) {
                AddUnsupportedShapeDiagnostic(diagnostics, smartArt,
                    "Skipped a PowerPoint SmartArt diagram because its semantic layout could not be rendered safely.");
            }
        }
    }
}
