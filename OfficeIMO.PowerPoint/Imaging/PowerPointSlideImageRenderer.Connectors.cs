using System;
using System.Collections.Generic;
using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    internal static partial class PowerPointSlideImageRenderer {
        private static bool TryAddZeroThicknessLine(OfficeDrawing drawing, PowerPointShape shape, List<OfficeImageExportDiagnostic> diagnostics, PowerPointShapeBoundsMapping mapping, string? presetName, A.ColorScheme? colorScheme) {
            if (!shape.TryGetBoundsPoints(out double rawLeft, out double rawTop, out double rawWidth, out double rawHeight)) {
                AddUnsupportedShapeDiagnostic(diagnostics, shape, "Skipped a PowerPoint shape because its bounds are outside the slide drawing canvas.");
                return true;
            }

            if (rawWidth > 0D && rawHeight > 0D) {
                return false;
            }

            double left = mapping.MapX(rawLeft);
            double top = mapping.MapY(rawTop);
            double width = mapping.MapWidth(rawWidth);
            double height = mapping.MapHeight(rawHeight);
            if ((width <= 0D && height <= 0D) ||
                left < 0D ||
                top < 0D ||
                left + Math.Max(1D, width) > drawing.Width ||
                top + Math.Max(1D, height) > drawing.Height) {
                AddUnsupportedShapeDiagnostic(diagnostics, shape, "Skipped a PowerPoint line shape because its bounds are outside the slide drawing canvas.");
                return true;
            }

            if (!OfficeShapePresets.TryCreate(presetName, Math.Max(1D, width), Math.Max(1D, height), out OfficeShape? drawingShape) ||
                drawingShape == null ||
                drawingShape.Kind != OfficeShapeKind.Line) {
                AddUnsupportedShapeDiagnostic(diagnostics, shape, "Skipped a PowerPoint line shape geometry that is not yet projected through OfficeIMO.Drawing.");
                return true;
            }

            drawingShape = OfficeShape.Line(0D, 0D, Math.Max(0D, width), Math.Max(0D, height));
            ApplyShapeStyle(drawingShape, shape, colorScheme, mapping, diagnostics);
            ApplyShapeTransform(drawingShape, shape, Math.Max(1D, width), Math.Max(1D, height));
            drawing.AddShape(drawingShape, left, top);
            return true;
        }

        private static bool IsZeroThicknessLinePreset(string? presetName) =>
            string.Equals(presetName, "line", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(presetName, "straightConnector1", StringComparison.OrdinalIgnoreCase);

        private static bool IsBentConnectorPreset(string? presetName) =>
            string.Equals(presetName, "bentConnector2", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(presetName, "bentConnector3", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(presetName, "bentConnector4", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(presetName, "bentConnector5", StringComparison.OrdinalIgnoreCase);

        private static bool IsCurvedConnectorPreset(string? presetName) =>
            string.Equals(presetName, "curvedConnector2", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(presetName, "curvedConnector3", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(presetName, "curvedConnector4", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(presetName, "curvedConnector5", StringComparison.OrdinalIgnoreCase);

        private static bool TryAddBentConnector(OfficeDrawing drawing, PowerPointShape shape, List<OfficeImageExportDiagnostic> diagnostics, PowerPointShapeBoundsMapping mapping, A.ColorScheme? colorScheme) {
            if (!TryGetBounds(shape, drawing, diagnostics, mapping, out double left, out double top, out double width, out double height)) {
                return true;
            }

            if (width <= 0D || height <= 0D) {
                AddUnsupportedShapeDiagnostic(diagnostics, shape, "Skipped a PowerPoint bent connector because its bounds do not describe a renderable elbow.");
                return true;
            }

            if (!TryCreateBentConnectorPath(GetAutoShapePresetName(shape), width, height, out List<OfficePathCommand> commands)) {
                AddUnsupportedShapeDiagnostic(diagnostics, shape, "Skipped a PowerPoint bent connector geometry that is not yet projected through OfficeIMO.Drawing.");
                return true;
            }

            OfficeShape drawingShape = OfficeShape.Path(commands);
            ApplyShapeStyle(drawingShape, shape, colorScheme, mapping, diagnostics);
            drawingShape.FillColor = null;
            drawingShape.FillGradient = null;
            ApplyShapeTransform(drawingShape, shape, width, height);
            drawing.AddShape(drawingShape, left, top);
            return true;
        }

        private static bool TryAddCurvedConnector(OfficeDrawing drawing, PowerPointShape shape, List<OfficeImageExportDiagnostic> diagnostics, PowerPointShapeBoundsMapping mapping, A.ColorScheme? colorScheme) {
            if (!TryGetBounds(shape, drawing, diagnostics, mapping, out double left, out double top, out double width, out double height)) {
                return true;
            }

            if (width <= 0D || height <= 0D) {
                AddUnsupportedShapeDiagnostic(diagnostics, shape, "Skipped a PowerPoint curved connector because its bounds do not describe a renderable curve.");
                return true;
            }

            if (!TryCreateCurvedConnectorPath(GetAutoShapePresetName(shape), width, height, out List<OfficePathCommand> commands)) {
                AddUnsupportedShapeDiagnostic(diagnostics, shape, "Skipped a PowerPoint curved connector geometry that is not yet projected through OfficeIMO.Drawing.");
                return true;
            }

            OfficeShape drawingShape = OfficeShape.Path(commands);
            ApplyShapeStyle(drawingShape, shape, colorScheme, mapping, diagnostics);
            drawingShape.FillColor = null;
            drawingShape.FillGradient = null;
            ApplyShapeTransform(drawingShape, shape, width, height);
            drawing.AddShape(drawingShape, left, top);
            return true;
        }

        private static bool TryCreateBentConnectorPath(string? presetName, double width, double height, out List<OfficePathCommand> commands) {
            commands = new List<OfficePathCommand>();
            if (!TryCreateBentConnectorWaypoints(presetName, width, height, out IReadOnlyList<OfficePoint>? waypoints, out bool useRightAngleFallback)) {
                return false;
            }

            List<OfficePoint> points = OfficeGeometry.BuildConnectorPolyline(
                new OfficePoint(0D, 0D),
                new OfficePoint(width, height),
                waypoints,
                useRightAngleFallback);

            commands.Add(OfficePathCommand.MoveTo(points[0]));
            for (int i = 1; i < points.Count; i++) {
                commands.Add(OfficePathCommand.LineTo(points[i]));
            }

            return true;
        }

        private static bool TryCreateCurvedConnectorPath(string? presetName, double width, double height, out List<OfficePathCommand> commands) {
            commands = new List<OfficePathCommand> {
                OfficePathCommand.MoveTo(0D, 0D)
            };

            if (string.Equals(presetName, "curvedConnector2", StringComparison.OrdinalIgnoreCase)) {
                commands.Add(OfficePathCommand.CubicBezierTo(0D, height, width, 0D, width, height));
                return true;
            }

            if (string.Equals(presetName, "curvedConnector3", StringComparison.OrdinalIgnoreCase)) {
                commands.Add(OfficePathCommand.CubicBezierTo(width * 0.5D, 0D, width * 0.5D, height, width * 0.5D, height));
                commands.Add(OfficePathCommand.CubicBezierTo(width * 0.5D, height, width, height * 0.5D, width, height));
                return true;
            }

            if (string.Equals(presetName, "curvedConnector4", StringComparison.OrdinalIgnoreCase)) {
                commands.Add(OfficePathCommand.CubicBezierTo(width * 0.5D, 0D, width * 0.5D, height * 0.5D, width * 0.5D, height * 0.5D));
                commands.Add(OfficePathCommand.CubicBezierTo(width * 0.5D, height * 0.5D, width, height * 0.5D, width, height * 0.5D));
                commands.Add(OfficePathCommand.CubicBezierTo(width, height * 0.5D, width * 0.5D, height, width, height));
                return true;
            }

            if (string.Equals(presetName, "curvedConnector5", StringComparison.OrdinalIgnoreCase)) {
                commands.Add(OfficePathCommand.CubicBezierTo(width / 3D, 0D, width / 3D, height / 2D, width / 3D, height / 2D));
                commands.Add(OfficePathCommand.CubicBezierTo(width / 3D, height / 2D, width * 2D / 3D, height / 2D, width * 2D / 3D, height / 2D));
                commands.Add(OfficePathCommand.CubicBezierTo(width * 2D / 3D, height / 2D, width * 2D / 3D, height, width * 2D / 3D, height));
                commands.Add(OfficePathCommand.CubicBezierTo(width * 2D / 3D, height, width, height, width, height));
                return true;
            }

            commands.Clear();
            return false;
        }

        private static bool TryCreateBentConnectorWaypoints(string? presetName, double width, double height, out IReadOnlyList<OfficePoint>? waypoints, out bool useRightAngleFallback) {
            useRightAngleFallback = false;
            waypoints = null;
            if (string.Equals(presetName, "bentConnector2", StringComparison.OrdinalIgnoreCase)) {
                useRightAngleFallback = true;
                return true;
            }

            if (string.Equals(presetName, "bentConnector3", StringComparison.OrdinalIgnoreCase)) {
                waypoints = new[] {
                    new OfficePoint(width / 2D, 0D),
                    new OfficePoint(width / 2D, height)
                };
                return true;
            }

            if (string.Equals(presetName, "bentConnector4", StringComparison.OrdinalIgnoreCase)) {
                waypoints = new[] {
                    new OfficePoint(width / 2D, 0D),
                    new OfficePoint(width / 2D, height / 2D),
                    new OfficePoint(width, height / 2D)
                };
                return true;
            }

            if (string.Equals(presetName, "bentConnector5", StringComparison.OrdinalIgnoreCase)) {
                waypoints = new[] {
                    new OfficePoint(width / 3D, 0D),
                    new OfficePoint(width / 3D, height / 2D),
                    new OfficePoint(width * 2D / 3D, height / 2D),
                    new OfficePoint(width * 2D / 3D, height)
                };
                return true;
            }

            return false;
        }
    }
}
