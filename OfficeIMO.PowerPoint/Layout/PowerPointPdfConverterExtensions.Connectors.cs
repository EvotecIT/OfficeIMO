using OfficeIMO.Drawing;
using PdfCore = OfficeIMO.Pdf;
using PptCore = OfficeIMO.PowerPoint;

namespace OfficeIMO.PowerPoint.Pdf;

public static partial class PowerPointPdfConverterExtensions {
    private static void RenderConnection(PdfCore.PdfPageCanvas canvas, PptCore.PowerPointConnectionShape source,
        double x, double y, double width, double height, int slideNumber, PowerPointToPdfOptions options) {
        string? preset = source.ShapeType.HasValue ? GetShapePresetName(source.ShapeType.ToOpenXml()) : null;
        OfficeShape? shape;
        if (preset is "line" or "straightConnector1") shape = OfficeShape.Line(0D, 0D, width, height);
        else if (TryCreateBentConnectorPath(preset, width, height, out var commands) ||
            TryCreateCurvedConnectorPath(preset, width, height, out commands)) shape = OfficeShape.Path(commands);
        else {
            AddWarning(options, slideNumber, "unsupported-connector", "Skipped unsupported PowerPoint connector geometry '" + preset + "'.",
                lossKind: OfficeConversionLossKind.Omission);
            return;
        }
        ApplyShapeAppearance(source, shape, slideNumber, options);
        shape.FillColor = null;
        shape.FillGradient = null;
        shape.FillRadialGradient = null;
        var frame = new OfficeImageFrameTransform(source.Rotation ?? 0D, x + width / 2D, y + height / 2D,
            source.HorizontalFlip == true, source.VerticalFlip == true);
        if (frame.HasTransform) canvas.Effect(frame.CreateDestinationTransform(), 1D, target => target.Shape(shape, x, y));
        else canvas.Shape(shape, x, y);
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
