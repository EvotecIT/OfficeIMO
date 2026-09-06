using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace OfficeIMO.Drawing;

public static partial class OfficeSvgDrawingReader {
    private static bool IsEmptySvgClipGeometry(XElement element, double width, double height) {
        bool IsZero(string name, double reference) =>
            TryViewportLength(element, name, reference, out double value) && value == 0D;
        return element.Name.LocalName.ToLowerInvariant() switch {
            "rect" => IsZero("width", width) || IsZero("height", height),
            "circle" => IsZero("r", NormalizedSvgDiagonal(width, height)),
            "ellipse" => IsZero("rx", width) || IsZero("ry", height),
            "path" => string.IsNullOrWhiteSpace(element.Attribute("d")?.Value),
            "polygon" => string.IsNullOrWhiteSpace(element.Attribute("points")?.Value),
            _ => false
        };
    }

    // Transform the clip geometry, not the painted canvas. Inverse-transforming a fixed-size
    // intermediate bitmap would discard valid paint before the final clip can see it.
    private static bool TryCreateDestinationSvgClip(OfficeDrawingShape shape, OfficeTransform transform,
        out OfficeClipPath? path, out double x, out double y) {
        x = shape.X;
        y = shape.Y;
        if (shape.Shape.Kind == OfficeShapeKind.Path)
            path = OfficeClipPath.Path(CloseSvgClipContours(shape.Shape.PathCommands), shape.Shape.FillRule);
        else if (!TryCreateShapeClipPath(shape.Shape, out path)) return false;
        if (transform == OfficeTransform.Identity) return true;
        IEnumerable<OfficePathCommand> source;
        if (path!.Kind == OfficeClipPathKind.Path) source = path.Commands;
        else {
            IReadOnlyList<OfficePoint> points = CreateRoundedStrokeContour(path.Width, path.Height,
                path.Kind == OfficeClipPathKind.RoundedRectangle ? path.CornerRadius : 0D);
            source = points.Select((point, index) => index == 0 ? OfficePathCommand.MoveTo(point) : OfficePathCommand.LineTo(point))
                .Concat(new[] { OfficePathCommand.Close() });
        }
        OfficeTransform coordinates = OfficeTransform.Translate(shape.X, shape.Y).Then(transform);
        var commands = new List<OfficePathCommand>();
        double left = double.PositiveInfinity, top = double.PositiveInfinity;
        double right = double.NegativeInfinity, bottom = double.NegativeInfinity;
        foreach (OfficePathCommand command in source) {
            OfficePathCommand mapped = command.Kind switch {
                OfficePathCommandKind.MoveTo => OfficePathCommand.MoveTo(coordinates.TransformPoint(command.Point)),
                OfficePathCommandKind.LineTo => OfficePathCommand.LineTo(coordinates.TransformPoint(command.Point)),
                OfficePathCommandKind.QuadraticBezierTo => OfficePathCommand.QuadraticBezierTo(
                    coordinates.TransformPoint(command.ControlPoint1), coordinates.TransformPoint(command.Point)),
                OfficePathCommandKind.CubicBezierTo => OfficePathCommand.CubicBezierTo(
                    coordinates.TransformPoint(command.ControlPoint1), coordinates.TransformPoint(command.ControlPoint2),
                    coordinates.TransformPoint(command.Point)),
                _ => OfficePathCommand.Close()
            };
            commands.Add(mapped);
            IncludeCommandBounds(mapped, ref left, ref top, ref right, ref bottom);
        }
        if (double.IsNaN(left) || double.IsInfinity(left) || double.IsNaN(top) || double.IsInfinity(top) ||
            right <= left || bottom <= top) return false;
        for (int index = 0; index < commands.Count; index++) commands[index] = commands[index].Translate(left, top);
        path = OfficeClipPath.Path(commands, shape.Shape.FillRule);
        x = left;
        y = top;
        return true;
    }

    private static IEnumerable<OfficePathCommand> CloseSvgClipContours(IEnumerable<OfficePathCommand> commands) {
        bool open = false;
        foreach (OfficePathCommand command in commands) {
            if (command.Kind == OfficePathCommandKind.MoveTo && open) yield return OfficePathCommand.Close();
            yield return command;
            if (command.Kind == OfficePathCommandKind.MoveTo) open = true;
            else if (command.Kind == OfficePathCommandKind.Close) open = false;
        }
        if (open) yield return OfficePathCommand.Close();
    }
}
