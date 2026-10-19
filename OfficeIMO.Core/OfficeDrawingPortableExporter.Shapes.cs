using System;
using System.Collections.Generic;
using System.Globalization;

namespace OfficeIMO.Drawing;

public static partial class OfficeDrawingPortableExporter {
    private sealed partial class ProjectionContext {
        internal Element ProjectShape(OfficeDrawingShape positioned, OfficeTransform parent, string location) {
            OfficeShape shape = positioned.Shape;
            Coordinate(positioned.X, location); Coordinate(positioned.Y, location);
            if (shape.FillGradient != null || shape.FillRadialGradient != null || shape.StrokeGradient != null ||
                shape.StrokeRadialGradient != null || shape.Shadow != null || shape.Glow != null ||
                shape.StrokeStartMarker != null || shape.StrokeEndMarker != null || shape.StrokeDashStyle != OfficeStrokeDashStyle.Solid ||
                shape.StrokeDashArray.Count != 0 || shape.StrokeDashOffset != 0)
                Fail("DRAWING_PAINT_UNSUPPORTED", "Gradients, effects, markers and dashed strokes are outside the portable profile.", location);
            OfficeTransform local = shape.Transform ?? OfficeTransform.Identity;
            ValidateTransform(local, location);
            OfficeTransform transform = local
                .Then(OfficeTransform.Translate(positioned.X, positioned.Y)).Then(parent);
            ValidateTransform(transform, location);
            string? fill = shape.Kind == OfficeShapeKind.Line ? null : Paint(shape.FillColor, shape.FillOpacity, location);
            Coordinate(shape.StrokeWidth, location);
            if (shape.StrokeWidth < 0) Fail("DRAWING_STROKE_INVALID", "Stroke width cannot be negative.", location);
            string? stroke = shape.StrokeWidth == 0 ? null : Paint(shape.StrokeColor, shape.StrokeOpacity, location);
            if (fill == null && stroke == null) Fail("DRAWING_PAINT_MISSING", "A portable path requires visible fill or stroke.", location);
            OfficeStrokeLineCap cap = shape.StrokeLineCap ?? OfficeStrokeLineCap.Butt;
            OfficeStrokeLineJoin join = shape.StrokeLineJoin ?? OfficeStrokeLineJoin.Miter;
            if (!Enum.IsDefined(typeof(OfficeStrokeLineCap), cap) || !Enum.IsDefined(typeof(OfficeStrokeLineJoin), join) ||
                !Enum.IsDefined(typeof(OfficeFillRule), shape.FillRule))
                Fail("DRAWING_PAINT_INVALID", "Unknown stroke or fill rule.", location);
            if (stroke != null && join == OfficeStrokeLineJoin.Miter && shape.StrokeMiterLimit != 10)
                Fail("DRAWING_MITER_UNSUPPORTED", "The portable miter-join profile requires the PDF miter limit of ten.", location);
            double strokeWidth = shape.StrokeWidth * transform.M11;
            Coordinate(strokeWidth, location);
            if (stroke != null && strokeWidth > 144) Fail("DRAWING_STROKE_LIMIT", "Portable stroke width cannot exceed 144 points.", location);
            RequireCommands(shape.Kind == OfficeShapeKind.Path ? shape.PathCommands.Count :
                shape.Kind == OfficeShapeKind.Polygon ? shape.Points.Count + 1 : shape.Kind == OfficeShapeKind.Line ? shape.Points.Count : 0, location);
            IReadOnlyList<OfficePathCommand> source = ShapeCommands(shape, location);
            CountCommands(source.Count, location);
            var commands = new List<OfficePathCommand>(source.Count);
            OfficePoint current = default, first = default;
            bool started = false;
            foreach (OfficePathCommand command in source) {
                if (command.Kind == OfficePathCommandKind.Close) {
                    if (!started) Fail("DRAWING_PATH_INVALID", "A closed path needs a preceding move.", location);
                    commands.Add(OfficePathCommand.Close()); current = first; continue;
                }
                if (command.Kind != OfficePathCommandKind.MoveTo && !started)
                    Fail("DRAWING_PATH_INVALID", "A path must start with a move.", location);
                OfficePoint end = Point(command.Point, transform, location);
                switch (command.Kind) {
                    case OfficePathCommandKind.MoveTo: commands.Add(OfficePathCommand.MoveTo(end)); first = command.Point; started = true; break;
                    case OfficePathCommandKind.LineTo: commands.Add(OfficePathCommand.LineTo(end)); break;
                    case OfficePathCommandKind.CubicBezierTo:
                        commands.Add(OfficePathCommand.CubicBezierTo(Point(command.ControlPoint1, transform, location),
                            Point(command.ControlPoint2, transform, location), end)); break;
                    case OfficePathCommandKind.QuadraticBezierTo:
                        var one = new OfficePoint(current.X + (command.ControlPoint1.X - current.X) * 2 / 3,
                            current.Y + (command.ControlPoint1.Y - current.Y) * 2 / 3);
                        var two = new OfficePoint(command.Point.X + (command.ControlPoint1.X - command.Point.X) * 2 / 3,
                            command.Point.Y + (command.ControlPoint1.Y - command.Point.Y) * 2 / 3);
                        commands.Add(OfficePathCommand.CubicBezierTo(Point(one, transform, location), Point(two, transform, location), end)); break;
                    default: Fail("DRAWING_PATH_INVALID", "Unknown path command.", location); break;
                }
                current = command.Point;
            }
            if (!started) Fail("DRAWING_PATH_INVALID", "A portable path needs a move command.", location);
            var result = new Element { Commands = commands, Fill = fill, Stroke = stroke, StrokeWidth = strokeWidth,
                FillRule = shape.FillRule == OfficeFillRule.EvenOdd ? "evenodd" : "nonzero",
                LineCap = cap.ToString().ToLowerInvariant(), LineJoin = join.ToString().ToLowerInvariant() };
            if (shape.ClipPath == null) return result;
            CountElement(location);
            return new Element { Kind = "group", Clip = Clip(shape.ClipPath, 0, 0, transform, location), Children = new List<Element> { result } };
        }
    }

    private static string? Paint(OfficeColor? color, double? opacity, string location) {
        if (!color.HasValue) return null;
        if (color.Value.A != 255 || opacity.HasValue && opacity.Value != 1)
            Fail("DRAWING_ALPHA_UNSUPPORTED", "Portable paths and text require opaque colors.", location);
        return color.Value.R.ToString("X2", CultureInfo.InvariantCulture) + color.Value.G.ToString("X2", CultureInfo.InvariantCulture)
            + color.Value.B.ToString("X2", CultureInfo.InvariantCulture);
    }

    private static IReadOnlyList<OfficePathCommand> ShapeCommands(OfficeShape shape, string location) {
        if (shape.Kind == OfficeShapeKind.Path) return shape.PathCommands;
        if (shape.Kind == OfficeShapeKind.Ellipse)
            return OfficeShapePresets.CreateEllipsePath(shape.Width / 2, shape.Height / 2, shape.Width / 2, shape.Height / 2, true);
        var commands = new List<OfficePathCommand>();
        if (shape.Kind == OfficeShapeKind.Line || shape.Kind == OfficeShapeKind.Polygon) {
            foreach (OfficePoint point in shape.Points)
                commands.Add(commands.Count == 0 ? OfficePathCommand.MoveTo(point) : OfficePathCommand.LineTo(point));
            if (shape.Kind == OfficeShapeKind.Polygon) commands.Add(OfficePathCommand.Close());
            return commands;
        }
        if (shape.Kind != OfficeShapeKind.Rectangle && shape.Kind != OfficeShapeKind.RoundedRectangle)
            Fail("DRAWING_SHAPE_UNSUPPORTED", "Unknown drawing shape.", location);
        double w = Length(shape.Width, location), h = Length(shape.Height, location);
        double r = shape.Kind == OfficeShapeKind.RoundedRectangle ? shape.CornerRadius : 0;
        if (r < 0 || r > Math.Min(w, h) / 2 || double.IsNaN(r)) Fail("DRAWING_RADIUS_INVALID", "Rounded rectangle radius is invalid.", location);
        if (r == 0) return new[] { OfficePathCommand.MoveTo(0, 0), OfficePathCommand.LineTo(w, 0),
            OfficePathCommand.LineTo(w, h), OfficePathCommand.LineTo(0, h), OfficePathCommand.Close() };
        const double k = 0.5522847498307936;
        double c = r * k;
        return new[] { OfficePathCommand.MoveTo(r, 0), OfficePathCommand.LineTo(w-r, 0),
            OfficePathCommand.CubicBezierTo(w-r+c, 0, w, r-c, w, r), OfficePathCommand.LineTo(w, h-r),
            OfficePathCommand.CubicBezierTo(w, h-r+c, w-r+c, h, w-r, h), OfficePathCommand.LineTo(r, h),
            OfficePathCommand.CubicBezierTo(r-c, h, 0, h-r+c, 0, h-r), OfficePathCommand.LineTo(0, r),
            OfficePathCommand.CubicBezierTo(0, r-c, r-c, 0, r, 0), OfficePathCommand.Close() };
    }
}
