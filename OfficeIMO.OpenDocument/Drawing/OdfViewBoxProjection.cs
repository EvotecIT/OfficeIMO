using OfficeIMO.Drawing;

namespace OfficeIMO.OpenDocument;

/// <summary>Normalizes native view-box coordinates before scaling or mirroring to a drawing canvas.</summary>
internal static class OdfViewBoxProjection {
    internal static IEnumerable<OfficePathCommand> Project(IEnumerable<OfficePathCommand> commands, OdfViewBox box,
        double width, double height, bool mirrorX = false, bool mirrorY = false) {
        foreach (OfficePathCommand command in commands) {
            yield return command.Kind switch {
                OfficePathCommandKind.MoveTo => OfficePathCommand.MoveTo(Point(command.Point)),
                OfficePathCommandKind.LineTo => OfficePathCommand.LineTo(Point(command.Point)),
                OfficePathCommandKind.QuadraticBezierTo => OfficePathCommand.QuadraticBezierTo(Point(command.ControlPoint1), Point(command.Point)),
                OfficePathCommandKind.CubicBezierTo => OfficePathCommand.CubicBezierTo(Point(command.ControlPoint1), Point(command.ControlPoint2), Point(command.Point)),
                _ => command
            };
        }
        OfficePoint Point(OfficePoint point) {
            // Dividing first maps declared edges to exactly 0 and 1. Multiplying first can
            // round a valid edge beyond the canvas and make shared PDF export reject it.
            double x = (point.X - box.X) / box.Width, y = (point.Y - box.Y) / box.Height;
            return new OfficePoint((mirrorX ? 1 - x : x) * width, (mirrorY ? 1 - y : y) * height);
        }
    }
}
