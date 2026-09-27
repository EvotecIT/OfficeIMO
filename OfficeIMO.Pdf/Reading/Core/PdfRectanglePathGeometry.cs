using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

// Shared rectangle evidence for retained page geometry and its consumers.
internal static class PdfRectanglePathGeometry {
    // Fills close an open rectangle implicitly; stroked candidates need the
    // final edge from Close or an explicit line back to the starting point.
    internal static bool IsRectangle(IReadOnlyList<OfficePathCommand> commands, bool allowImplicitClose, double tolerance = 0.000001D) {
        if (commands.Count < 4 || commands.Count > 6 ||
            commands[0].Kind != OfficePathCommandKind.MoveTo ||
            commands[1].Kind != OfficePathCommandKind.LineTo ||
            commands[2].Kind != OfficePathCommandKind.LineTo ||
            commands[3].Kind != OfficePathCommandKind.LineTo ||
            !IsAxisAlignedRectangle(commands, tolerance)) return false;
        if (commands.Count == 4) return allowImplicitClose;
        if (commands.Count == 5 && commands[4].Kind == OfficePathCommandKind.Close) return true;
        return commands[4].Kind == OfficePathCommandKind.LineTo &&
            Math.Abs(commands[4].Point.X - commands[0].Point.X) <= tolerance &&
            Math.Abs(commands[4].Point.Y - commands[0].Point.Y) <= tolerance &&
            (commands.Count == 5 || commands[5].Kind == OfficePathCommandKind.Close);
    }

    private static bool IsAxisAlignedRectangle(IReadOnlyList<OfficePathCommand> commands, double tolerance) {
        OfficePoint p0 = commands[0].Point;
        OfficePoint p1 = commands[1].Point;
        OfficePoint p2 = commands[2].Point;
        OfficePoint p3 = commands[3].Point;
        return (Math.Abs(p0.X - p1.X) <= tolerance && Math.Abs(p1.Y - p2.Y) <= tolerance &&
                Math.Abs(p2.X - p3.X) <= tolerance && Math.Abs(p3.Y - p0.Y) <= tolerance ||
                Math.Abs(p0.Y - p1.Y) <= tolerance && Math.Abs(p1.X - p2.X) <= tolerance &&
                Math.Abs(p2.Y - p3.Y) <= tolerance && Math.Abs(p3.X - p0.X) <= tolerance) &&
            Math.Abs(p0.X - p2.X) > tolerance && Math.Abs(p0.Y - p2.Y) > tolerance;
    }

}
