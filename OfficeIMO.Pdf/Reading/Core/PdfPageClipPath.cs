using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal readonly struct PdfPageClipPath {
    private PdfPageClipPath(double x, double y, double width, double height, bool isRectangle, OfficeFillRule fillRule, IReadOnlyList<OfficePathCommand> commands) {
        X = x;
        Y = y;
        Width = width;
        Height = height;
        IsRectangle = isRectangle;
        FillRule = fillRule;
        Commands = commands;
    }

    public static PdfPageClipPath Rectangle(double x, double y, double width, double height) =>
        new PdfPageClipPath(x, y, width, height, true, OfficeFillRule.EvenOdd, Array.Empty<OfficePathCommand>());

    public static PdfPageClipPath ResolveActiveClip(PdfPageClipPath? activeClipPath, PdfPageClipPath clipPath) {
        if (!activeClipPath.HasValue) {
            return clipPath;
        }

        PdfPageClipPath active = activeClipPath.Value;
        if (!active.IsRectangle || !clipPath.IsRectangle) {
            if (active.IsRectangle) {
                return IntersectClipBounds(active, clipPath, out PdfPageClipPath intersection)
                    ? clipPath.WithBounds(intersection)
                    : Rectangle(Math.Max(active.X, clipPath.X), Math.Max(active.Y, clipPath.Y), 0D, 0D);
            }

            if (clipPath.IsRectangle) {
                return IntersectClipBounds(active, clipPath, out PdfPageClipPath intersection)
                    ? active.WithBounds(intersection)
                    : Rectangle(Math.Max(active.X, clipPath.X), Math.Max(active.Y, clipPath.Y), 0D, 0D);
            }

            return active;
        }

        return IntersectClipBounds(active, clipPath, out PdfPageClipPath rectangleIntersection)
            ? rectangleIntersection
            : Rectangle(Math.Max(active.X, clipPath.X), Math.Max(active.Y, clipPath.Y), 0D, 0D);
    }

    private static bool IntersectClipBounds(PdfPageClipPath first, PdfPageClipPath second, out PdfPageClipPath intersection) {
        double left = Math.Max(first.X, second.X);
        double top = Math.Max(first.Y, second.Y);
        double right = Math.Min(first.X + first.Width, second.X + second.Width);
        double bottom = Math.Min(first.Y + first.Height, second.Y + second.Height);
        double width = right - left;
        double height = bottom - top;
        if (width <= 0D || height <= 0D) {
            intersection = default;
            return false;
        }

        intersection = Rectangle(left, top, width, height);
        return true;
    }

    public static bool TryCreatePath(IReadOnlyList<OfficePathCommand> commands, OfficeFillRule fillRule, out PdfPageClipPath clipPath) {
        clipPath = default;
        if (commands.Count == 0 || commands[0].Kind != OfficePathCommandKind.MoveTo) {
            return false;
        }

        bool hasPoint = false;
        bool hasDraw = false;
        double left = 0D;
        double top = 0D;
        double right = 0D;
        double bottom = 0D;
        for (int i = 0; i < commands.Count; i++) {
            OfficePathCommand command = commands[i];
            switch (command.Kind) {
                case OfficePathCommandKind.MoveTo:
                    Include(command.Point, ref hasPoint, ref left, ref top, ref right, ref bottom);
                    break;
                case OfficePathCommandKind.LineTo:
                    Include(command.Point, ref hasPoint, ref left, ref top, ref right, ref bottom);
                    hasDraw = true;
                    break;
                case OfficePathCommandKind.QuadraticBezierTo:
                    Include(command.ControlPoint1, ref hasPoint, ref left, ref top, ref right, ref bottom);
                    Include(command.Point, ref hasPoint, ref left, ref top, ref right, ref bottom);
                    hasDraw = true;
                    break;
                case OfficePathCommandKind.CubicBezierTo:
                    Include(command.ControlPoint1, ref hasPoint, ref left, ref top, ref right, ref bottom);
                    Include(command.ControlPoint2, ref hasPoint, ref left, ref top, ref right, ref bottom);
                    Include(command.Point, ref hasPoint, ref left, ref top, ref right, ref bottom);
                    hasDraw = true;
                    break;
                case OfficePathCommandKind.Close:
                    break;
            }
        }

        double width = right - left;
        double height = bottom - top;
        if (!hasDraw || width <= 0D || height <= 0D) {
            return false;
        }

        clipPath = new PdfPageClipPath(left, top, width, height, false, fillRule, new List<OfficePathCommand>(commands));
        return true;
    }

    public double X { get; }

    public double Y { get; }

    public double Width { get; }

    public double Height { get; }

    public bool IsRectangle { get; }

    public OfficeFillRule FillRule { get; }

    public IReadOnlyList<OfficePathCommand> Commands { get; }

    private PdfPageClipPath WithBounds(PdfPageClipPath bounds) =>
        new PdfPageClipPath(bounds.X, bounds.Y, bounds.Width, bounds.Height, IsRectangle, FillRule, Commands);

    public OfficeClipPath? ToOfficeClipPath(double primitiveX, double primitiveY) {
        if (!NearlyEqual(X, primitiveX) || !NearlyEqual(Y, primitiveY) || Width <= 0D || Height <= 0D) {
            return null;
        }

        if (IsRectangle) {
            return OfficeClipPath.Rectangle(Width, Height);
        }

        try {
            return OfficeClipPath.Path(TranslateCommands(primitiveX, primitiveY), FillRule);
        } catch (ArgumentException) {
            return null;
        }
    }

    private List<OfficePathCommand> TranslateCommands(double offsetX, double offsetY) {
        var result = new List<OfficePathCommand>(Commands.Count);
        for (int i = 0; i < Commands.Count; i++) {
            OfficePathCommand command = Commands[i];
            switch (command.Kind) {
                case OfficePathCommandKind.MoveTo:
                    result.Add(OfficePathCommand.MoveTo(command.Point.X - offsetX, command.Point.Y - offsetY));
                    break;
                case OfficePathCommandKind.LineTo:
                    result.Add(OfficePathCommand.LineTo(command.Point.X - offsetX, command.Point.Y - offsetY));
                    break;
                case OfficePathCommandKind.QuadraticBezierTo:
                    result.Add(OfficePathCommand.QuadraticBezierTo(
                        command.ControlPoint1.X - offsetX,
                        command.ControlPoint1.Y - offsetY,
                        command.Point.X - offsetX,
                        command.Point.Y - offsetY));
                    break;
                case OfficePathCommandKind.CubicBezierTo:
                    result.Add(OfficePathCommand.CubicBezierTo(
                        command.ControlPoint1.X - offsetX,
                        command.ControlPoint1.Y - offsetY,
                        command.ControlPoint2.X - offsetX,
                        command.ControlPoint2.Y - offsetY,
                        command.Point.X - offsetX,
                        command.Point.Y - offsetY));
                    break;
                case OfficePathCommandKind.Close:
                    result.Add(OfficePathCommand.Close());
                    break;
            }
        }

        return result;
    }

    private static void Include(OfficePoint point, ref bool hasPoint, ref double left, ref double top, ref double right, ref double bottom) {
        if (!hasPoint) {
            left = right = point.X;
            top = bottom = point.Y;
            hasPoint = true;
            return;
        }

        if (point.X < left) {
            left = point.X;
        }

        if (point.Y < top) {
            top = point.Y;
        }

        if (point.X > right) {
            right = point.X;
        }

        if (point.Y > bottom) {
            bottom = point.Y;
        }
    }

    private static bool NearlyEqual(double left, double right) => Math.Abs(left - right) <= 0.001D;
}
