using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal readonly struct HtmlResolvedBorderRadii {
    internal HtmlResolvedBorderRadii(
        double topLeftX,
        double topLeftY,
        double topRightX,
        double topRightY,
        double bottomRightX,
        double bottomRightY,
        double bottomLeftX,
        double bottomLeftY) {
        TopLeftX = Math.Max(0D, topLeftX);
        TopLeftY = Math.Max(0D, topLeftY);
        TopRightX = Math.Max(0D, topRightX);
        TopRightY = Math.Max(0D, topRightY);
        BottomRightX = Math.Max(0D, bottomRightX);
        BottomRightY = Math.Max(0D, bottomRightY);
        BottomLeftX = Math.Max(0D, bottomLeftX);
        BottomLeftY = Math.Max(0D, bottomLeftY);
    }

    internal double TopLeftX { get; }
    internal double TopLeftY { get; }
    internal double TopRightX { get; }
    internal double TopRightY { get; }
    internal double BottomRightX { get; }
    internal double BottomRightY { get; }
    internal double BottomLeftX { get; }
    internal double BottomLeftY { get; }

    internal bool IsZero => MaximumRadius <= 0.0001D;

    internal bool IsUniformCircular =>
        NearlyEqual(TopLeftX, TopLeftY)
        && NearlyEqual(TopLeftX, TopRightX)
        && NearlyEqual(TopLeftX, TopRightY)
        && NearlyEqual(TopLeftX, BottomRightX)
        && NearlyEqual(TopLeftX, BottomRightY)
        && NearlyEqual(TopLeftX, BottomLeftX)
        && NearlyEqual(TopLeftX, BottomLeftY);

    internal double UniformRadius => IsUniformCircular ? TopLeftX : 0D;

    private double MaximumRadius => Math.Max(
        Math.Max(Math.Max(TopLeftX, TopLeftY), Math.Max(TopRightX, TopRightY)),
        Math.Max(Math.Max(BottomRightX, BottomRightY), Math.Max(BottomLeftX, BottomLeftY)));

    internal HtmlResolvedBorderRadii Normalize(double width, double height) {
        double scale = 1D;
        scale = LimitScale(scale, width, TopLeftX + TopRightX);
        scale = LimitScale(scale, width, BottomLeftX + BottomRightX);
        scale = LimitScale(scale, height, TopLeftY + BottomLeftY);
        scale = LimitScale(scale, height, TopRightY + BottomRightY);
        return scale >= 0.999999D ? this : Scale(scale);
    }

    internal HtmlResolvedBorderRadii Inset(
        double left,
        double top,
        double right,
        double bottom,
        double targetWidth,
        double targetHeight) => new HtmlResolvedBorderRadii(
            Math.Max(0D, TopLeftX - left),
            Math.Max(0D, TopLeftY - top),
            Math.Max(0D, TopRightX - right),
            Math.Max(0D, TopRightY - top),
            Math.Max(0D, BottomRightX - right),
            Math.Max(0D, BottomRightY - bottom),
            Math.Max(0D, BottomLeftX - left),
            Math.Max(0D, BottomLeftY - bottom)).Normalize(targetWidth, targetHeight);

    internal HtmlResolvedBorderRadii Expand(double amount, double targetWidth, double targetHeight) =>
        new HtmlResolvedBorderRadii(
            Math.Max(0D, TopLeftX + amount),
            Math.Max(0D, TopLeftY + amount),
            Math.Max(0D, TopRightX + amount),
            Math.Max(0D, TopRightY + amount),
            Math.Max(0D, BottomRightX + amount),
            Math.Max(0D, BottomRightY + amount),
            Math.Max(0D, BottomLeftX + amount),
            Math.Max(0D, BottomLeftY + amount)).Normalize(targetWidth, targetHeight);

    internal IReadOnlyList<OfficePathCommand> CreatePathCommands(double width, double height) {
        const double kappa = 0.5522847498307936D;
        return new[] {
            OfficePathCommand.MoveTo(TopLeftX, 0D),
            OfficePathCommand.LineTo(width - TopRightX, 0D),
            OfficePathCommand.CubicBezierTo(width - TopRightX + TopRightX * kappa, 0D, width, TopRightY - TopRightY * kappa, width, TopRightY),
            OfficePathCommand.LineTo(width, height - BottomRightY),
            OfficePathCommand.CubicBezierTo(width, height - BottomRightY + BottomRightY * kappa, width - BottomRightX + BottomRightX * kappa, height, width - BottomRightX, height),
            OfficePathCommand.LineTo(BottomLeftX, height),
            OfficePathCommand.CubicBezierTo(BottomLeftX - BottomLeftX * kappa, height, 0D, height - BottomLeftY + BottomLeftY * kappa, 0D, height - BottomLeftY),
            OfficePathCommand.LineTo(0D, TopLeftY),
            OfficePathCommand.CubicBezierTo(0D, TopLeftY - TopLeftY * kappa, TopLeftX - TopLeftX * kappa, 0D, TopLeftX, 0D),
            OfficePathCommand.Close()
        };
    }

    private HtmlResolvedBorderRadii Scale(double scale) => new HtmlResolvedBorderRadii(
        TopLeftX * scale,
        TopLeftY * scale,
        TopRightX * scale,
        TopRightY * scale,
        BottomRightX * scale,
        BottomRightY * scale,
        BottomLeftX * scale,
        BottomLeftY * scale);

    private static double LimitScale(double current, double available, double required) =>
        required <= 0.0001D ? current : Math.Min(current, Math.Max(0D, available) / required);

    private static bool NearlyEqual(double left, double right) => Math.Abs(left - right) <= 0.0001D;
}
