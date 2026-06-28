using System.Text;

namespace OfficeIMO.Drawing;

/// <summary>
/// Shared helpers for mapping Office stroke dash styles to renderer dash patterns.
/// </summary>
public static class OfficeStrokeDashStyleExtensions {
    /// <summary>
    /// Returns a dash/gap pattern in rendered units for the supplied stroke width.
    /// </summary>
    /// <param name="style">Stroke dash style.</param>
    /// <param name="strokeWidth">Rendered stroke width.</param>
    /// <returns>Alternating dash and gap lengths. An empty array means a solid stroke.</returns>
    public static double[] GetDashPattern(this OfficeStrokeDashStyle style, double strokeWidth) {
        double width = strokeWidth <= 0D ? 1D : strokeWidth;
        switch (style) {
            case OfficeStrokeDashStyle.Dash:
                return new[] { width * 4D, width * 2D };
            case OfficeStrokeDashStyle.Dot:
                return new[] { width, width * 2D };
            case OfficeStrokeDashStyle.DashDot:
                return new[] { width * 4D, width * 2D, width, width * 2D };
            case OfficeStrokeDashStyle.DashDotDot:
                return new[] { width * 4D, width * 2D, width, width * 2D, width, width * 2D };
            default:
                return System.Array.Empty<double>();
        }
    }

    /// <summary>
    /// Returns an SVG <c>stroke-dasharray</c> value for the supplied stroke width.
    /// </summary>
    /// <param name="style">Stroke dash style.</param>
    /// <param name="strokeWidth">Rendered stroke width.</param>
    /// <returns>SVG dash-array value, or <c>null</c> for solid strokes.</returns>
    public static string? GetSvgDashArray(this OfficeStrokeDashStyle style, double strokeWidth) {
        double[] pattern = style.GetDashPattern(strokeWidth);
        if (pattern.Length == 0) {
            return null;
        }

        var builder = new StringBuilder();
        for (int i = 0; i < pattern.Length; i++) {
            if (i > 0) {
                builder.Append(' ');
            }

            builder.Append(OfficeSvgFormatting.FormatNumber(pattern[i]));
        }

        return builder.ToString();
    }
}
