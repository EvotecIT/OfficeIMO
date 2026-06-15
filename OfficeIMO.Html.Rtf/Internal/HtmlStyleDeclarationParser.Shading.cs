using System.Globalization;

namespace OfficeIMO.Html.Rtf;

internal static partial class HtmlStyleDeclarationParser {
    private static int? ParseRtfShadingInteger(string value) {
        return int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsed) && parsed >= 0
            ? parsed
            : null;
    }

    private static int? ParseRtfShadingPercent(string value) {
        if (value.EndsWith("%", StringComparison.Ordinal) &&
            double.TryParse(value.Substring(0, value.Length - 1), NumberStyles.Float, CultureInfo.InvariantCulture, out double percent) &&
            percent >= 0d) {
            return (int)Math.Round(Math.Min(percent, 100d) * 100d, MidpointRounding.AwayFromZero);
        }

        return int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsed) && parsed >= 0
            ? Math.Min(parsed, 10000)
            : null;
    }

    private static RtfShadingPattern? ParseRtfShadingPattern(string value) {
        switch (NormalizeRtfShadingPattern(value)) {
            case "horizontal":
            case "horiz":
            case "bghoriz":
            case "trbghoriz":
            case "clbghoriz":
                return RtfShadingPattern.Horizontal;
            case "vertical":
            case "vert":
            case "bgvert":
            case "trbgvert":
            case "clbgvert":
                return RtfShadingPattern.Vertical;
            case "forwarddiagonal":
            case "fdiag":
            case "bgfdiag":
            case "trbgfdiag":
            case "clbgfdiag":
                return RtfShadingPattern.ForwardDiagonal;
            case "backwarddiagonal":
            case "bdiag":
            case "bgbdiag":
            case "trbgbdiag":
            case "clbgbdiag":
                return RtfShadingPattern.BackwardDiagonal;
            case "cross":
            case "bgcross":
            case "trbgcross":
            case "clbgcross":
                return RtfShadingPattern.Cross;
            case "diagonalcross":
            case "dcross":
            case "bgdcross":
            case "trbgdcross":
            case "clbgdcross":
                return RtfShadingPattern.DiagonalCross;
            case "darkhorizontal":
            case "darkhoriz":
            case "dkhorizontal":
            case "dkhoriz":
            case "dkhor":
            case "bgdkhoriz":
            case "bgdkhor":
            case "trbgdkhor":
            case "clbgdkhor":
                return RtfShadingPattern.DarkHorizontal;
            case "darkvertical":
            case "darkvert":
            case "dkvertical":
            case "dkvert":
            case "bgdkvert":
            case "trbgdkvert":
            case "clbgdkvert":
                return RtfShadingPattern.DarkVertical;
            case "darkforwarddiagonal":
            case "darkfdiag":
            case "dkforwarddiagonal":
            case "dkfdiag":
            case "bgdkfdiag":
            case "trbgdkfdiag":
            case "clbgdkfdiag":
                return RtfShadingPattern.DarkForwardDiagonal;
            case "darkbackwarddiagonal":
            case "darkbdiag":
            case "dkbackwarddiagonal":
            case "dkbdiag":
            case "bgdkbdiag":
            case "trbgdkbdiag":
            case "clbgdkbdiag":
                return RtfShadingPattern.DarkBackwardDiagonal;
            case "darkcross":
            case "dkcross":
            case "bgdkcross":
            case "trbgdkcross":
            case "clbgdkcross":
                return RtfShadingPattern.DarkCross;
            case "darkdiagonalcross":
            case "darkdcross":
            case "dkdiagonalcross":
            case "dkdcross":
            case "bgdkdcross":
            case "trbgdkdcross":
            case "clbgdkdcross":
                return RtfShadingPattern.DarkDiagonalCross;
            case "none":
                return RtfShadingPattern.None;
            default:
                return null;
        }
    }

    private static string NormalizeRtfShadingPattern(string value) {
        string token = value.Trim().TrimStart('\\').ToLowerInvariant();
        return token.Replace("-", string.Empty).Replace("_", string.Empty);
    }
}
