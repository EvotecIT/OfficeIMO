using System.Globalization;

namespace OfficeIMO.Html;

internal static class HtmlCssReplacedElementParser {
    private const double MinimumAspectRatio = 0.000001D;
    private const double MaximumAspectRatio = 1000000D;
    private const double MaximumPositionScalar = 1000000000D;
    internal static string NormalizeObjectFit(string value, out string unsupported) {
        unsupported = string.Empty;
        string normalized = string.IsNullOrWhiteSpace(value) ? "fill" : value.Trim().ToLowerInvariant();
        if (normalized == "fill" || normalized == "contain" || normalized == "cover" || normalized == "none" || normalized == "scale-down") {
            return normalized;
        }
        unsupported = "object-fit=" + normalized;
        return "fill";
    }

    internal static string NormalizeObjectPosition(string value, double fontSize, double rootFontSize, out string unsupported) {
        unsupported = string.Empty;
        string normalized = string.IsNullOrWhiteSpace(value) ? "50% 50%" : value.Trim().ToLowerInvariant();
        if (TryParsePosition(normalized, fontSize, rootFontSize, out _, out _)) return normalized;
        unsupported = "object-position=" + normalized;
        return "50% 50%";
    }

    internal static bool TryParseAspectRatio(
        string value,
        out double? ratio,
        out bool prefersIntrinsic,
        out string unsupported) {
        ratio = null;
        prefersIntrinsic = true;
        unsupported = string.Empty;
        string normalized = string.IsNullOrWhiteSpace(value) ? "auto" : value.Trim().ToLowerInvariant();
        IReadOnlyList<string> tokens = HtmlRenderCssValues.SplitWhitespace(normalized);
        if (tokens.Count == 1 && tokens[0] == "auto") return true;

        var ratioTokens = new List<string>();
        bool autoSeen = false;
        foreach (string token in tokens) {
            if (token == "auto") {
                if (autoSeen) return InvalidAspect(normalized, out unsupported);
                autoSeen = true;
            } else {
                ratioTokens.Add(token);
            }
        }
        if (ratioTokens.Count == 0 || !TryPositiveRatio(string.Concat(ratioTokens), out double parsed)) {
            return InvalidAspect(normalized, out unsupported);
        }
        ratio = parsed;
        prefersIntrinsic = autoSeen;
        return true;
    }

    internal static bool TryResolveObjectPosition(
        string value,
        double areaWidth,
        double areaHeight,
        double objectWidth,
        double objectHeight,
        double fontSize,
        double rootFontSize,
        out double offsetX,
        out double offsetY) {
        offsetX = 0D;
        offsetY = 0D;
        if (!TryParsePosition(value, fontSize, rootFontSize, out AxisPosition horizontal, out AxisPosition vertical)) return false;
        offsetX = horizontal.Resolve(areaWidth, areaWidth - objectWidth, fontSize, rootFontSize);
        offsetY = vertical.Resolve(areaHeight, areaHeight - objectHeight, fontSize, rootFontSize);
        return !double.IsNaN(offsetX) && !double.IsInfinity(offsetX)
            && !double.IsNaN(offsetY) && !double.IsInfinity(offsetY);
    }

    internal static bool IsSupportedObjectFitSyntax(string value) {
        NormalizeObjectFit(value, out string unsupported);
        return unsupported.Length == 0;
    }

    internal static bool IsSupportedObjectPositionSyntax(string value) =>
        TryParsePosition(value.Trim().ToLowerInvariant(), 16D, 16D, out _, out _);

    internal static bool IsSupportedAspectRatioSyntax(string value) =>
        TryParseAspectRatio(value, out _, out _, out _);

    private static bool InvalidAspect(string normalized, out string unsupported) {
        unsupported = "aspect-ratio=" + normalized;
        return false;
    }

    private static bool TryPositiveRatio(string value, out double ratio) {
        ratio = 0D;
        string[] parts = value.Split('/');
        if (parts.Length < 1 || parts.Length > 2 || !TryPositiveNumber(parts[0], out double numerator)) return false;
        double denominator = 1D;
        if (parts.Length == 2 && !TryPositiveNumber(parts[1], out denominator)) return false;
        ratio = numerator / denominator;
        return ratio >= MinimumAspectRatio && ratio <= MaximumAspectRatio && !double.IsNaN(ratio) && !double.IsInfinity(ratio);
    }

    private static bool TryPositiveNumber(string value, out double number) =>
        double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out number)
        && number > 0D
        && !double.IsNaN(number)
        && !double.IsInfinity(number);

    private static bool TryParsePosition(
        string value,
        double fontSize,
        double rootFontSize,
        out AxisPosition horizontal,
        out AxisPosition vertical) {
        horizontal = AxisPosition.Center;
        vertical = AxisPosition.Center;
        IReadOnlyList<string> tokens = HtmlRenderCssValues.SplitWhitespace(value);
        if (tokens.Count < 1 || tokens.Count > 4) return false;
        if (tokens.Count == 1) return TryParseOnePosition(tokens[0], fontSize, rootFontSize, out horizontal, out vertical);
        if (tokens.Count == 2) return TryParseTwoPositions(tokens[0], tokens[1], fontSize, rootFontSize, out horizontal, out vertical);
        return TryParseEdgeOffsetPositions(tokens, fontSize, rootFontSize, out horizontal, out vertical);
    }

    private static bool TryParseOnePosition(
        string token,
        double fontSize,
        double rootFontSize,
        out AxisPosition horizontal,
        out AxisPosition vertical) {
        horizontal = AxisPosition.Center;
        vertical = AxisPosition.Center;
        if (IsVerticalEdge(token)) return TryParseAxis(token, horizontalAxis: false, fontSize, rootFontSize, out vertical);
        return TryParseAxis(token, horizontalAxis: true, fontSize, rootFontSize, out horizontal);
    }

    private static bool TryParseTwoPositions(
        string first,
        string second,
        double fontSize,
        double rootFontSize,
        out AxisPosition horizontal,
        out AxisPosition vertical) {
        horizontal = AxisPosition.Center;
        vertical = AxisPosition.Center;
        bool verticalFirst = IsVerticalEdge(first) || IsHorizontalEdge(second);
        string horizontalToken = verticalFirst ? second : first;
        string verticalToken = verticalFirst ? first : second;
        return TryParseAxis(horizontalToken, horizontalAxis: true, fontSize, rootFontSize, out horizontal)
            && TryParseAxis(verticalToken, horizontalAxis: false, fontSize, rootFontSize, out vertical);
    }

    private static bool TryParseEdgeOffsetPositions(
        IReadOnlyList<string> tokens,
        double fontSize,
        double rootFontSize,
        out AxisPosition horizontal,
        out AxisPosition vertical) {
        horizontal = AxisPosition.Center;
        vertical = AxisPosition.Center;
        bool horizontalSet = false;
        bool verticalSet = false;
        bool offsetSeen = false;
        for (int index = 0; index < tokens.Count; index++) {
            string token = tokens[index];
            if (token == "center") {
                if (!horizontalSet) {
                    horizontal = AxisPosition.Center;
                    horizontalSet = true;
                } else if (!verticalSet) {
                    vertical = AxisPosition.Center;
                    verticalSet = true;
                } else {
                    return false;
                }
                continue;
            }

            bool horizontalAxis = IsHorizontalEdge(token);
            if (!horizontalAxis && !IsVerticalEdge(token)) return false;
            if (horizontalAxis ? horizontalSet : verticalSet) return false;
            string? offset = null;
            if (index + 1 < tokens.Count && IsLengthPercentage(tokens[index + 1], fontSize, rootFontSize)) {
                offset = tokens[++index];
                offsetSeen = true;
            }
            AxisPosition position = AxisPosition.Edge(token == "right" || token == "bottom", offset);
            if (horizontalAxis) {
                horizontal = position;
                horizontalSet = true;
            } else {
                vertical = position;
                verticalSet = true;
            }
        }
        return offsetSeen && horizontalSet && verticalSet;
    }

    private static bool TryParseAxis(
        string token,
        bool horizontalAxis,
        double fontSize,
        double rootFontSize,
        out AxisPosition position) {
        position = AxisPosition.Center;
        if (token == "center") return true;
        if (horizontalAxis && IsHorizontalEdge(token)) {
            position = AxisPosition.Edge(token == "right", null);
            return true;
        }
        if (!horizontalAxis && IsVerticalEdge(token)) {
            position = AxisPosition.Edge(token == "bottom", null);
            return true;
        }
        if (TryPercentage(token, out double percentage)) {
            position = AxisPosition.Aligned(percentage);
            return true;
        }
        if (!IsLength(token, fontSize, rootFontSize)) return false;
        position = AxisPosition.Edge(end: false, token);
        return true;
    }

    private static bool IsLengthPercentage(string value, double fontSize, double rootFontSize) =>
        TryPercentage(value, out _) || IsLength(value, fontSize, rootFontSize);

    private static bool IsLength(string value, double fontSize, double rootFontSize) {
        return !value.EndsWith("%", StringComparison.Ordinal)
            && HtmlRenderCssValues.TryLength(value, 100D, fontSize, rootFontSize, out double length)
            && Math.Abs(length) <= MaximumPositionScalar;
    }

    private static bool TryPercentage(string value, out double percentage) {
        percentage = 0D;
        if (!value.EndsWith("%", StringComparison.Ordinal)) return false;
        string number = value.Substring(0, value.Length - 1);
        if (!double.TryParse(number, NumberStyles.Float, CultureInfo.InvariantCulture, out percentage)
            || double.IsNaN(percentage)
            || double.IsInfinity(percentage)
            || Math.Abs(percentage) > MaximumPositionScalar) return false;
        percentage /= 100D;
        return true;
    }

    private static bool IsHorizontalEdge(string token) => token == "left" || token == "right";
    private static bool IsVerticalEdge(string token) => token == "top" || token == "bottom";

    private readonly struct AxisPosition {
        private AxisPosition(double alignment, bool end, string? edgeOffset, bool usesAlignment) {
            Alignment = alignment;
            End = end;
            EdgeOffset = edgeOffset;
            UsesAlignment = usesAlignment;
        }

        internal static AxisPosition Center => Aligned(0.5D);
        internal static AxisPosition Aligned(double alignment) => new AxisPosition(alignment, false, null, true);
        internal static AxisPosition Edge(bool end, string? offset) => new AxisPosition(0D, end, offset, false);

        private double Alignment { get; }
        private bool End { get; }
        private string? EdgeOffset { get; }
        private bool UsesAlignment { get; }

        internal double Resolve(double areaLength, double freeSpace, double fontSize, double rootFontSize) {
            if (UsesAlignment) return freeSpace * Alignment;
            double offset = 0D;
            if (EdgeOffset != null) {
                if (TryPercentage(EdgeOffset, out double percentage)) offset = areaLength * percentage;
                else HtmlRenderCssValues.TryLength(EdgeOffset, areaLength, fontSize, rootFontSize, out offset);
            }
            return End ? freeSpace - offset : offset;
        }
    }
}
