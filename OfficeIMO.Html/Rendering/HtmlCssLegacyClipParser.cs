using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

/// <summary>Resolves the legacy absolutely positioned border-box clip rectangle.</summary>
internal static class HtmlCssLegacyClipParser {
    internal static bool IsSupportedSyntax(string value) =>
        TryResolve(value, 100D, 100D, 16D, 16D, 100D, 100D, 100D, 100D, out _);

    internal static bool TryResolve(
        string value,
        double boxWidth,
        double boxHeight,
        double fontSize,
        double rootFontSize,
        double viewportWidth,
        double viewportHeight,
        double containerWidth,
        double containerHeight,
        out HtmlCssResolvedClipPath? resolved,
        double characterAdvance = double.NaN) {
        resolved = null;
        string normalized = value.Trim().ToLowerInvariant();
        if (normalized == "auto") return true;
        if (!normalized.StartsWith("rect(", StringComparison.Ordinal)
            || HtmlRenderCssValues.FindMatchingParenthesis(normalized, 4) != normalized.Length - 1) return false;

        string arguments = normalized.Substring(5, normalized.Length - 6);
        if (!HtmlRenderCssValues.TrySplitTopLevelCommas(arguments, 4, out IReadOnlyList<string> commaParts)) return false;
        // CSS2 requires commas and also permits the historical whitespace-only form.
        // A mixture is rejected instead of inventing missing or ambiguous edges.
        IReadOnlyList<string> edges = commaParts.Count == 1
            ? HtmlRenderCssValues.SplitWhitespace(arguments)
            : commaParts;
        if (edges.Count != 4) return false;
        var lengths = new double[4];
        for (int index = 0; index < lengths.Length; index++) {
            string edge = edges[index].Trim();
            if (edge == "auto") {
                lengths[index] = index == 1 ? boxWidth : index == 2 ? boxHeight : 0D;
            } else if (edge.IndexOf('%') >= 0
                || !HtmlRenderCssValues.TryLength(edge, double.NaN, fontSize, rootFontSize,
                    viewportWidth, viewportHeight, containerWidth, containerHeight, out lengths[index], characterAdvance)) {
                return false;
            }
        }

        double width = lengths[1] - lengths[3];
        double height = lengths[2] - lengths[0];
        if (!IsFinite(width) || !IsFinite(height)) return false;
        resolved = new HtmlCssResolvedClipPath(lengths[3], lengths[0],
            width <= 0D || height <= 0D ? OfficeClipPath.Empty() : OfficeClipPath.Rectangle(width, height));
        return true;
    }

    private static bool IsFinite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);
}
