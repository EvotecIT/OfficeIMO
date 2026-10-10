using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private static string ApplyTextTransform(string text, HtmlRenderBoxStyle style) =>
        ApplyTextTransformSegments(new[] { text }, style)[0];

    private static void ApplyPendingInlineTextTransforms(IList<HtmlInlineRun> runs) {
        int start = 0;
        while (start < runs.Count) {
            HtmlInlineRun first = runs[start];
            if (!first.TextTransformPending) {
                start++;
                continue;
            }

            var textRuns = new List<HtmlInlineRun> { first };
            int end = start + 1;
            while (end < runs.Count) {
                HtmlInlineRun next = runs[end];
                if (IsInlineTextTransformContextMarker(next)) {
                    end++;
                    continue;
                }
                if (!next.TextTransformPending || !HasEquivalentTextTransform(first.Style, next.Style)) break;
                textRuns.Add(next);
                end++;
            }

            IReadOnlyList<string> transformed = ApplyTextTransformSegments(
                textRuns.Select(static run => run.Text).ToArray(), first.Style);
            for (int index = 0; index < textRuns.Count; index++) {
                textRuns[index].CompleteTextTransform(transformed[index]);
            }
            start = end;
        }
        // Remove suppressed logical input before first-letter/first-line styling,
        // shaping, layout, semantic assignment, or PDF text serialization.
        for (int index = runs.Count - 1; index >= 0; index--) {
            if (runs[index].IsTextTransformContextOnly) runs.RemoveAt(index);
        }
    }

    // Layout-only markers have no logical text. Retain their runs and styles for
    // layout, but do not let them reset contextual casing of adjacent text.
    private static bool IsInlineTextTransformContextMarker(HtmlInlineRun run) =>
        !run.TextTransformPending
        && (run.IsInlineStrutMarker
            || run.InlineEdgeBoundary != null
            || run.IsFlowMarker
            || run.RunningStringElement != null
            || run.RunningElementAssignment != null
            || run.PositionedMarkerElement != null
            || (run.Text.Length > 0
                && run.Text.All(static character => OfficeTextElements.ContainsBidiControl(character.ToString()))));

    private static bool HasEquivalentTextTransform(HtmlRenderBoxStyle left, HtmlRenderBoxStyle right) =>
        string.Equals(left.TextTransform, right.TextTransform, StringComparison.OrdinalIgnoreCase)
        && string.Equals(left.Language, right.Language, StringComparison.OrdinalIgnoreCase)
        && left.ApproximateSmallCaps == right.ApproximateSmallCaps;

    private static IReadOnlyList<string> ApplyTextTransformSegments(
        IReadOnlyList<string> segments,
        HtmlRenderBoxStyle style) {
        CultureInfo culture = ResolveTextTransformCulture(style.Language);
        OfficeTextCase textCase = style.TextTransform.ToLowerInvariant() switch {
            "uppercase" => OfficeTextCase.Uppercase,
            "lowercase" => OfficeTextCase.Lowercase,
            "capitalize" => OfficeTextCase.Capitalize,
            _ => OfficeTextCase.None
        };
        IReadOnlyList<string> transformed = string.Equals(style.TextTransform, "math-auto", StringComparison.OrdinalIgnoreCase)
            ? segments.Select(OfficeMathTextTransform.MathAuto).ToArray()
            : OfficeTextCaseTransformer.ApplySegments(segments, textCase, culture);
        return style.ApproximateSmallCaps
            ? OfficeTextCaseTransformer.ApplySegments(transformed, OfficeTextCase.Uppercase, culture)
            : transformed;
    }

    private static CultureInfo ResolveTextTransformCulture(string language) {
        if (string.IsNullOrWhiteSpace(language)) return CultureInfo.InvariantCulture;
        try {
            return CultureInfo.GetCultureInfo(language);
        } catch (CultureNotFoundException) {
            return CultureInfo.InvariantCulture;
        }
    }

}
