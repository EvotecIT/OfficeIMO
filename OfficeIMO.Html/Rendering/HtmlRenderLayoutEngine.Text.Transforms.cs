using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private static string ApplyTextTransform(string text, HtmlRenderBoxStyle style) =>
        ApplyTextTransformSegments(new[] { text }, style)[0];

    private static void ApplyPendingInlineTextTransforms(IList<HtmlInlineRun> runs, bool retainSuppressedText = false) {
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
                textRuns.Select(static run => run.Text).ToArray(), first.Style, applySmallCaps: false);
            for (int index = 0; index < textRuns.Count; index++) {
                textRuns[index].CompleteTextTransform(transformed[index]);
            }
            start = end;
            if (first.Style.ApproximateSmallCaps) {
                start += SynthesizeSmallCaps(runs, textRuns);
            }
        }
        if (!retainSuppressedText) RemoveSuppressedInlineText(runs);
    }

    // Synthetic small caps, as browsers draw them when a font has no small-cap glyphs: lowercase letters
    // become capitals at 70 % of the font size, while capitals, digits and punctuation keep their size.
    internal const double SyntheticSmallCapsScale = 0.7D;

    private static int SynthesizeSmallCaps(IList<HtmlInlineRun> runs, IReadOnlyList<HtmlInlineRun> textRuns) {
        int inserted = 0;
        foreach (HtmlInlineRun run in textRuns) {
            string text = run.Text;
            int index = runs.IndexOf(run);
            if (index < 0 || text.Length == 0) continue;
            CultureInfo culture = ResolveTextTransformCulture(run.Style.Language);
            var pieces = new List<(string Text, bool Small)>();
            int pieceStart = 0;
            bool? small = null;
            for (int position = 0; position < text.Length; position++) {
                bool isSmall = char.IsLower(text[position]);
                if (small.HasValue && small.Value != isSmall) {
                    pieces.Add((text.Substring(pieceStart, position - pieceStart), small.Value));
                    pieceStart = position;
                }
                small = isSmall;
            }
            pieces.Add((text.Substring(pieceStart), small ?? false));
            if (pieces.Count == 1 && !pieces[0].Small) continue;

            HtmlRenderBoxStyle smallStyle = run.Style.Clone();
            smallStyle.Font = run.Style.Font.WithSize(run.Style.Font.Size * SyntheticSmallCapsScale);
            for (int piece = 0; piece < pieces.Count; piece++) {
                string upper = OfficeTextCaseTransformer.ApplySegments(new[] { pieces[piece].Text }, OfficeTextCase.Uppercase, culture)[0];
                if (piece == 0) {
                    if (pieces[piece].Small) {
                        HtmlInlineRun replacement = run.CloneText(upper, upper, smallStyle);
                        replacement.CompleteTextTransform(upper);
                        runs[index] = replacement;
                    } else {
                        run.CompleteTextTransform(upper);
                    }
                    continue;
                }

                HtmlInlineRun clone = run.CloneText(upper, upper, pieces[piece].Small ? smallStyle : run.Style);
                clone.CompleteTextTransform(upper);
                runs.Insert(index + piece, clone);
                inserted++;
            }
        }

        return inserted;
    }

    // A font-changing pseudo-element may restore the selected source fragment.
    // All other zero-font input must leave before shaping, painting, or extraction.
    // Its CSS whitespace can still permit a zero-advance line break.
    private static void RemoveSuppressedInlineText(IList<HtmlInlineRun> runs) {
        for (int index = runs.Count - 1; index >= 0; index--) {
            HtmlInlineRun run = runs[index];
            if (!run.IsTextTransformContextOnly) continue;
            if (run.Text.Any(IsCollapsibleCssWhitespace) && !run.Style.PreventTextWrapping) {
                HtmlInlineRun marker = run.CloneText(string.Empty, string.Empty, run.Style);
                marker.IsTextTransformContextOnly = false;
                marker.IsInlineStrutMarker = true;
                marker.IsSuppressedWhitespaceBreak = true;
                runs[index] = marker;
            } else runs.RemoveAt(index);
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
        HtmlRenderBoxStyle style,
        bool applySmallCaps = true) {
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
        return applySmallCaps && style.ApproximateSmallCaps
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
