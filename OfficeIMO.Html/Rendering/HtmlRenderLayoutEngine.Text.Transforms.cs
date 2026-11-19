using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    // Markers, generated and running strings draw small caps as full-size capitals.
    private string ApplyTextTransform(string text, HtmlRenderBoxStyle style) {
        if (style.ApproximateSmallCaps && ContainsSmallCapital(text, IsAllSmallCaps(style))) {
            ReportSyntheticSmallCaps(style, "generated-text", style.FontVariant);
        }
        return ApplyTextTransformSegments(new[] { text }, style)[0];
    }

    private void ApplyPendingInlineTextTransforms(IList<HtmlInlineRun> runs, bool retainSuppressedText = false, bool intrinsic = false) {
        int start = 0;
        while (start < runs.Count) {
            HtmlInlineRun first = runs[start];
            if (!first.TextTransformPending) {
                start++;
                continue;
            }

            if (!first.Style.ApproximateSmallCaps
                && string.Equals(first.Style.TextTransform, "none", StringComparison.OrdinalIgnoreCase)) {
                first.CompleteTextTransform(first.Text);
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
                textRuns.Select(static run => run.Text).ToArray(), first.Style, applySmallCaps: false,
                segmentStyles: textRuns.Select(static run => run.Style).ToArray());
            for (int index = 0; index < textRuns.Count; index++) {
                textRuns[index].CompleteTextTransform(transformed[index]);
            }
            start = end;
            if (first.Style.ApproximateSmallCaps) {
                start += ApplySmallCaps(runs, textRuns, intrinsic);
            }
        }
        if (!retainSuppressedText) RemoveSuppressedInlineText(runs);
    }

    // Synthetic small caps, as browsers draw them when a font has no small-cap glyphs: lowercase letters
    // become capitals at 70 % of the font size, while capitals, digits and punctuation keep their size.
    internal const double SyntheticSmallCapsScale = 0.7D;

    /// <summary>
    /// Draws small caps with the face's own <c>smcp</c> (and, for all-small-caps, <c>c2sc</c>) glyphs,
    /// keeping the source text and run intact. Faces without them fall back to synthesis.
    /// </summary>
    /// <remarks>
    /// Intrinsic sizing keeps one run per source run, so its synthesis uppercases in place.
    /// </remarks>
    private int ApplySmallCaps(IList<HtmlInlineRun> runs, IReadOnlyList<HtmlInlineRun> textRuns, bool intrinsic) {
        int inserted = 0;
        Dictionary<HtmlRenderBoxStyle, HtmlRenderBoxStyle>? fontStyles = null;
        foreach (HtmlInlineRun run in textRuns) {
            int index = runs.IndexOf(run);
            if (index < 0 || run.Text.Length == 0) continue;
            bool capitals = IsAllSmallCaps(run.Style);
            if (!ContainsSmallCapital(run.Text, capitals)) continue;
            if (CanUseFontSmallCaps(run.Text, run.Style, capitals)) {
                fontStyles ??= new Dictionary<HtmlRenderBoxStyle, HtmlRenderBoxStyle>();
                if (!fontStyles.TryGetValue(run.Style, out HtmlRenderBoxStyle? fontStyle)) {
                    fontStyle = run.Style.Clone();
                    fontStyle.TextFeatureSettings = OfficeOpenTypeSmallCaps.WithSmallCaps(run.Style.TextFeatureSettings, capitals);
                    fontStyles.Add(run.Style, fontStyle);
                }
                runs[index] = run.CloneText(run.Text, run.LogicalText, fontStyle, run.IsFirstLetter);
                continue;
            }
            if (!intrinsic) ReportSyntheticSmallCaps(run);
            inserted += SynthesizeSmallCaps(runs, index, run, capitals, split: !intrinsic);
        }
        return inserted;
    }

    private static bool IsAllSmallCaps(HtmlRenderBoxStyle style) =>
        style.FontVariant.IndexOf("all-small-caps", StringComparison.OrdinalIgnoreCase) >= 0
        || style.FontVariant.IndexOf("all-petite-caps", StringComparison.OrdinalIgnoreCase) >= 0;

    private static bool ContainsSmallCapital(string text, bool capitals) {
        foreach (char character in text) {
            if (OfficeOpenTypeSmallCaps.IsSmallCapital(character, capitals)) return true;
        }
        return false;
    }

    private bool CanUseFontSmallCaps(string text, HtmlRenderBoxStyle style, bool capitals) {
        // Managed shaping covers horizontal text; vertical runs keep the synthesized route.
        if (!string.Equals(style.WritingMode, "horizontal-tb", StringComparison.OrdinalIgnoreCase)) return false;
        foreach (OfficeFontFallbackRun fallback in _fonts.PlanFallbackRuns(text, style.Font.FamilyName, style.FontDescriptor)) {
            CheckCancellation();
            if (!ContainsSmallCapital(fallback.Text, capitals)) continue;
            if (!_fonts.TryResolveFaceForText(fallback.Text, fallback.FamilyName, style.FontDescriptor, style.Font.Size,
                    out OfficeFontFace? face) || face?.Program == null) return false;
            if (!OfficeOpenTypeSmallCaps.CanSubstitute(face.Program, fallback.Text, capitals, style.TextFeatureSettings,
                    _cancellationToken)) return false;
        }
        return true;
    }

    private void ReportSyntheticSmallCaps(HtmlInlineRun run) =>
        ReportSyntheticSmallCaps((object?)run.OwnerElement ?? run.Source,
            run.OwnerElement != null ? HtmlRenderStyleResolver.DescribeSource(run.OwnerElement) : run.Source,
            run.Style.FontVariant);

    private void ReportSyntheticSmallCaps(object key, string source, string fontVariantCaps) {
        if (!_reportedSyntheticSmallCaps.Add(key)) return;
        _diagnostics.Add(
            ComponentName,
            HtmlRenderDiagnosticCodes.FontVariantApproximated,
            "CSS small-caps were synthesized from capitals because the font has no small-cap glyphs for this text.",
            HtmlDiagnosticSeverity.Warning,
            source,
            "font-variant-caps=" + fontVariantCaps,
            OfficeConversionLossKind.Approximation);
    }

    private readonly HashSet<object> _reportedSyntheticSmallCaps = new HashSet<object>();

    // Synthesized pieces keep their source case as logical text, so extraction returns what the author wrote.
    private static int SynthesizeSmallCaps(IList<HtmlInlineRun> runs, int index, HtmlInlineRun run, bool capitals, bool split) {
        string text = run.Text;
        CultureInfo culture = ResolveTextTransformCulture(run.Style.Language);
        if (!split) {
            (string upperText, string logicalText) = Uppercase(text, culture);
            run.CompleteTextTransform(upperText, logicalText);
            return 0;
        }
        var pieces = new List<(string Text, bool Small)>();
        int pieceStart = 0;
        bool? small = null;
        for (int position = 0; position < text.Length; position++) {
            bool isSmall = OfficeOpenTypeSmallCaps.IsSmallCapital(text[position], capitals);
            if (small.HasValue && small.Value != isSmall) {
                pieces.Add((text.Substring(pieceStart, position - pieceStart), small.Value));
                pieceStart = position;
            }
            small = isSmall;
        }
        pieces.Add((text.Substring(pieceStart), small ?? false));

        int inserted = 0;
        HtmlRenderBoxStyle smallStyle = run.Style.Clone();
        smallStyle.Font = run.Style.Font.WithSize(run.Style.Font.Size * SyntheticSmallCapsScale);
        for (int piece = 0; piece < pieces.Count; piece++) {
            (string upper, string logical) = Uppercase(pieces[piece].Text, culture);
            if (piece == 0) {
                if (pieces[piece].Small) runs[index] = run.CloneText(upper, logical, smallStyle);
                else run.CompleteTextTransform(upper, logical);
                continue;
            }

            runs.Insert(index + piece, run.CloneText(upper, logical, pieces[piece].Small ? smallStyle : run.Style));
            inserted++;
        }
        return inserted;

        // Logical slices follow paint slices by length, so a case mapping that changes length keeps paint text.
        static (string Upper, string Logical) Uppercase(string value, CultureInfo culture) {
            string upper = OfficeTextCaseTransformer.ApplySegments(new[] { value }, OfficeTextCase.Uppercase, culture)[0];
            return (upper, upper.Length == value.Length ? value : upper);
        }
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
        bool applySmallCaps = true,
        IReadOnlyList<HtmlRenderBoxStyle>? segmentStyles = null) {
        CultureInfo culture = ResolveTextTransformCulture(style.Language);
        // text-transform combines one case keyword with full-width and full-size-kana.
        OfficeTextCase textCase = OfficeTextCase.None;
        bool fullWidth = false, fullSizeKana = false;
        foreach (string keyword in HtmlRenderCssValues.SplitWhitespace(style.TextTransform.ToLowerInvariant())) {
            switch (keyword) {
                case "uppercase": textCase = OfficeTextCase.Uppercase; break;
                case "lowercase": textCase = OfficeTextCase.Lowercase; break;
                case "capitalize": textCase = OfficeTextCase.Capitalize; break;
                case "full-width": fullWidth = true; break;
                case "full-size-kana": fullSizeKana = true; break;
            }
        }
        IReadOnlyList<string> transformed = string.Equals(style.TextTransform, "math-auto", StringComparison.OrdinalIgnoreCase)
            ? segments.Select(OfficeMathTextTransform.MathAuto).ToArray()
            : OfficeTextCaseTransformer.ApplySegments(segments, textCase, culture);
        if (fullWidth || fullSizeKana) {
            // Layout collapses, trims and breaks lines at ASCII spaces after this transform, so a
            // space becomes an ideographic space only where it does none of that: white space
            // that is preserved and does not wrap. Casing context spans the runs; white space is each run's own.
            transformed = transformed.Select((segment, index) => {
                HtmlRenderBoxStyle owner = segmentStyles != null && index < segmentStyles.Count ? segmentStyles[index] : style;
                bool fullWidthSpaces = owner.PreserveWhitespace && !owner.BreakSpaces && owner.PreventTextWrapping;
                return HtmlTextWidthTransform.Apply(segment, fullWidth, fullSizeKana, fullWidthSpaces);
            }).ToArray();
        }
        return applySmallCaps && style.ApproximateSmallCaps
            ? OfficeTextCaseTransformer.ApplySegments(transformed, OfficeTextCase.Uppercase, culture)
            : transformed;
    }

    private static CultureInfo ResolveTextTransformCulture(string language) {
        if (string.IsNullOrWhiteSpace(language)) return CultureInfo.InvariantCulture;
        if (TextTransformCultures.TryGetValue(language, out CultureInfo? cached)) return cached;
        CultureInfo culture;
        try {
            culture = CultureInfo.GetCultureInfo(language);
        } catch (CultureNotFoundException) {
            culture = CultureInfo.InvariantCulture;
        }
        if (TextTransformCultures.Count < MaximumTextTransformCultures) TextTransformCultures.TryAdd(language, culture);
        return culture;
    }

    // The framework takes a process-wide lock for each lookup, and this runs per text run.
    private const int MaximumTextTransformCultures = 256;
    private static readonly System.Collections.Concurrent.ConcurrentDictionary<string, CultureInfo> TextTransformCultures =
        new(StringComparer.Ordinal);

}
