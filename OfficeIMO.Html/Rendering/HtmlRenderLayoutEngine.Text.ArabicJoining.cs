using System.Text;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private int _arabicJoiningEligibilityRemainder;

    // Core owns joining; this projection only retains context across paint-only
    // inline boundaries and maps the length-preserving result to source slices.
    private IReadOnlyList<HtmlInlineRun> PrepareArabicJoiningRuns(IReadOnlyList<HtmlInlineRun> runs) {
        if (!HasArabicJoiningSource(runs, static run => run.LogicalText)) return runs;
        CheckCancellation();
        ChargeLayoutOperations(runs.Count, "Arabic joining source mapping");
        var items = new ArabicJoiningItem[runs.Count];
        for (int index = 0; index < runs.Count; index++) {
            CheckCancellation();
            HtmlInlineRun run = runs[index];
            bool barrier = run.AtomicBlock != null && !run.IsFlowMarker
                || run.FloatingBlock != null || run.PositionedMarkerElement != null
                || run.RunningStringElement != null || run.RunningElementAssignment != null
                || run.IsEmptyInlineBox || run.LeaderPattern != null || run.IsSuppressedWhitespaceBreak
                || OfficeTextElements.ContainsBidiControl(run.LogicalText)
                || run.InlineEdgeBoundary != null && HasArabicJoiningInsets(run.Style);
            items[index] = new ArabicJoiningItem(run.LogicalText, run.Style, barrier, run.EndsFirstLine);
        }
        string[] contextual = PrepareArabicJoiningText(items);
        var result = new List<HtmlInlineRun>(runs.Count);
        for (int index = 0; index < runs.Count; index++) {
            CheckCancellation();
            HtmlInlineRun run = runs[index];
            result.Add(string.Equals(contextual[index], run.Text, StringComparison.Ordinal)
                ? run : run.CloneText(contextual[index], run.LogicalText, run.Style, run.IsFirstLetter));
        }
        return result;
    }

    private IReadOnlyList<IntrinsicTextRun> PrepareArabicJoiningIntrinsicRuns(IReadOnlyList<IntrinsicTextRun> runs) {
        if (!HasArabicJoiningSource(runs, static run => run.Text)) return runs;
        CheckCancellation();
        ChargeLayoutOperations(runs.Count, "Arabic joining source mapping");
        var items = new ArabicJoiningItem[runs.Count];
        for (int index = 0; index < runs.Count; index++) {
            CheckCancellation();
            IntrinsicTextRun run = runs[index];
            items[index] = new ArabicJoiningItem(run.Text, run.Style,
                run.IsParagraphStart || run.IsForcedBreak || run.IsReplaced
                    || OfficeTextElements.ContainsBidiControl(run.Text), endsContext: false);
        }
        string[] contextual = PrepareArabicJoiningText(items);
        var result = new List<IntrinsicTextRun>(runs.Count);
        for (int index = 0; index < runs.Count; index++) {
            CheckCancellation();
            IntrinsicTextRun run = runs[index];
            result.Add(string.Equals(contextual[index], run.Text, StringComparison.Ordinal)
                ? run : new IntrinsicTextRun(contextual[index], run.Style, textTransformPending: false));
        }
        return PrepareConfiguredJoiningIntrinsicRuns(result);
    }

    // Eligibility only reads run/text fields. Carry partial chunks across calls:
    // nested intrinsic lists must not each incur a rounded-up minimum operation.
    // Full context mapping retains its separate per-run and character charges.
    private bool HasArabicJoiningSource<T>(IReadOnlyList<T> runs, Func<T, string> textSelector) {
        CheckCancellation();
        for (int index = 0; index < runs.Count; index++) {
            CheckCancellation();
            string text = textSelector(runs[index]);
            ChargeArabicJoiningEligibility(1L + text.Length);
            if (OfficeTextElements.ContainsJoiningScript(text)) return true;
        }
        return false;
    }

    private void ChargeArabicJoiningEligibility(long work) {
        long pending = work + _arabicJoiningEligibilityRemainder;
        long operations = pending / 32L;
        _arabicJoiningEligibilityRemainder = (int)(pending % 32L);
        if (operations > 0L) ChargeLayoutOperations(operations, "Arabic joining eligibility");
    }

    private string[] PrepareArabicJoiningText(IReadOnlyList<ArabicJoiningItem> items) {
        CheckCancellation();
        var result = new string[items.Count];
        for (int index = 0; index < items.Count; index++) {
            CheckCancellation();
            result[index] = items[index].Text;
        }
        int start = 0;
        while (start < items.Count) {
            CheckCancellation();
            if (items[start].Text.Length == 0) { start++; continue; }
            int end = start + 1;
            int lastText = start;
            if (!items[start].Barrier) {
                while (end < items.Count && !items[end].Barrier && !items[lastText].EndsContext) {
                    CheckCancellation();
                    if (items[end].Text.Length == 0) { end++; continue; }
                    if (!CanShareArabicJoiningContext(items[lastText].Style, items[end].Style)) break;
                    lastText = end++;
                }
            }
            int length = 0;
            bool hasJoining = false;
            for (int index = start; index < end; index++) {
                CheckCancellation();
                ChargeLayoutOperations((items[index].Text.Length + 31L) / 32L, "Arabic joining text scan");
                length = checked(length + items[index].Text.Length);
                hasJoining |= OfficeTextElements.ContainsJoiningScript(items[index].Text);
            }
            if (hasJoining) {
                ChargeLayoutOperations(1L + (length + 31L) / 32L, "Arabic joining context");
                var context = new StringBuilder(length);
                for (int index = start; index < end; index++) {
                    CheckCancellation();
                    context.Append(items[index].Text);
                }
                CheckCancellation();
                string shaped = OfficeArabicTextShaper.Shape(context.ToString());
                CheckCancellation();
                if (shaped.Length != length) throw new InvalidOperationException("Arabic joining changed source slice lengths.");
                int offset = 0;
                for (int index = start; index < end; index++) {
                    CheckCancellation();
                    result[index] = shaped.Substring(offset, items[index].Text.Length);
                    offset += items[index].Text.Length;
                }
            }
            start = end;
        }
        return result;
    }

    private bool CanShareArabicJoiningContext(HtmlRenderBoxStyle left, HtmlRenderBoxStyle right) =>
        _options.TextShapingProvider == null
        && left.TextFeatureSettings.IsDefault && right.TextFeatureSettings.IsDefault
        && left.Font.Equals(right.Font) && left.FontDescriptor == right.FontDescriptor
        && string.Equals(left.Language, right.Language, StringComparison.OrdinalIgnoreCase)
        && left.Direction == right.Direction && left.UnicodeBidi == "normal" && right.UnicodeBidi == "normal"
        && left.WritingMode == "horizontal-tb" && right.WritingMode == "horizontal-tb"
        && left.LetterSpacing == 0D && right.LetterSpacing == 0D
        && left.WordSpacing.Equals(right.WordSpacing)
        && left.WordBreak == "normal" && right.WordBreak == "normal"
        && left.OverflowWrap == "normal" && right.OverflowWrap == "normal"
        && !HasArabicJoiningInsets(left) && !HasArabicJoiningInsets(right);

    private static bool HasArabicJoiningInsets(HtmlRenderBoxStyle style) =>
        style.HorizontalInsets != 0D || style.MarginLeft != 0D || style.MarginRight != 0D;

    private readonly struct ArabicJoiningItem {
        internal ArabicJoiningItem(string text, HtmlRenderBoxStyle style, bool barrier, bool endsContext) {
            Text = text; Style = style; Barrier = barrier; EndsContext = endsContext;
        }
        internal string Text { get; }
        internal HtmlRenderBoxStyle Style { get; }
        internal bool Barrier { get; }
        internal bool EndsContext { get; }
    }

    private List<HtmlInlineRun> ApplyScopedFontFallbacks(IEnumerable<HtmlInlineRun> sourceRuns) {
        IReadOnlyList<HtmlInlineRun> materialized = sourceRuns as IReadOnlyList<HtmlInlineRun> ?? sourceRuns.ToArray();
        var resolvedRuns = new List<HtmlInlineRun>();
        foreach (HtmlInlineRun run in PrepareArabicJoiningRuns(materialized)) {
            if (run.Text.Length == 0 || run.AtomicBlock != null || run.FloatingBlock != null || run.PositionedMarkerElement != null) {
                resolvedRuns.Add(run);
                continue;
            }
            _fontUsage?.Observe(run.LogicalText, run.Style.Font.FamilyName, run.Style.FontDescriptor);
            IReadOnlyList<OfficeFontFallbackRun> fallbacks = _fonts.PlanFallbackRuns(
                run.LogicalText, run.Style.Font.FamilyName, run.Style.FontDescriptor);
            if (fallbacks.Count == 1
                && string.Equals(fallbacks[0].Text, run.LogicalText, StringComparison.Ordinal)
                && string.Equals(InlineFallbackFamilyRequest(fallbacks[0]), run.Style.Font.FamilyName, StringComparison.Ordinal)) {
                resolvedRuns.Add(run);
                continue;
            }
            int fallbackOffset = 0;
            foreach (OfficeFontFallbackRun fallback in fallbacks) {
                CheckCancellation();
                HtmlRenderBoxStyle style = run.Style.Clone();
                style.Font = style.Font.WithFamilyName(InlineFallbackFamilyRequest(fallback));
                HtmlInlineRun resolvedRun = run.CloneText(
                    run.Text.Substring(fallbackOffset, fallback.Text.Length), fallback.Text, style, run.IsFirstLetter);
                resolvedRun.PreparedHyphenation = run.PreparedHyphenation?.SliceSource(fallbackOffset, fallback.Text.Length);
                fallbackOffset += fallback.Text.Length;
                resolvedRun.EndsFirstLine = run.EndsFirstLine && fallbackOffset == run.Text.Length;
                resolvedRun.FirstLineHyphen = resolvedRun.EndsFirstLine ? run.FirstLineHyphen : string.Empty;
                resolvedRuns.Add(resolvedRun);
            }
        }
        return resolvedRuns.Count == 0 ? resolvedRuns : PrepareConfiguredJoiningRuns(resolvedRuns, resolvedRuns[0].Style);
    }
}
