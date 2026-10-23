using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private IReadOnlyList<IntrinsicTextRun> PrepareConfiguredJoiningIntrinsicRuns(IReadOnlyList<IntrinsicTextRun> runs) {
        if (runs.Count == 0) return runs;
        var sources = new List<HtmlInlineRun>(runs.Count);
        foreach (IntrinsicTextRun run in runs) {
            bool boundary = run.IsParagraphStart || run.IsForcedBreak || run.IsReplaced;
            sources.Add(new HtmlInlineRun(boundary ? "\u2028" : run.Text, run.Style, null, "intrinsic-text",
                logicalText: boundary ? "\u2028" : OfficeArabicTextShaper.ToLogicalText(run.Text)));
        }
        Dictionary<HtmlInlineRun, List<PreparedInlineSlice>> prepared = PrepareConfiguredJoiningSlices(sources, runs[0].Style, reportLoss: false);
        var result = new List<IntrinsicTextRun>(runs.Count);
        for (int index = 0; index < runs.Count; index++) {
            CheckCancellation();
            IntrinsicTextRun run = runs[index];
            if (!prepared.TryGetValue(sources[index], out List<PreparedInlineSlice>? slices)) { result.Add(run); continue; }
            int offset = 0;
            foreach (PreparedInlineSlice slice in slices) {
                if (slice.Offset > offset) result.Add(new IntrinsicTextRun(run.Text.Substring(offset, slice.Offset - offset), run.Style, textTransformPending: false));
                result.Add(new IntrinsicTextRun(slice.Shaping?.Text ?? run.Text.Substring(slice.Offset, slice.Length), run.Style, textTransformPending: false) {
                    PreparedShaping = slice.Shaping,
                    PreserveContextualFallback = slice.Shaping == null
                });
                offset = slice.Offset + slice.Length;
            }
            if (offset < run.Text.Length) result.Add(new IntrinsicTextRun(run.Text.Substring(offset), run.Style, textTransformPending: false));
        }
        return result;
    }
}
