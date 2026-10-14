using AngleSharp.Dom;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    // Prepare the same contextual casing and first-letter fragment used by final
    // inline layout before whitespace normalization or intrinsic measurement.
    // Paragraph markers carry the originating box; inline descendants retain it.
    private IReadOnlyList<IntrinsicTextRun> PrepareIntrinsicFirstLetterRuns(IReadOnlyList<IntrinsicTextRun> rawRuns) {
        if (rawRuns.Count == 0) return rawRuns;
        var transformRuns = new List<HtmlInlineRun>(rawRuns.Count);
        foreach (IntrinsicTextRun run in rawRuns) {
            transformRuns.Add(new HtmlInlineRun(run.Text, run.Style, null, "intrinsic-text",
                textTransformPending: run.TextTransformPending && !run.IsParagraphStart && !run.IsForcedBreak && !run.IsReplaced) {
                IsInlineStrutMarker = run.IsInlineInset
            });
        }
        ApplyPendingInlineTextTransforms(transformRuns, retainSuppressedText: true);

        var prepared = new List<IntrinsicTextRun>(rawRuns.Count);
        IElement? firstLetterOwner = null;
        double firstLetterWidth = 0D;
        for (int index = 0; index < rawRuns.Count; index++) {
            IntrinsicTextRun run = rawRuns[index];
            if (run.IsParagraphStart) {
                firstLetterWidth = run.FirstLetterContainingWidth;
                firstLetterOwner = run.FirstLetterOwner != null
                    && TryResolveFirstLetterStyle(run.FirstLetterOwner, firstLetterWidth, run.Style, out _)
                    ? run.FirstLetterOwner
                    : null;
                prepared.Add(run);
                continue;
            }
            if (run.IsInlineInset) {
                prepared.Add(run);
                continue;
            }
            if (run.IsForcedBreak || run.IsReplaced) {
                firstLetterOwner = null;
                prepared.Add(run);
                continue;
            }

            string text = transformRuns[index].Text;
            bool firstLineBlocked = false;
            if (firstLetterOwner == null
                || !TryFindFirstLetterBounds(text, run.Style.PreserveWhitespace, out int start, out int end,
                    out firstLineBlocked)) {
                if (firstLineBlocked) firstLetterOwner = null;
                prepared.Add(new IntrinsicTextRun(text, run.Style, textTransformPending: false));
                continue;
            }
            if (TryResolveFirstLetterStyle(firstLetterOwner, firstLetterWidth, run.Style, out HtmlRenderBoxStyle letterStyle)) {
                if (start > 0) prepared.Add(new IntrinsicTextRun(text.Substring(0, start), run.Style, textTransformPending: false));
                prepared.Add(new IntrinsicTextRun(text.Substring(start, end - start), letterStyle, textTransformPending: false));
                if (end < text.Length) prepared.Add(new IntrinsicTextRun(text.Substring(end), run.Style, textTransformPending: false));
            } else {
                prepared.Add(new IntrinsicTextRun(text, run.Style, textTransformPending: false));
            }
            // A genuine zero-font first letter still owns selection. Do not style
            // a later visible descendant when this fragment remains suppressed.
            firstLetterOwner = null;
        }
        return prepared;
    }
}
