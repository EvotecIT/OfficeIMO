using AngleSharp.Dom;
using System.Globalization;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private void ApplyFirstLetterStyle(
        IElement? formattingContainer,
        double width,
        HtmlRenderBoxStyle parentStyle,
        List<HtmlInlineRun> runs) {
        if (formattingContainer == null
            || !_styleResolver.TryResolvePseudo(
                formattingContainer,
                HtmlPseudoElementKind.FirstLetter,
                width,
                parentStyle,
                out HtmlRenderBoxStyle firstLetterStyle)) return;

        for (int runIndex = 0; runIndex < runs.Count; runIndex++) {
            HtmlInlineRun run = runs[runIndex];
            if (run.AtomicBlock != null || run.Text.Length == 0) continue;
            IReadOnlyList<string> elements = OfficeTextElements.Split(run.Text);
            int start = 0;
            while (start < elements.Count && string.IsNullOrWhiteSpace(elements[start])) start++;
            if (start >= elements.Count) continue;

            int end = start;
            while (end < elements.Count && IsFirstLetterPunctuation(elements[end])) end++;
            if (end < elements.Count) end++;
            while (end < elements.Count && IsFirstLetterPunctuation(elements[end])) end++;
            if (end <= start) return;

            string prefix = string.Concat(elements.Take(start));
            string firstLetter = string.Concat(elements.Skip(start).Take(end - start));
            string suffix = string.Concat(elements.Skip(end));
            var replacement = new List<HtmlInlineRun>(3);
            if (prefix.Length > 0) replacement.Add(run.CloneText(prefix, prefix, run.Style));
            replacement.Add(run.CloneText(firstLetter, firstLetter, firstLetterStyle, isFirstLetter: true));
            if (suffix.Length > 0) replacement.Add(run.CloneText(suffix, suffix, run.Style));
            runs.RemoveAt(runIndex);
            runs.InsertRange(runIndex, replacement);
            return;
        }
    }

    private void ApplyFirstLineStyle(
        IElement? formattingContainer,
        double width,
        HtmlRenderBoxStyle parentStyle,
        List<HtmlInlineRun> runs) {
        if (formattingContainer == null
            || !_styleResolver.TryResolvePseudo(
                formattingContainer,
                HtmlPseudoElementKind.FirstLine,
                width,
                parentStyle,
                out HtmlRenderBoxStyle firstLineStyle)) return;

        var styledRuns = new List<HtmlInlineRun>(runs.Count);
        double lineWidth = 0D;
        bool firstLine = true;
        bool hasContent = false;
        foreach (HtmlInlineRun run in runs) {
            if (!firstLine || run.AtomicBlock != null || run.Text.Length == 0) {
                styledRuns.Add(run);
                if (firstLine && run.AtomicBlock != null) {
                    if (hasContent && lineWidth + run.AtomicBlock.Width > width) firstLine = false;
                    else {
                        lineWidth += run.AtomicBlock.Width;
                        hasContent = true;
                    }
                }
                continue;
            }

            IReadOnlyList<string> tokens = Tokenize(run.Text, run.Style.PreserveWhitespace, run.Style.BreakSpaces).ToList();
            for (int tokenIndex = 0; tokenIndex < tokens.Count; tokenIndex++) {
                string token = tokens[tokenIndex];
                if (token == "\u2028" || run.Style.PreserveWhitespace && (token == "\n" || token == "\r\n")) {
                    styledRuns.Add(run.CloneText(token, token, run.IsFirstLetter ? run.Style : firstLineStyle, run.IsFirstLetter));
                    firstLine = false;
                    continue;
                }

                HtmlRenderBoxStyle tokenStyle = run.IsFirstLetter ? run.Style : firstLineStyle;
                string measuredToken = !run.Style.PreserveWhitespace && IsWhitespaceToken(token) ? " " : token;
                double tokenWidth = MeasureInlineText(measuredToken, tokenStyle);
                if (hasContent && lineWidth + tokenWidth > width) firstLine = false;
                HtmlRenderBoxStyle appliedStyle = firstLine ? tokenStyle : run.Style;
                styledRuns.Add(run.CloneText(token, token, appliedStyle, run.IsFirstLetter));
                if (firstLine) {
                    lineWidth += tokenWidth;
                    if (!IsWhitespaceToken(token)) hasContent = true;
                }
            }
        }
        runs.Clear();
        runs.AddRange(styledRuns);
    }

    private static bool IsFirstLetterPunctuation(string textElement) {
        if (string.IsNullOrEmpty(textElement)) return false;
        UnicodeCategory category = char.GetUnicodeCategory(textElement, 0);
        return category == UnicodeCategory.OpenPunctuation
            || category == UnicodeCategory.ClosePunctuation
            || category == UnicodeCategory.InitialQuotePunctuation
            || category == UnicodeCategory.FinalQuotePunctuation
            || category == UnicodeCategory.OtherPunctuation;
    }
}
