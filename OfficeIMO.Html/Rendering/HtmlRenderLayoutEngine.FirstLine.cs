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
                if (!hasContent
                    && tokenWidth > width
                    && !IsWhitespaceToken(token)
                    && TryResolveFirstLineTokenSplit(token, run.Style, tokenStyle, width, out int split)) {
                    string prefix = token.Substring(0, split);
                    string suffix = token.Substring(split);
                    styledRuns.Add(run.CloneText(prefix, prefix, tokenStyle, run.IsFirstLetter));
                    if (suffix.Length > 0) styledRuns.Add(run.CloneText(suffix, suffix, run.Style, run.IsFirstLetter));
                    firstLine = false;
                    hasContent = true;
                    continue;
                }
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

    private bool TryResolveFirstLineTokenSplit(
        string token,
        HtmlRenderBoxStyle layoutStyle,
        HtmlRenderBoxStyle firstLineStyle,
        double width,
        out int split) {
        split = -1;
        HyphenationToken hyphenation = PrepareHyphenationToken(token, token, layoutStyle);
        if (hyphenation.HasBreaks) {
            foreach (int point in hyphenation.PrimaryBreaks.Concat(hyphenation.SecondaryBreaks).Distinct().OrderBy(point => point)) {
                if (point <= 0 || point >= hyphenation.LogicalText.Length || point >= hyphenation.SourceBoundaries.Count) continue;
                string candidate = hyphenation.PaintText.Substring(0, point) + layoutStyle.HyphenateCharacter;
                if (MeasureInlineText(candidate, firstLineStyle) <= width + 0.0001D) {
                    split = hyphenation.SourceBoundaries[point];
                }
            }
            if (split > 0 && split < token.Length) return true;
        }

        IReadOnlyList<int> preferred = OfficeTextLineBreaks.GetBreakPositions(
            token,
            allowCjkBreaks: layoutStyle.WordBreak != "keep-all");
        foreach (int point in preferred) {
            if (point <= 0 || point >= token.Length) continue;
            if (MeasureInlineText(token.Substring(0, point), firstLineStyle) <= width + 0.0001D) split = point;
        }
        if (split > 0) return true;
        if (!AllowsEmergencyTokenBreak(layoutStyle)) return false;

        int sourceLength = 0;
        foreach (string element in OfficeTextElements.Enumerate(token)) {
            int candidateLength = sourceLength + element.Length;
            if (candidateLength >= token.Length
                || MeasureInlineText(token.Substring(0, candidateLength), firstLineStyle) > width + 0.0001D) break;
            sourceLength = candidateLength;
        }
        split = sourceLength;
        return split > 0 && split < token.Length;
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
