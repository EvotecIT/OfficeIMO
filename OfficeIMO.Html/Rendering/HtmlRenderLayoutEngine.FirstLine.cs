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
            if (!TryFindFirstLetterBounds(run.Text, out int start, out int end)) continue;
            string prefix = run.Text.Substring(0, start);
            string firstLetter = run.Text.Substring(start, end - start);
            string suffix = run.Text.Substring(end);
            var replacement = new List<HtmlInlineRun>(3);
            if (prefix.Length > 0) replacement.Add(run.CloneText(prefix, prefix, run.Style));
            replacement.Add(run.CloneText(firstLetter, firstLetter, firstLetterStyle, isFirstLetter: true));
            if (suffix.Length > 0) replacement.Add(run.CloneText(suffix, suffix, run.Style));
            runs.RemoveAt(runIndex);
            runs.InsertRange(runIndex, replacement);
            return;
        }
    }

    private bool TryFindFirstLetterBounds(string text, out int start, out int end) {
        start = end = 0;
        TextElementEnumerator enumerator = StringInfo.GetTextElementEnumerator(text);
        int examined = 0;
        bool MoveNext() {
            if (!enumerator.MoveNext()) return false;
            ChargeLayoutOperation("first-letter scanning");
            if ((++examined & 0xFF) == 0) CheckCancellation();
            return true;
        }

        if (!MoveNext()) return false;
        while (string.IsNullOrWhiteSpace(enumerator.GetTextElement())) {
            if (!MoveNext()) return false;
        }
        start = enumerator.ElementIndex;
        while (IsFirstLetterPunctuation(enumerator.GetTextElement())) {
            if (!MoveNext()) {
                end = text.Length;
                return true;
            }
        }
        if (!MoveNext()) {
            end = text.Length;
            return true;
        }
        while (IsFirstLetterPunctuation(enumerator.GetTextElement())) {
            if (!MoveNext()) {
                end = text.Length;
                return true;
            }
        }
        end = enumerator.ElementIndex;
        return end > start;
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
                double remainingWidth = Math.Max(0D, width - lineWidth);
                if (tokenWidth > remainingWidth
                    && !IsWhitespaceToken(token)
                    && remainingWidth > 0.0001D
                    && TryResolveFirstLineTokenSplit(token, run.Style, tokenStyle, remainingWidth, out int split)) {
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
        ChargeLayoutOperations(token.Length, "first-line token search");
        string searchToken = GetFirstLineSearchToken(token, firstLineStyle, width);
        HyphenationToken hyphenation = PrepareHyphenationToken(token, token, layoutStyle);
        if (hyphenation.HasBreaks) {
            int[] candidates = hyphenation.PrimaryBreaks.Concat(hyphenation.SecondaryBreaks)
                .Where(point => point > 0 && point < searchToken.Length &&
                    point < hyphenation.SourceBoundaries.Count)
                .Distinct().OrderBy(point => point).ToArray();
            int point = FindLargestFittingFirstLineBreak(
                hyphenation.PaintText, candidates, layoutStyle.HyphenateCharacter, firstLineStyle, width);
            if (point > 0) split = hyphenation.SourceBoundaries[point];
            if (split > 0 && split < token.Length) return true;
        }

        IReadOnlyList<int> preferred = OfficeTextLineBreaks.GetBreakPositions(
            searchToken,
            allowCjkBreaks: layoutStyle.WordBreak != "keep-all");
        split = FindLargestFittingFirstLineBreak(searchToken, preferred, string.Empty, firstLineStyle, width);
        if (split > 0) return true;
        if (!AllowsEmergencyTokenBreak(layoutStyle)) return false;

        int[] graphemeStarts = StringInfo.ParseCombiningCharacters(searchToken);
        split = FindLargestFittingFirstLineBreak(searchToken, graphemeStarts, string.Empty, firstLineStyle, width);
        return split > 0 && split < token.Length;
    }

    private string GetFirstLineSearchToken(string token, HtmlRenderBoxStyle style, double width) {
        const int initialSearchCharacters = 16384;
        if (style.LetterSpacing < 0D) return token;
        if (token.Length <= initialSearchCharacters) return token;

        int target = initialSearchCharacters;
        var elements = StringInfo.GetTextElementEnumerator(token);
        while (elements.MoveNext()) {
            int boundary = elements.ElementIndex;
            if (boundary < target) continue;
            CheckCancellation();
            ChargeLayoutOperations((boundary + 31L) / 32L, "first-line token search");
            if (MeasureInlineText(token.Substring(0, boundary), style) > width + 0.0001D) {
                return token.Substring(0, boundary);
            }
            target = target > token.Length / 2 ? token.Length : target * 2;
        }
        return token;
    }

    private int FindLargestFittingFirstLineBreak(
        string text,
        IReadOnlyList<int> points,
        string suffix,
        HtmlRenderBoxStyle style,
        double width) {
        if (style.LetterSpacing < 0D) {
            for (int index = points.Count - 1; index >= 0; index--) {
                CheckCancellation();
                int point = points[index];
                if (point <= 0 || point >= text.Length) continue;
                ChargeLayoutOperations((long)point + suffix.Length, "first-line token splitting");
                if (MeasureInlineText(text.Substring(0, point) + suffix, style) <= width + 0.0001D) {
                    return point;
                }
            }
            return -1;
        }
        int low = 0;
        int high = points.Count - 1;
        int best = -1;
        while (low <= high) {
            CheckCancellation();
            ChargeLayoutOperation("first-line token splitting");
            int middle = low + (high - low) / 2;
            int point = points[middle];
            if (point <= 0) {
                low = middle + 1;
                continue;
            }
            if (point >= text.Length) {
                high = middle - 1;
                continue;
            }
            ChargeLayoutOperations((point + 31L) / 32L, "first-line token splitting");
            string candidate = text.Substring(0, point) + suffix;
            if (MeasureInlineText(candidate, style) <= width + 0.0001D) {
                best = point;
                low = middle + 1;
            } else {
                high = middle - 1;
            }
        }
        return best;
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
