using AngleSharp.Dom;
using System.Globalization;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
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

        width = Math.Max(0.01D, width - (parentStyle.TextIndent?.Resolve(width) ?? 0D));
        var styledRuns = new List<HtmlInlineRun>(runs.Count);
        PrepareInlineEdgeScopes(runs);
        var previewLine = new InlineLine();
        bool firstLine = true;
        bool hasContent = false;
        for (int runIndex = 0; runIndex < runs.Count; runIndex++) {
            HtmlInlineRun run = runs[runIndex];
            if (run.InlineEdgeBoundary != null) {
                if (firstLine) ProcessInlineEdgeBoundary(run, previewLine);
                styledRuns.Add(run);
                continue;
            }
            if (!firstLine || run.AtomicBlock != null || run.Text.Length == 0) {
                styledRuns.Add(run);
                if (firstLine && run.AtomicBlock != null) {
                    if (hasContent && previewLine.PreviewAdvance(run, run.AtomicBlock.Width) > width) firstLine = false;
                    else {
                        previewLine.Add(new InlineSegment(string.Empty, run.AtomicBlock.Width, run));
                        hasContent = true;
                    }
                }
                continue;
            }

            HtmlRenderBoxStyle runFirstLineStyle = firstLineStyle.Clone();
            // Language comes from the originating element, not a CSS pseudo-element.
            runFirstLineStyle.Language = run.Style.Language;
            // ::first-line changes eligible font/paint properties. It must not
            // replace the originating run's whitespace and tab-stop contract.
            runFirstLineStyle.PreserveWhitespace = run.Style.PreserveWhitespace;
            runFirstLineStyle.BreakSpaces = run.Style.BreakSpaces;
            runFirstLineStyle.PreventTextWrapping = run.Style.PreventTextWrapping;
            runFirstLineStyle.TabSize = run.Style.TabSize;
            runFirstLineStyle.TabSizeIsLength = run.Style.TabSizeIsLength;
            IReadOnlyList<string> tokens = Tokenize(run.Text, run.Style.PreserveWhitespace, run.Style.BreakSpaces).ToList();
            for (int tokenIndex = 0; tokenIndex < tokens.Count; tokenIndex++) {
                string token = tokens[tokenIndex];
                run.InlineTokenEndsRun = tokenIndex == tokens.Count - 1;
                if (!firstLine) {
                    styledRuns.Add(run.CloneText(token, token, run.Style, run.IsFirstLetter));
                    continue;
                }
                if (token == "\u2028" || run.Style.PreserveWhitespace && (token == "\n" || token == "\r\n")) {
                    styledRuns.Add(run.CloneText(token, token, run.IsFirstLetter ? run.Style : runFirstLineStyle, run.IsFirstLetter));
                    firstLine = false;
                    continue;
                }

                HtmlRenderBoxStyle tokenStyle = run.IsFirstLetter ? run.Style : runFirstLineStyle;
                string measuredToken = !run.Style.PreserveWhitespace && IsWhitespaceToken(token) ? " " : token;
                double tokenWidth = tokenStyle.PreserveWhitespace && measuredToken.IndexOf('\t') >= 0
                    ? MeasureTabExpandedText(measuredToken, tokenStyle, previewLine.PreviewContentStart(run))
                    : MeasureInlineText(measuredToken, tokenStyle);
                double remainingWidth = Math.Max(0D, width - previewLine.PreviewAdvance(run, 0D, false));
                bool preventWrapping = parentStyle.PreventTextWrapping || run.Style.PreventTextWrapping;
                bool moveWholeToken = !preventWrapping && hasContent && !IsWhitespaceToken(token)
                    && tokenWidth > remainingWidth + 0.0001D
                    && (tokenStyle.HyphenateLimitZone > 0D && remainingWidth <= tokenStyle.HyphenateLimitZone + 0.0001D
                        || tokenStyle.HyphenateLimitLast == "always"
                            && !HasRemainingInlineFlowContent(runs, runIndex, tokens, tokenIndex)
                            && MeasureInlineText(measuredToken, run.Style) <= width + 0.0001D);
                if (moveWholeToken) {
                    firstLine = false;
                    styledRuns.Add(run.CloneText(token, token, run.Style, run.IsFirstLetter));
                    continue;
                }
                if (!preventWrapping
                    && previewLine.PreviewAdvance(run, tokenWidth) > width
                    && !IsWhitespaceToken(token)
                    && remainingWidth > 0.0001D
                    && TryResolveFirstLineTokenSplit(token, run.Style, tokenStyle, remainingWidth, out int split, out HyphenationToken? prepared)) {
                    string prefix = token.Substring(0, split);
                    string suffix = token.Substring(split);
                    HtmlInlineRun prefixRun = run.CloneText(prefix, prefix, tokenStyle, run.IsFirstLetter);
                    prefixRun.EndsFirstLine = true;
                    prefixRun.FirstLineHyphen = prepared.HasValue ? run.Style.HyphenateCharacter : string.Empty;
                    prefixRun.PreparedHyphenation = prepared?.SliceSource(0, split);
                    styledRuns.Add(prefixRun);
                    if (suffix.Length > 0) {
                        HtmlInlineRun suffixRun = run.CloneText(suffix, suffix, run.Style, run.IsFirstLetter);
                        suffixRun.PreparedHyphenation = prepared?.SliceSource(split, token.Length - split);
                        styledRuns.Add(suffixRun);
                    }
                    firstLine = false;
                    hasContent = true;
                    continue;
                }
                if (hasContent && previewLine.PreviewAdvance(run, tokenWidth) > width) firstLine = false;
                HtmlRenderBoxStyle appliedStyle = firstLine ? tokenStyle : run.Style;
                styledRuns.Add(run.CloneText(token, token, appliedStyle, run.IsFirstLetter));
                if (firstLine) {
                    previewLine.Add(new InlineSegment(token, tokenWidth, run));
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
        out int split,
        out HyphenationToken? prepared) {
        split = -1;
        prepared = null;
        ChargeLayoutOperations(token.Length, "first-line token search");
        string searchToken = GetFirstLineSearchToken(token, firstLineStyle, width);
        HyphenationToken hyphenation = PrepareHyphenationToken(token, token, layoutStyle);
        if (hyphenation.HasBreaks && firstLineStyle.HyphenateLimitLines != 0) {
            int FindBreak(IReadOnlyList<int> points) => FindLargestFittingFirstLineBreak(
                hyphenation.PaintText, points.Where(point => point > 0 && point < searchToken.Length
                    && point < hyphenation.SourceBoundaries.Count).ToArray(),
                layoutStyle.HyphenateCharacter, firstLineStyle, width);
            int point = FindBreak(hyphenation.PrimaryBreaks);
            if (point < 0) point = FindBreak(hyphenation.SecondaryBreaks);
            if (point > 0) split = hyphenation.SourceBoundaries[point];
            if (split > 0 && split < token.Length) {
                prepared = hyphenation;
                return true;
            }
        }

        IReadOnlyList<int> preferred = GetHtmlPreferredBreakPositions(
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


}
