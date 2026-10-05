using System.Text;
using AngleSharp.Dom;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private double MeasureMinContentRuns(IReadOnlyList<IntrinsicTextRun> runs) {
        double maximum = 1D;
        double current = 0D;
        foreach (IntrinsicTextRun run in runs) {
            if (run.IsForcedBreak) {
                maximum = Math.Max(maximum, current);
                current = 0D;
                continue;
            }
            if (run.IsReplaced) {
                current += run.ReplacedWidth;
                maximum = Math.Max(maximum, current);
                if (!run.Style.PreventTextWrapping) current = 0D;
                continue;
            }
            if (run.Style.PreventTextWrapping) {
                current += MeasureInlineText(run.Text, run.Style);
                maximum = Math.Max(maximum, current);
                continue;
            }

            int start = 0;
            for (int index = 0; index <= run.Text.Length; index++) {
                bool atEnd = index == run.Text.Length;
                if (!atEnd && !char.IsWhiteSpace(run.Text[index])) continue;
                if (index > start) {
                    string token = run.Text.Substring(start, index - start);
                    bool breakEverywhere = run.Style.WordBreak == "break-all" || run.Style.OverflowWrap == "anywhere";
                    if (breakEverywhere) {
                        foreach (string element in OfficeTextElements.Split(token)) {
                            current += MeasureInlineText(element, run.Style);
                            maximum = Math.Max(maximum, current);
                            current = 0D;
                        }
                    } else {
                        HyphenationToken hyphenation = PrepareHyphenationToken(token, token, run.Style);
                        string paintToken = hyphenation.PaintText;
                        var hyphenationBreaks = new HashSet<int>(hyphenation.PrimaryBreaks.Concat(hyphenation.SecondaryBreaks));
                        IReadOnlyList<int> preferredBreaks = OfficeTextLineBreaks.GetBreakPositions(
                            paintToken,
                            run.Style.WordBreak != "keep-all");
                        int segmentStart = 0;
                        foreach (int end in preferredBreaks.Concat(hyphenationBreaks).Distinct().OrderBy(point => point)) {
                            if (end <= segmentStart || end > paintToken.Length) continue;
                            string segment = paintToken.Substring(segmentStart, end - segmentStart);
                            if (hyphenationBreaks.Contains(end)) segment += run.Style.HyphenateCharacter;
                            current += MeasureInlineText(segment, run.Style);
                            maximum = Math.Max(maximum, current);
                            current = 0D;
                            segmentStart = end;
                        }
                        if (segmentStart < paintToken.Length) current += MeasureInlineText(paintToken.Substring(segmentStart), run.Style);
                        maximum = Math.Max(maximum, current);
                    }
                }
                if (!atEnd) {
                    if (run.Style.BreakSpaces) {
                        current += MeasureInlineText(run.Text[index].ToString(), run.Style);
                    }
                    maximum = Math.Max(maximum, current);
                    current = 0D;
                    start = index + 1;
                }
            }
        }
        return Math.Max(maximum, current);
    }

    private double MeasureMaxContentRuns(IReadOnlyList<IntrinsicTextRun> runs) {
        double maximum = 1D;
        double current = 0D;
        foreach (IntrinsicTextRun run in runs) {
            if (run.IsForcedBreak) {
                maximum = Math.Max(maximum, current);
                current = 0D;
                continue;
            }
            if (run.IsReplaced) {
                current += run.ReplacedWidth;
                continue;
            }
            // Line layout measures tokens separately. Measuring a whole run can
            // kern across token boundaries and underestimate its natural width.
            foreach (string token in Tokenize(run.Text, run.Style.PreserveWhitespace, run.Style.BreakSpaces)) {
                if (token == "\u2028" || run.Style.PreserveWhitespace && token == "\n") {
                    maximum = Math.Max(maximum, current);
                    current = 0D;
                    continue;
                }
                string paint = token.IndexOf('\u00AD') < 0 ? token : token.Replace("\u00AD", string.Empty);
                if (paint.Length == 0) continue;
                current += run.Style.PreserveWhitespace && paint.IndexOf('\t') >= 0
                    ? MeasureTabExpandedText(paint, run.Style, current)
                    : MeasureInlineText(paint, run.Style);
            }
        }
        return Math.Max(maximum, current);
    }

    /// <summary>Collects styled in-flow and generated content for grid and flex intrinsic sizing.</summary>
    private IReadOnlyList<IntrinsicTextRun> ResolveInFlowIntrinsicTextRuns(FlexItem item, double availableSize) {
        var rawRuns = new List<IntrinsicTextRun>();
        if (item.Element == null) {
            rawRuns.Add(new IntrinsicTextRun(item.TextContent, item.Style));
        } else {
            AppendInFlowIntrinsicTextRuns(item.Element, item.Style, availableSize, 1, rawRuns);
        }

        var normalized = new List<IntrinsicTextRun>();
        bool pendingWhitespace = false;
        foreach (IntrinsicTextRun run in rawRuns) {
            if (run.IsForcedBreak) {
                pendingWhitespace = false;
                if (normalized.Count > 0 && !normalized[normalized.Count - 1].IsForcedBreak) {
                    normalized.Add(run);
                }
                continue;
            }
            if (run.IsReplaced) {
                if (pendingWhitespace) {
                    AppendNormalizedIntrinsicText(normalized, " ", run.Style);
                    pendingWhitespace = false;
                }
                normalized.Add(run);
                continue;
            }
            string transformed = ApplyTextTransform(run.Text, run.Style);
            if (run.Style.PreserveWhitespace) {
                if (pendingWhitespace) {
                    AppendNormalizedIntrinsicText(normalized, " ", run.Style);
                    pendingWhitespace = false;
                }

                int segmentStart = 0;
                for (int index = 0; index <= transformed.Length; index++) {
                    bool atEnd = index == transformed.Length;
                    bool atLineBreak = !atEnd && (transformed[index] == '\r' || transformed[index] == '\n');
                    if (!atEnd && !atLineBreak) continue;
                    if (index > segmentStart) {
                        AppendNormalizedIntrinsicText(normalized, transformed.Substring(segmentStart, index - segmentStart), run.Style);
                    }
                    if (atLineBreak) {
                        if (transformed[index] == '\r' && index + 1 < transformed.Length && transformed[index + 1] == '\n') index++;
                        if (normalized.Count > 0 && !normalized[normalized.Count - 1].IsForcedBreak) {
                            normalized.Add(IntrinsicTextRun.ForcedBreak(run.Style));
                        }
                        segmentStart = index + 1;
                    }
                }
                continue;
            }

            var text = new StringBuilder();
            foreach (char current in transformed) {
                if (char.IsWhiteSpace(current)) {
                    pendingWhitespace = normalized.Count > 0 || text.Length > 0;
                    continue;
                }
                if (pendingWhitespace) {
                    text.Append(' ');
                    pendingWhitespace = false;
                }
                text.Append(current);
            }
            if (text.Length == 0) continue;
            AppendNormalizedIntrinsicText(normalized, text.ToString(), run.Style);
        }
        return normalized;
    }

    private static void AppendNormalizedIntrinsicText(List<IntrinsicTextRun> runs, string text, HtmlRenderBoxStyle style) {
        if (text.Length == 0) return;
        if (runs.Count > 0 && !runs[runs.Count - 1].IsForcedBreak && ReferenceEquals(runs[runs.Count - 1].Style, style)) {
            IntrinsicTextRun previous = runs[runs.Count - 1];
            runs[runs.Count - 1] = new IntrinsicTextRun(previous.Text + text, style);
        } else {
            runs.Add(new IntrinsicTextRun(text, style));
        }
    }

    private void AppendInFlowIntrinsicTextRuns(
        IElement parent,
        HtmlRenderBoxStyle parentStyle,
        double availableSize,
        int depth,
        ICollection<IntrinsicTextRun> result) {
        AppendGeneratedIntrinsicText(parent, HtmlPseudoElementKind.Before, parentStyle, availableSize, result);
        foreach (INode node in parent.ChildNodes) {
            if (node is IText text) {
                if (text.Data.Length > 0) result.Add(new IntrinsicTextRun(text.Data, parentStyle));
                continue;
            }
            if (node is not IElement child || ShouldSkipElement(child)) continue;
            EnsureDepth(depth, child);
            HtmlRenderBoxStyle childStyle = _styleResolver.Resolve(child, availableSize, parentStyle);
            if (childStyle.Display == "none" || childStyle.Position == "absolute" || childStyle.Position == "fixed") continue;
            if (string.Equals(child.LocalName, "br", StringComparison.OrdinalIgnoreCase)) {
                result.Add(IntrinsicTextRun.ForcedBreak(childStyle));
                continue;
            }
            bool establishesLineBoundary = HtmlRenderStyleResolver.IsBlockElement(child, childStyle);
            if (establishesLineBoundary) result.Add(IntrinsicTextRun.ForcedBreak(childStyle));
            if (IsReplacedImageElement(child)) {
                double width = ResolveReplacedImageBoxWidth(child, childStyle) + childStyle.MarginLeft + childStyle.MarginRight;
                result.Add(IntrinsicTextRun.Replaced(width, childStyle));
            } else {
                AppendInFlowIntrinsicTextRuns(child, childStyle, availableSize, depth + 1, result);
            }
            if (establishesLineBoundary) result.Add(IntrinsicTextRun.ForcedBreak(childStyle));
        }
        AppendGeneratedIntrinsicText(parent, HtmlPseudoElementKind.After, parentStyle, availableSize, result);
    }

    private void AppendGeneratedIntrinsicText(
        IElement element,
        HtmlPseudoElementKind kind,
        HtmlRenderBoxStyle parentStyle,
        double availableSize,
        ICollection<IntrinsicTextRun> result) {
        if (!_generatedContent.TryGet(element, kind, out string content)
            || content.Length == 0
            || !_styleResolver.TryResolvePseudo(element, kind, availableSize, parentStyle, out HtmlRenderBoxStyle style)
            || style.Display == "none"
            || style.Position == "absolute"
            || style.Position == "fixed") return;
        bool establishesLineBoundary = style.Display == "block" || style.Display == "flow-root" || style.Display == "list-item" || style.Display == "table" || style.Display == "flex" || style.Display == "grid";
        if (establishesLineBoundary) result.Add(IntrinsicTextRun.ForcedBreak(style));
        result.Add(new IntrinsicTextRun(content, style));
        if (establishesLineBoundary) result.Add(IntrinsicTextRun.ForcedBreak(style));
    }

    private sealed class IntrinsicTextRun {
        internal IntrinsicTextRun(string text, HtmlRenderBoxStyle style, bool isForcedBreak = false, bool isReplaced = false, double replacedWidth = 0D) {
            Text = text;
            Style = style;
            IsForcedBreak = isForcedBreak;
            IsReplaced = isReplaced;
            ReplacedWidth = replacedWidth;
        }

        internal static IntrinsicTextRun ForcedBreak(HtmlRenderBoxStyle style) => new(string.Empty, style, isForcedBreak: true);
        internal static IntrinsicTextRun Replaced(double width, HtmlRenderBoxStyle style) => new(string.Empty, style, isReplaced: true, replacedWidth: width);

        internal string Text { get; }
        internal HtmlRenderBoxStyle Style { get; }
        internal bool IsForcedBreak { get; }
        internal bool IsReplaced { get; }
        internal double ReplacedWidth { get; }
    }

}
