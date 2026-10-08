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
                current += run.ReplacedMinWidth;
                maximum = Math.Max(maximum, current);
                if (!run.Style.PreventTextWrapping && !run.IsInlineInset) current = 0D;
                continue;
            }
            if (run.Style.PreventTextWrapping) {
                current += run.Text.IndexOf('\t') >= 0
                    ? MeasureTabExpandedText(run.Text, run.Style, current)
                    : MeasureInlineText(run.Text, run.Style);
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
                        IReadOnlyList<int> preferredBreaks = GetHtmlPreferredBreakPositions(
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
            current += run.IsReplaced
                ? run.ReplacedWidth
                : run.Text.IndexOf('\t') >= 0
                    ? MeasureTabExpandedText(run.Text, run.Style, current)
                    : MeasureMaxContentTextRun(run);
        }
        return Math.Max(maximum, current);
    }

    private double MeasureMaxContentTextRun(IntrinsicTextRun run) {
        double fullWidth = MeasureInlineText(run.Text, run.Style);
        if (!run.Text.Any(char.IsWhiteSpace)) return fullWidth;

        // Inline layout measures separate words and spaces. Some font shapers
        // return a slightly smaller width for the whole string, which can make
        // an auto-sized flex/grid item wrap despite using its max-content width.
        double inlineWidth = 0D;
        foreach (string token in Tokenize(run.Text, run.Style.PreserveWhitespace, run.Style.BreakSpaces)) {
            inlineWidth += MeasureInlineText(!run.Style.PreserveWhitespace && IsWhitespaceToken(token) ? " " : token, run.Style);
        }
        return Math.Max(fullWidth, inlineWidth);
    }

    private IReadOnlyList<IntrinsicTextRun> ResolveInFlowIntrinsicTextRuns(FlexItem item, double availableSize, int depth = 1, bool skipSizedNestedTables = false, bool includeDescendantInsets = false) {
        var rawRuns = new List<IntrinsicTextRun>();
        if (item.Element == null) {
            if (item.Style.Font.Size > 0D) rawRuns.Add(new IntrinsicTextRun(item.TextContent, item.Style));
        } else {
            AppendInFlowIntrinsicTextRuns(item.Element, item.Style, availableSize, depth, rawRuns, skipSizedNestedTables, includeDescendantInsets);
        }
        return NormalizeIntrinsicTextRuns(rawRuns);
    }

    private IReadOnlyList<IntrinsicTextRun> NormalizeIntrinsicTextRuns(IReadOnlyList<IntrinsicTextRun> rawRuns) {
        var normalized = new List<IntrinsicTextRun>();
        HtmlRenderBoxStyle? pendingWhitespaceStyle = null;
        foreach (IntrinsicTextRun run in rawRuns) {
            if (run.IsForcedBreak) {
                pendingWhitespaceStyle = null;
                if (normalized.Count > 0 && !normalized[normalized.Count - 1].IsForcedBreak) {
                    normalized.Add(run);
                }
                continue;
            }
            if (run.IsReplaced) {
                if (pendingWhitespaceStyle != null) {
                    AppendNormalizedIntrinsicText(normalized, " ", pendingWhitespaceStyle);
                    pendingWhitespaceStyle = null;
                }
                normalized.Add(run);
                continue;
            }
            string transformed = ApplyTextTransform(run.Text, run.Style);
            if (run.Style.PreserveWhitespace) {
                if (pendingWhitespaceStyle != null) {
                    AppendNormalizedIntrinsicText(normalized, " ", pendingWhitespaceStyle);
                    pendingWhitespaceStyle = null;
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
                    if (pendingWhitespaceStyle == null
                        && (normalized.Count > 0 && !normalized[normalized.Count - 1].IsForcedBreak || text.Length > 0)) {
                        pendingWhitespaceStyle = run.Style;
                    }
                    continue;
                }
                if (pendingWhitespaceStyle != null) {
                    if (ReferenceEquals(pendingWhitespaceStyle, run.Style)) {
                        text.Append(' ');
                    } else {
                        AppendNormalizedIntrinsicText(normalized, text.ToString(), run.Style);
                        text.Clear();
                        AppendNormalizedIntrinsicText(normalized, " ", pendingWhitespaceStyle);
                    }
                    pendingWhitespaceStyle = null;
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
        if (runs.Count > 0 && !runs[runs.Count - 1].IsForcedBreak
            && !runs[runs.Count - 1].IsReplaced && ReferenceEquals(runs[runs.Count - 1].Style, style)) {
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
        ICollection<IntrinsicTextRun> result,
        bool skipSizedNestedTables,
        bool includeDescendantInsets) {
        if (parentStyle.Display is "flex" or "inline-flex"
            && TryCollectFlexItems(parent, availableSize, parentStyle, depth, captureRunningElements: false,
                out List<FlexItem> items, out _, registerOutOfFlowElements: false)) {
            bool column = parentStyle.FlexDirection is "column" or "column-reverse";
            double width = 0D;
            double minimumWidth = 0D;
            foreach (FlexItem item in items) {
                NormalizeFlexIntrinsicConstraints(item);
                // Resolve the child's content once for both contributions. Recursive
                // flex containers must not duplicate that traversal at every level.
                bool definite = TryResolveDefiniteGridContribution(item, availableSize, out double minimumContribution);
                IReadOnlyList<IntrinsicTextRun>? childRuns = definite ? null : ResolveInFlowIntrinsicTextRuns(item, availableSize, depth + 1);
                double contribution = column ? ResolveColumnFlexCrossBasis(item, availableSize, depth + 1, childRuns)
                    : ClampFlexMainSize(item, ResolveFlexBasis(item, availableSize, depth + 1, childRuns), vertical: false);
                if (!definite) minimumContribution = ResolveGridMeasuredContribution(item.Style, MeasureMinContentRuns(childRuns!));
                if (!column && item.Style.FlexShrink == 0D) minimumContribution = Math.Max(minimumContribution, contribution);
                width = column ? Math.Max(width, contribution) : width + contribution;
                minimumWidth = column || parentStyle.FlexWrap != "nowrap"
                    ? Math.Max(minimumWidth, minimumContribution) : minimumWidth + minimumContribution;
            }
            if (!column) width += parentStyle.ColumnGap * Math.Max(0, items.Count - 1);
            if (!column && parentStyle.FlexWrap == "nowrap") minimumWidth += parentStyle.ColumnGap * Math.Max(0, items.Count - 1);
            result.Add(IntrinsicTextRun.Replaced(minimumWidth, width, parentStyle));
            return;
        }
        // Subgrid axes are sized in their parent's inherited track context.
        // Only standalone grids can use the independent intrinsic-width owner.
        if (parentStyle.Display is "grid" or "inline-grid"
            && !IsSubgridTrackList(parentStyle.GridTemplateColumns) && !IsSubgridTrackList(parentStyle.GridTemplateRows)) {
            (double minimum, double maximum) = ResolveIntrinsicGridWidths(parent, availableSize, parentStyle, depth + 1, clampToAvailable: false);
            double insets = parentStyle.HorizontalInsets + parentStyle.MarginLeft + parentStyle.MarginRight;
            result.Add(IntrinsicTextRun.Replaced(Math.Max(0D, minimum - insets), Math.Max(0D, maximum - insets), parentStyle));
            return;
        }
        AppendGeneratedIntrinsicText(parent, HtmlPseudoElementKind.Before, parentStyle, availableSize, result);
        foreach (INode node in parent.ChildNodes) {
            if (IsClosedDisclosureChild(node)) continue;
            if (node is IText text) {
                if (text.Data.Length > 0 && parentStyle.Font.Size > 0D) result.Add(new IntrinsicTextRun(text.Data, parentStyle));
                continue;
            }
            if (node is not IElement child || ShouldSkipElement(child)) continue;
            EnsureDepth(depth, child);
            HtmlRenderBoxStyle childStyle = _styleResolver.Resolve(child, availableSize, parentStyle);
            childStyle = PrepareButtonChildStyle(child, childStyle);
            if (childStyle.Display == "none" || childStyle.Position == "absolute" || childStyle.Position == "fixed") continue;
            if (includeDescendantInsets && !IsReplacedImageElement(child) && !IsFormControlElement(child.LocalName)) {
                childStyle = _styleResolver.ResolveIntrinsicMeasurementStyle(child, childStyle);
            }
            if (skipSizedNestedTables && string.Equals(child.LocalName, "table", StringComparison.OrdinalIgnoreCase)
                && HtmlRenderStyleResolver.IsBlockElement(child, childStyle)
                && childStyle.ExplicitWidth.HasValue && !childStyle.ExplicitWidthUsesPercentage) {
                // Do not flatten nested table cells into surrounding text. A
                // definite table contributes its declared outer box once.
                result.Add(IntrinsicTextRun.ForcedBreak(childStyle));
                double width = ResolveBoxWidth(availableSize, childStyle) + childStyle.MarginLeft + childStyle.MarginRight;
                result.Add(IntrinsicTextRun.Replaced(width, childStyle));
                result.Add(IntrinsicTextRun.ForcedBreak(childStyle));
                continue;
            }
            if (string.Equals(child.LocalName, "br", StringComparison.OrdinalIgnoreCase)) {
                result.Add(IntrinsicTextRun.ForcedBreak(childStyle));
                continue;
            }
            bool establishesLineBoundary = HtmlRenderStyleResolver.IsBlockElement(child, childStyle)
                && (!includeDescendantInsets || childStyle.FloatSide == "none");
            bool isReplacedChild = IsReplacedImageElement(child)
                || (IsFormControlElement(child.LocalName.ToLowerInvariant()) && !UsesButtonChildLayout(child));
            if (includeDescendantInsets && establishesLineBoundary && !isReplacedChild) {
                HtmlRenderBoxStyle intrinsicStyle = childStyle;
                IReadOnlyList<IntrinsicTextRun> childRuns = ResolveInFlowIntrinsicTextRuns(
                    new FlexItem(child, intrinsicStyle, 0), availableSize, depth + 1,
                    skipSizedNestedTables, includeDescendantInsets);
                double measuredMinimum = MeasureMinContentRuns(childRuns);
                double measuredMaximum = MeasureMaxContentRuns(childRuns);
                double minimum = ResolveGridMeasuredContribution(intrinsicStyle, measuredMinimum);
                double maximum = ResolveGridMeasuredContribution(intrinsicStyle, measuredMaximum);
                if (intrinsicStyle.HasIntrinsicWidths) {
                    (minimum, maximum) = ResolveOrdinaryIntrinsicContributions(intrinsicStyle,
                        measuredMinimum, measuredMaximum, availableSize);
                } else if (intrinsicStyle.ExplicitWidth.HasValue && !intrinsicStyle.ExplicitWidthUsesPercentage) {
                    TryResolveDefiniteGridContribution(new FlexItem(child, intrinsicStyle, 0), availableSize, out double authored);
                    minimum = authored;
                    maximum = authored;
                }
                result.Add(IntrinsicTextRun.ForcedBreak(childStyle));
                result.Add(IntrinsicTextRun.Replaced(minimum, Math.Max(minimum, maximum), intrinsicStyle));
                result.Add(IntrinsicTextRun.ForcedBreak(childStyle));
                continue;
            }
            if (establishesLineBoundary) result.Add(IntrinsicTextRun.ForcedBreak(childStyle));
            if (IsFormControlElement(child.LocalName.ToLowerInvariant()) && !UsesButtonChildLayout(child)) {
                result.Add(IntrinsicTextRun.Replaced(
                    ResolveFormControlIntrinsicOuterWidth(child, childStyle, availableSize), childStyle));
            } else if (IsReplacedImageElement(child)) {
                double width = ResolveIntrinsicReplacedImageBoxWidth(child, childStyle) + childStyle.MarginLeft + childStyle.MarginRight;
                // A cyclic percentage width or maximum can compress replaced content
                // during minimum sizing; its definite constraints still apply.
                double minimumWidth = childStyle.ExplicitWidthUsesPercentage || childStyle.MaxWidthUsesPercentage
                    ? ResolveCompressibleReplacedMinimumWidth(childStyle)
                    : width;
                result.Add(IntrinsicTextRun.Replaced(minimumWidth, width, childStyle));
            } else if (childStyle.Display is "inline-block" or "inline-flex" or "inline-grid"
                && child.LocalName != "math" && (!IsFormControlElement(child.LocalName) || UsesButtonChildLayout(child))) {
                var atomic = new FlexItem(child, childStyle, 0);
                (double Minimum, double Maximum) widths = ResolveGridContentContributions(atomic, availableSize, depth + 1, includeDescendantInsets);
                result.Add(IntrinsicTextRun.Replaced(widths.Minimum, widths.Maximum, parentStyle));
            } else {
                double leadingInset = includeDescendantInsets && childStyle.Display != "contents"
                    ? Math.Max(0D, childStyle.BorderLeftWidth + childStyle.PaddingLeft + childStyle.MarginLeft)
                    : 0D;
                double trailingInset = includeDescendantInsets && childStyle.Display != "contents"
                    ? Math.Max(0D, childStyle.BorderRightWidth + childStyle.PaddingRight + childStyle.MarginRight)
                    : 0D;
                if (leadingInset > 0D) result.Add(IntrinsicTextRun.InlineInset(leadingInset, childStyle));
                AppendInFlowIntrinsicTextRuns(child, childStyle, availableSize, depth + 1, result, skipSizedNestedTables, includeDescendantInsets);
                if (trailingInset > 0D) result.Add(IntrinsicTextRun.InlineInset(trailingInset, childStyle));
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
        if (!_generatedContent.TryGetContent(element, kind, out _)
            || !_styleResolver.TryResolvePseudo(element, kind, availableSize, parentStyle, out HtmlRenderBoxStyle style)
            || style.Display == "none"
            || style.Position == "absolute"
            || style.Position == "fixed") return;
        bool establishesLineBoundary = style.Display == "block" || style.Display == "flow-root" || style.Display == "list-item" || style.Display == "table" || style.Display == "flex" || style.Display == "grid";
        if (establishesLineBoundary) result.Add(IntrinsicTextRun.ForcedBreak(style));
        var inlineRuns = new List<HtmlInlineRun>();
        AddGeneratedInlineRun(element, kind, availableSize, null, parentStyle, null, 0D, 0D, inlineRuns);
        foreach (HtmlInlineRun run in inlineRuns) {
            if (run.AtomicBlock != null) {
                result.Add(IntrinsicTextRun.Replaced(run.AtomicBlock.Width, run.Style));
            } else if (run.Text.Length > 0) {
                result.Add(new IntrinsicTextRun(run.Text, run.Style));
            }
        }
        if (establishesLineBoundary) result.Add(IntrinsicTextRun.ForcedBreak(style));
    }

    private sealed class IntrinsicTextRun {
        internal IntrinsicTextRun(string text, HtmlRenderBoxStyle style, bool isForcedBreak = false, bool isReplaced = false, double replacedWidth = 0D, double? replacedMinWidth = null, bool isInlineInset = false) {
            Text = text;
            Style = style;
            IsForcedBreak = isForcedBreak;
            IsReplaced = isReplaced;
            IsInlineInset = isInlineInset;
            ReplacedWidth = replacedWidth;
            ReplacedMinWidth = replacedMinWidth ?? replacedWidth;
        }

        internal static IntrinsicTextRun ForcedBreak(HtmlRenderBoxStyle style) => new(string.Empty, style, isForcedBreak: true);
        internal static IntrinsicTextRun Replaced(double width, HtmlRenderBoxStyle style) => new(string.Empty, style, isReplaced: true, replacedWidth: width);
        internal static IntrinsicTextRun Replaced(double minimum, double maximum, HtmlRenderBoxStyle style) =>
            new(string.Empty, style, isReplaced: true, replacedWidth: maximum, replacedMinWidth: minimum);
        internal static IntrinsicTextRun InlineInset(double width, HtmlRenderBoxStyle style) =>
            new(string.Empty, style, isReplaced: true, replacedWidth: width, isInlineInset: true);

        internal string Text { get; }
        internal HtmlRenderBoxStyle Style { get; }
        internal bool IsForcedBreak { get; }
        internal bool IsReplaced { get; }
        internal bool IsInlineInset { get; }
        internal double ReplacedWidth { get; }
        internal double ReplacedMinWidth { get; }
    }

}
