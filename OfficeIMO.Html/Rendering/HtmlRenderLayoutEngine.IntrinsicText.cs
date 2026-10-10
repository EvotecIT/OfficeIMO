using System.Text;
using AngleSharp.Dom;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private double MeasureMinContentRuns(IReadOnlyList<IntrinsicTextRun> runs) {
        double maximum = 1D;
        double current = 0D;
        // Indentation contributes to required width without shifting the
        // content-relative tab stops used by ordinary inline layout.
        double firstLineIndent = 0D;
        bool hasContent = false;
        bool pendingAtomicBreak = false;
        var cloned = new List<(HtmlRenderBoxStyle Style, double Start, double End)>();
        var ends = new Dictionary<HtmlRenderBoxStyle, double>();
        foreach (IntrinsicTextRun edge in runs.Where(run => run.IsInlineInset && run.IsInlineInsetEnd)) {
            ends[edge.Style] = edge.ReplacedWidth;
        }
        void AddAdvance(double advance) {
            current += advance;
            hasContent = true;
        }
        void FinishSegment() {
            if (hasContent) maximum = Math.Max(maximum, firstLineIndent + current + cloned.Sum(edge => edge.End));
            current = cloned.Sum(edge => edge.Start);
            firstLineIndent = 0D;
            hasContent = false;
            pendingAtomicBreak = false;
        }
        foreach (IntrinsicTextRun run in runs) {
            if (run.IsParagraphStart) {
                FinishSegment();
                firstLineIndent = run.Style.TextIndent?.ResolveIntrinsicContribution() ?? 0D;
                continue;
            }
            if (run.IsInlineInset) {
                if (!run.IsInlineInsetEnd && pendingAtomicBreak) FinishSegment();
                AddAdvance(run.ReplacedMinWidth);
                if (run.Style.BoxDecorationBreak == "clone") {
                    if (run.IsInlineInsetEnd) {
                        cloned.RemoveAll(edge => ReferenceEquals(edge.Style, run.Style));
                    } else {
                        ends.TryGetValue(run.Style, out double trailing);
                        cloned.Add((run.Style, run.ReplacedWidth, trailing));
                    }
                }
                continue;
            }
            if (run.IsForcedBreak) {
                FinishSegment();
                continue;
            }
            if (pendingAtomicBreak) FinishSegment();
            if (run.IsReplaced) {
                AddAdvance(run.ReplacedMinWidth);
                pendingAtomicBreak = !run.Style.PreventTextWrapping;
                continue;
            }
            if (run.Style.PreventTextWrapping) {
                AddAdvance(run.Text.IndexOf('\t') >= 0
                    ? MeasureTabExpandedText(run.Text, run.Style, current)
                    : MeasureInlineText(run.Text, run.Style));
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
                        IReadOnlyList<string> elements = OfficeTextElements.Split(token);
                        for (int part = 0; part < elements.Count; part++) {
                            AddAdvance(MeasureInlineText(elements[part], run.Style));
                            if (part < elements.Count - 1) FinishSegment();
                        }
                    } else {
                        HyphenationToken hyphenation = PrepareHyphenationToken(token, token, run.Style);
                        string paintToken = hyphenation.PaintText;
                        var hyphenationBreaks = new HashSet<int>(hyphenation.PrimaryBreaks.Concat(hyphenation.SecondaryBreaks));
                        IReadOnlyList<int> preferredBreaks = GetHtmlPreferredBreakPositions(paintToken, run.Style.WordBreak != "keep-all");
                        int segmentStart = 0;
                        foreach (int end in preferredBreaks.Concat(hyphenationBreaks).Distinct().OrderBy(point => point)) {
                            if (end <= segmentStart || end > paintToken.Length) continue;
                            string segment = paintToken.Substring(segmentStart, end - segmentStart);
                            if (hyphenationBreaks.Contains(end)) segment += run.Style.HyphenateCharacter;
                            AddAdvance(MeasureInlineText(segment, run.Style));
                            FinishSegment();
                            segmentStart = end;
                        }
                        if (segmentStart < paintToken.Length) AddAdvance(MeasureInlineText(paintToken.Substring(segmentStart), run.Style));
                    }
                }
                if (!atEnd) {
                    if (run.Style.BreakSpaces) AddAdvance(MeasureInlineText(run.Text[index].ToString(), run.Style));
                    FinishSegment();
                    start = index + 1;
                }
            }
        }
        FinishSegment();
        return maximum;
    }

    private double MeasureMaxContentRuns(IReadOnlyList<IntrinsicTextRun> runs) {
        double maximum = 1D;
        double current = 0D;
        double firstLineIndent = 0D;
        bool hasContent = false;
        void FinishLine() {
            if (hasContent) maximum = Math.Max(maximum, firstLineIndent + current);
            current = 0D;
            firstLineIndent = 0D;
            hasContent = false;
        }
        foreach (IntrinsicTextRun run in runs) {
            if (run.IsParagraphStart) {
                FinishLine();
                firstLineIndent = run.Style.TextIndent?.ResolveIntrinsicContribution() ?? 0D;
                continue;
            }
            if (run.IsForcedBreak) {
                FinishLine();
                continue;
            }
            hasContent = true;
            current += run.IsReplaced
                ? run.ReplacedWidth
                : run.Text.IndexOf('\t') >= 0
                    ? MeasureTabExpandedText(run.Text, run.Style, current)
                    : MeasureMaxContentTextRun(run);
        }
        FinishLine();
        return maximum;
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
            if (item.Style.Font.Size > 0D) {
                rawRuns.Add(IntrinsicTextRun.ParagraphStart(item.Style));
                rawRuns.Add(new IntrinsicTextRun(item.TextContent, item.Style));
            }
        } else {
            AppendInFlowIntrinsicTextRuns(item.Element, item.Style, availableSize, depth, rawRuns, skipSizedNestedTables, includeDescendantInsets);
        }
        return NormalizeIntrinsicTextRuns(rawRuns);
    }

    private IReadOnlyList<IntrinsicTextRun> NormalizeIntrinsicTextRuns(IReadOnlyList<IntrinsicTextRun> rawRuns) {
        var normalized = new List<IntrinsicTextRun>();
        HtmlRenderBoxStyle? pendingWhitespaceStyle = null;
        bool hasLineContent = false;
        bool hasAnyContent = false;
        foreach (IntrinsicTextRun run in rawRuns) {
            if (run.IsParagraphStart) {
                pendingWhitespaceStyle = null;
                hasLineContent = false;
                normalized.Add(run);
                continue;
            }
            if (run.IsForcedBreak) {
                pendingWhitespaceStyle = null;
                hasLineContent = false;
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
                // Inline edges contribute width, but do not make leading
                // collapsible whitespace into content on the formatted line.
                if (!run.IsInlineInset) hasLineContent = true;
                hasAnyContent = true;
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
                        hasLineContent = hasAnyContent = true;
                    }
                    if (atLineBreak) {
                        if (transformed[index] == '\r' && index + 1 < transformed.Length && transformed[index + 1] == '\n') index++;
                        if (normalized.Count > 0 && !normalized[normalized.Count - 1].IsForcedBreak) {
                            normalized.Add(IntrinsicTextRun.ForcedBreak(run.Style));
                        }
                        hasLineContent = false;
                        segmentStart = index + 1;
                    }
                }
                continue;
            }

            var text = new StringBuilder();
            foreach (char current in transformed) {
                if (char.IsWhiteSpace(current)) {
                    if (pendingWhitespaceStyle == null
                        && (hasLineContent || text.Length > 0)) {
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
            hasLineContent = hasAnyContent = true;
        }
        return hasAnyContent ? normalized : Array.Empty<IntrinsicTextRun>();
    }

    private static void AppendNormalizedIntrinsicText(List<IntrinsicTextRun> runs, string text, HtmlRenderBoxStyle style) {
        if (text.Length == 0) return;
        if (runs.Count > 0 && !runs[runs.Count - 1].IsForcedBreak
            && !runs[runs.Count - 1].IsParagraphStart
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
        bool includeDescendantInsets,
        bool startsParagraph = true, IEnumerable<INode>? sourceNodes = null, bool flattenInternalTableBoxes = false) {
        if (parent.LocalName == "table") flattenInternalTableBoxes = parentStyle.Display == "inline";
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
        // Inline descendants retain the surrounding paragraph owner even when
        // they override text-indent. Blocks and atomic boxes measure their own
        // first formatted line; flex/grid containers contribute their items above.
        if (startsParagraph) result.Add(IntrinsicTextRun.ParagraphStart(parentStyle));
        if (sourceNodes == null) AppendGeneratedIntrinsicText(parent, HtmlPseudoElementKind.Before, parentStyle, availableSize, result);
        foreach (INode node in sourceNodes ?? parent.ChildNodes) {
            CheckCancellation();
            if (IsClosedDisclosureChild(node)) continue;
            if (node is IText text) {
                if (text.Data.Length > 0 && parentStyle.Font.Size > 0D) result.Add(new IntrinsicTextRun(text.Data, parentStyle));
                continue;
            }
            if (node is not IElement child || ShouldSkipElement(child)) continue;
            EnsureDepth(depth, child);
            // Empty descendants and cached style/text measurements still cost
            // traversal work before they contribute intrinsic content.
            ChargeLayoutOperation(HtmlRenderStyleResolver.DescribeSource(child));
            HtmlRenderBoxStyle childStyle = _styleResolver.Resolve(child, availableSize, parentStyle);
            childStyle = PrepareButtonChildStyle(child, childStyle);
            childStyle = ForwardContainingHeightBasis(child, childStyle, parentStyle);
            if (childStyle.Display == "none" || childStyle.Position == "absolute" || childStyle.Position == "fixed") continue;
            if (includeDescendantInsets && !IsReplacedImageElement(child) && !IsFormControlElement(child.LocalName)) {
                childStyle = _styleResolver.ResolveIntrinsicMeasurementStyle(child, childStyle);
            }
            if (childStyle.Display == "inline-table" || skipSizedNestedTables && childStyle.Display == "table") {
                bool blockTable = childStyle.Display == "table";
                if (blockTable) result.Add(IntrinsicTextRun.ForcedBreak(childStyle));
                (double Minimum, double Maximum) widths = ResolveIntrinsicTableOuterWidths(child, childStyle, availableSize, depth + 1);
                result.Add(IntrinsicTextRun.Replaced(widths.Minimum, widths.Maximum, childStyle));
                if (blockTable) result.Add(IntrinsicTextRun.ForcedBreak(childStyle));
                continue;
            }
            if (string.Equals(child.LocalName, "br", StringComparison.OrdinalIgnoreCase)) {
                result.Add(IntrinsicTextRun.ForcedBreak(childStyle));
                continue;
            }
            // A native table displayed inline retains its legacy inline text flow;
            // its internal row and cell boxes cannot split the surrounding paragraph.
            bool establishesLineBoundary = HtmlRenderStyleResolver.IsBlockElement(child, childStyle)
                && !(flattenInternalTableBoxes && IsInternalTableDisplay(childStyle.Display))
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
                    ? childStyle.BorderLeftWidth + childStyle.PaddingLeft + childStyle.MarginLeft
                    : 0D;
                double trailingInset = includeDescendantInsets && childStyle.Display != "contents"
                    ? childStyle.BorderRightWidth + childStyle.PaddingRight + childStyle.MarginRight
                    : 0D;
                if (leadingInset != 0D || trailingInset != 0D) result.Add(IntrinsicTextRun.InlineInset(leadingInset, childStyle));
                AppendInFlowIntrinsicTextRuns(child, childStyle, availableSize, depth + 1, result,
                    skipSizedNestedTables, includeDescendantInsets, startsParagraph: establishesLineBoundary,
                    flattenInternalTableBoxes: flattenInternalTableBoxes);
                if (leadingInset != 0D || trailingInset != 0D) result.Add(IntrinsicTextRun.InlineInset(trailingInset, childStyle, closing: true));
            }
            if (establishesLineBoundary) result.Add(IntrinsicTextRun.ForcedBreak(childStyle));
        }
        if (sourceNodes == null) AppendGeneratedIntrinsicText(parent, HtmlPseudoElementKind.After, parentStyle, availableSize, result);
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
        if (establishesLineBoundary) {
            result.Add(IntrinsicTextRun.ForcedBreak(style));
            result.Add(IntrinsicTextRun.ParagraphStart(style));
        }
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
        internal IntrinsicTextRun(string text, HtmlRenderBoxStyle style, bool isForcedBreak = false, bool isReplaced = false, double replacedWidth = 0D, double? replacedMinWidth = null, bool isInlineInset = false, bool isInlineInsetEnd = false, bool isParagraphStart = false) {
            Text = text;
            Style = style;
            IsForcedBreak = isForcedBreak;
            IsParagraphStart = isParagraphStart;
            IsReplaced = isReplaced;
            IsInlineInset = isInlineInset;
            IsInlineInsetEnd = isInlineInsetEnd;
            ReplacedWidth = replacedWidth;
            ReplacedMinWidth = replacedMinWidth ?? replacedWidth;
        }

        internal static IntrinsicTextRun ForcedBreak(HtmlRenderBoxStyle style) => new(string.Empty, style, isForcedBreak: true);
        internal static IntrinsicTextRun ParagraphStart(HtmlRenderBoxStyle style) => new(string.Empty, style, isParagraphStart: true);
        internal static IntrinsicTextRun Replaced(double width, HtmlRenderBoxStyle style) => new(string.Empty, style, isReplaced: true, replacedWidth: width);
        internal static IntrinsicTextRun Replaced(double minimum, double maximum, HtmlRenderBoxStyle style) =>
            new(string.Empty, style, isReplaced: true, replacedWidth: maximum, replacedMinWidth: minimum);
        internal static IntrinsicTextRun InlineInset(double width, HtmlRenderBoxStyle style, bool closing = false) =>
            new(string.Empty, style, isReplaced: true, replacedWidth: width, isInlineInset: true, isInlineInsetEnd: closing);

        internal string Text { get; }
        internal HtmlRenderBoxStyle Style { get; }
        internal bool IsForcedBreak { get; }
        internal bool IsParagraphStart { get; }
        internal bool IsReplaced { get; }
        internal bool IsInlineInset { get; }
        internal bool IsInlineInsetEnd { get; }
        internal double ReplacedWidth { get; }
        internal double ReplacedMinWidth { get; }
    }

}
