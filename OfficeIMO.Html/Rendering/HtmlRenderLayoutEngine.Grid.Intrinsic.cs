using System.Text;
using AngleSharp.Dom;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private List<double> ResolveGridIntrinsicTrackBases(
        IReadOnlyList<GridTrack> tracks,
        IReadOnlyList<GridIntrinsicContribution> contributions,
        double availableSize,
        double gap,
        bool includeFractionTracks) {
        var sizes = tracks.Select(track => track.IsCollapsed ? 0D : Math.Max(0D, track.Kind == GridTrackKind.Fixed ? Math.Max(track.Value, track.Minimum) : track.Minimum)).ToList();
        foreach (GridIntrinsicContribution contribution in contributions.OrderBy(value => value.Item.ColumnSpan)) {
            GridItem item = contribution.Item;
            IReadOnlyList<GridTrack> spannedTracks = tracks.Skip(item.Column).Take(item.ColumnSpan).ToList();
            bool usesMaxContentContribution = includeFractionTracks && spannedTracks.Any(track => track.Kind == GridTrackKind.Fraction)
                || GridTracksUseMaxContentContribution(spannedTracks);
            double required = (usesMaxContentContribution
                ? ResolveGridMaxContentContribution(item.Item, availableSize)
                : ResolveGridMinContentContribution(item.Item, availableSize)) + contribution.EdgeInsets;
            double current = sizes.Skip(item.Column).Take(item.ColumnSpan).Sum() + gap * Math.Max(0, item.ColumnSpan - 1);
            double deficit = Math.Max(0D, required - current);
            if (deficit <= 0D) continue;
            List<int> intrinsicTracks = Enumerable.Range(item.Column, item.ColumnSpan)
                .Where(index => tracks[index].Kind == GridTrackKind.Auto
                    || tracks[index].Kind == GridTrackKind.Intrinsic
                    || includeFractionTracks && tracks[index].Kind == GridTrackKind.Fraction)
                .ToList();
            if (intrinsicTracks.Count == 0) {
                intrinsicTracks.AddRange(Enumerable.Range(item.Column, item.ColumnSpan)
                    .Where(index => tracks[index].MinimumSizing != GridIntrinsicSizing.None));
            }
            if (intrinsicTracks.Count == 0) continue;
            double remainingDeficit = deficit;
            var growableTracks = new List<int>(intrinsicTracks);
            while (remainingDeficit > 0.0001D && growableTracks.Count > 0) {
                double addition = remainingDeficit / growableTracks.Count;
                double distributed = 0D;
                var nextGrowableTracks = new List<int>(growableTracks.Count);
                foreach (int index in growableTracks) {
                    double availableGrowth = usesMaxContentContribution && tracks[index].GrowthLimit.HasValue
                        ? Math.Max(0D, tracks[index].GrowthLimit!.Value - sizes[index])
                        : double.PositiveInfinity;
                    double growth = Math.Min(addition, availableGrowth);
                    sizes[index] += growth;
                    distributed += growth;
                    if (availableGrowth > growth + 0.0001D) nextGrowableTracks.Add(index);
                }
                if (distributed <= 0.0001D) break;
                remainingDeficit = Math.Max(0D, remainingDeficit - distributed);
                growableTracks = nextGrowableTracks;
            }

            // fit-content() limits max-content growth, but its automatic minimum still
            // has to satisfy the item's min-content contribution.
            double minContentRequired = ResolveGridMinContentContribution(item.Item, availableSize) + contribution.EdgeInsets;
            double minContentAllocated = sizes.Skip(item.Column).Take(item.ColumnSpan).Sum() + gap * Math.Max(0, item.ColumnSpan - 1);
            double minContentDeficit = Math.Max(0D, minContentRequired - minContentAllocated);
            if (minContentDeficit <= 0D) continue;
            List<int> fitContentTracks = intrinsicTracks
                .Where(index => tracks[index].GrowthLimit.HasValue && tracks[index].MinimumSizing == GridIntrinsicSizing.MinContent)
                .ToList();
            if (fitContentTracks.Count == 0) continue;
            double floorAddition = minContentDeficit / fitContentTracks.Count;
            foreach (int index in fitContentTracks) sizes[index] += floorAddition;
        }
        return sizes;
    }

    private void ReportFractionalMinimumFallbacks(
        IReadOnlyList<GridTrack> tracks,
        IReadOnlyList<GridIntrinsicContribution> contributions,
        IReadOnlyList<double> sizes,
        double gap,
        double availableSize) {
        foreach (GridIntrinsicContribution contribution in contributions) {
            GridItem item = contribution.Item;
            bool spansFraction = Enumerable.Range(item.Column, item.ColumnSpan)
                .Any(index => tracks[index].Kind == GridTrackKind.Fraction && !tracks[index].HasExplicitMinimum);
            if (!spansFraction) continue;
            double required = ResolveGridMinContentContribution(item.Item, availableSize) + contribution.EdgeInsets;
            double allocated = sizes.Skip(item.Column).Take(item.ColumnSpan).Sum() + gap * Math.Max(0, item.ColumnSpan - 1);
            if (required <= allocated + 0.0001D) continue;
            string source = item.Item.Element == null
                ? item.Item.TagName
                : HtmlRenderStyleResolver.DescribeSource(item.Item.Element);
            ReportUnsupportedGridValue(source, "fractional automatic minimum exceeds allocated track share");
        }
    }

    private static bool GridTracksUseMaxContentContribution(IReadOnlyList<GridTrack> tracks) =>
        tracks.Any(track =>
            track.MaximumSizing == GridIntrinsicSizing.MaxContent
            || track.MinimumSizing == GridIntrinsicSizing.MaxContent
            || track.Kind == GridTrackKind.Auto);

    private double ResolveGridMinContentContribution(FlexItem item, double availableSize) {
        HtmlRenderBoxStyle style = item.Style;
        if (TryResolveDefiniteGridContribution(item, availableSize, out double definite)) return definite;

        IReadOnlyList<GridIntrinsicTextRun> textRuns = ResolveGridInFlowTextRuns(item, availableSize);
        double measured;
        if (textRuns.Count == 0) {
            measured = 1D;
        } else {
            measured = MeasureGridMinContentRuns(textRuns);
        }
        measured = Math.Max(measured, ResolveDescendantReplacedGridContribution(item, availableSize));

        return ResolveGridMeasuredContribution(style, measured);
    }

    private double ResolveGridMaxContentContribution(FlexItem item, double availableSize) {
        HtmlRenderBoxStyle style = item.Style;
        if (TryResolveDefiniteGridContribution(item, availableSize, out double definite)) return definite;

        IReadOnlyList<GridIntrinsicTextRun> textRuns = ResolveGridInFlowTextRuns(item, availableSize);
        double measured = textRuns.Count == 0
            ? 1D
            : MeasureGridMaxContentRuns(textRuns);
        measured = Math.Max(measured, ResolveDescendantReplacedGridContribution(item, availableSize));
        return ResolveGridMeasuredContribution(style, measured);
    }

    private double MeasureGridMinContentRuns(IReadOnlyList<GridIntrinsicTextRun> runs) {
        double maximum = 1D;
        double current = 0D;
        foreach (GridIntrinsicTextRun run in runs) {
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

    private double MeasureGridMaxContentRuns(IReadOnlyList<GridIntrinsicTextRun> runs) {
        double maximum = 1D;
        double current = 0D;
        foreach (GridIntrinsicTextRun run in runs) {
            if (run.IsForcedBreak) {
                maximum = Math.Max(maximum, current);
                current = 0D;
                continue;
            }
            current += run.IsReplaced ? run.ReplacedWidth : MeasureInlineText(run.Text, run.Style);
        }
        return Math.Max(maximum, current);
    }

    private IReadOnlyList<GridIntrinsicTextRun> ResolveGridInFlowTextRuns(FlexItem item, double availableSize) {
        var rawRuns = new List<GridIntrinsicTextRun>();
        if (item.Element == null) {
            rawRuns.Add(new GridIntrinsicTextRun(item.TextContent, item.Style));
        } else {
            AppendGridInFlowTextRuns(item.Element, item.Style, availableSize, 1, rawRuns);
        }

        var normalized = new List<GridIntrinsicTextRun>();
        bool pendingWhitespace = false;
        foreach (GridIntrinsicTextRun run in rawRuns) {
            if (run.IsForcedBreak) {
                pendingWhitespace = false;
                if (normalized.Count > 0 && !normalized[normalized.Count - 1].IsForcedBreak) {
                    normalized.Add(run);
                }
                continue;
            }
            if (run.IsReplaced) {
                if (pendingWhitespace) {
                    AppendNormalizedGridIntrinsicText(normalized, " ", run.Style);
                    pendingWhitespace = false;
                }
                normalized.Add(run);
                continue;
            }
            string transformed = ApplyTextTransform(run.Text, run.Style);
            if (run.Style.PreserveWhitespace) {
                if (pendingWhitespace) {
                    AppendNormalizedGridIntrinsicText(normalized, " ", run.Style);
                    pendingWhitespace = false;
                }

                int segmentStart = 0;
                for (int index = 0; index <= transformed.Length; index++) {
                    bool atEnd = index == transformed.Length;
                    bool atLineBreak = !atEnd && (transformed[index] == '\r' || transformed[index] == '\n');
                    if (!atEnd && !atLineBreak) continue;
                    if (index > segmentStart) {
                        AppendNormalizedGridIntrinsicText(normalized, transformed.Substring(segmentStart, index - segmentStart), run.Style);
                    }
                    if (atLineBreak) {
                        if (transformed[index] == '\r' && index + 1 < transformed.Length && transformed[index + 1] == '\n') index++;
                        if (normalized.Count > 0 && !normalized[normalized.Count - 1].IsForcedBreak) {
                            normalized.Add(GridIntrinsicTextRun.ForcedBreak(run.Style));
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
            AppendNormalizedGridIntrinsicText(normalized, text.ToString(), run.Style);
        }
        return normalized;
    }

    private static void AppendNormalizedGridIntrinsicText(List<GridIntrinsicTextRun> runs, string text, HtmlRenderBoxStyle style) {
        if (text.Length == 0) return;
        if (runs.Count > 0 && !runs[runs.Count - 1].IsForcedBreak && ReferenceEquals(runs[runs.Count - 1].Style, style)) {
            GridIntrinsicTextRun previous = runs[runs.Count - 1];
            runs[runs.Count - 1] = new GridIntrinsicTextRun(previous.Text + text, style);
        } else {
            runs.Add(new GridIntrinsicTextRun(text, style));
        }
    }

    private void AppendGridInFlowTextRuns(
        IElement parent,
        HtmlRenderBoxStyle parentStyle,
        double availableSize,
        int depth,
        ICollection<GridIntrinsicTextRun> result) {
        AppendGeneratedGridIntrinsicText(parent, HtmlPseudoElementKind.Before, parentStyle, availableSize, result);
        foreach (INode node in parent.ChildNodes) {
            if (node is IText text) {
                if (text.Data.Length > 0) result.Add(new GridIntrinsicTextRun(text.Data, parentStyle));
                continue;
            }
            if (node is not IElement child || ShouldSkipElement(child)) continue;
            EnsureDepth(depth, child);
            HtmlRenderBoxStyle childStyle = _styleResolver.Resolve(child, availableSize, parentStyle);
            if (childStyle.Display == "none" || childStyle.Position == "absolute" || childStyle.Position == "fixed") continue;
            if (string.Equals(child.LocalName, "br", StringComparison.OrdinalIgnoreCase)) {
                result.Add(GridIntrinsicTextRun.ForcedBreak(childStyle));
                continue;
            }
            bool establishesLineBoundary = HtmlRenderStyleResolver.IsBlockElement(child, childStyle);
            if (establishesLineBoundary) result.Add(GridIntrinsicTextRun.ForcedBreak(childStyle));
            if (IsReplacedImageElement(child)) {
                double width = ResolveReplacedImageBoxWidth(child, childStyle) + childStyle.MarginLeft + childStyle.MarginRight;
                result.Add(GridIntrinsicTextRun.Replaced(width, childStyle));
            } else {
                AppendGridInFlowTextRuns(child, childStyle, availableSize, depth + 1, result);
            }
            if (establishesLineBoundary) result.Add(GridIntrinsicTextRun.ForcedBreak(childStyle));
        }
        AppendGeneratedGridIntrinsicText(parent, HtmlPseudoElementKind.After, parentStyle, availableSize, result);
    }

    private void AppendGeneratedGridIntrinsicText(
        IElement element,
        HtmlPseudoElementKind kind,
        HtmlRenderBoxStyle parentStyle,
        double availableSize,
        ICollection<GridIntrinsicTextRun> result) {
        if (!_generatedContent.TryGet(element, kind, out string content)
            || content.Length == 0
            || !_styleResolver.TryResolvePseudo(element, kind, availableSize, parentStyle, out HtmlRenderBoxStyle style)
            || style.Display == "none"
            || style.Position == "absolute"
            || style.Position == "fixed") return;
        bool establishesLineBoundary = style.Display == "block" || style.Display == "flow-root" || style.Display == "list-item" || style.Display == "table" || style.Display == "flex" || style.Display == "grid";
        if (establishesLineBoundary) result.Add(GridIntrinsicTextRun.ForcedBreak(style));
        result.Add(new GridIntrinsicTextRun(content, style));
        if (establishesLineBoundary) result.Add(GridIntrinsicTextRun.ForcedBreak(style));
    }

    private double ResolveDescendantReplacedGridContribution(FlexItem item, double availableSize) {
        if (item.Element == null) return 0D;
        return ResolveDescendantReplacedGridContribution(item.Element, item.Style, availableSize, 1);
    }

    private double ResolveDescendantReplacedGridContribution(IElement parent, HtmlRenderBoxStyle parentStyle, double availableSize, int depth) {
        double maximum = 0D;
        foreach (IElement child in parent.Children) {
            EnsureDepth(depth, child);
            if (ShouldSkipElement(child)) continue;
            HtmlRenderBoxStyle childStyle = _styleResolver.Resolve(child, availableSize, parentStyle);
            if (childStyle.Display == "none" || childStyle.Position == "absolute" || childStyle.Position == "fixed") continue;
            double contribution;
            if (IsReplacedImageElement(child)) {
                contribution = ResolveReplacedImageBoxWidth(child, childStyle) + childStyle.MarginLeft + childStyle.MarginRight;
            } else {
                double descendant = ResolveDescendantReplacedGridContribution(child, childStyle, availableSize, depth + 1);
                contribution = descendant > 0D ? ResolveGridMeasuredContribution(childStyle, descendant) : 0D;
            }
            maximum = Math.Max(maximum, contribution);
        }
        return maximum;
    }

    private bool TryResolveDefiniteGridContribution(FlexItem item, double availableSize, out double contribution) {
        HtmlRenderBoxStyle style = item.Style;
        if (style.ExplicitWidth.HasValue) {
            double boxWidth = style.ExplicitWidth.Value + (style.BorderBox ? 0D : style.HorizontalInsets);
            if (style.MaxWidth.HasValue) boxWidth = Math.Min(boxWidth, style.MaxWidth.Value + (style.BorderBox ? 0D : style.HorizontalInsets));
            if (style.MinWidth.HasValue) boxWidth = Math.Max(boxWidth, style.MinWidth.Value + (style.BorderBox ? 0D : style.HorizontalInsets));
            contribution = Math.Max(1D, boxWidth + style.MarginLeft + style.MarginRight);
            return true;
        }
        if (IsReplacedImageElementTag(item.TagName) && item.Element != null) {
            contribution = Math.Max(1D, ResolveReplacedImageBoxWidth(item.Element, style) + style.MarginLeft + style.MarginRight);
            return true;
        }
        if (item.TagName == "table") {
            contribution = ResolveColumnFlexCrossBasis(item, availableSize);
            return true;
        }

        contribution = 0D;
        return false;
    }

    private static double ResolveGridMeasuredContribution(HtmlRenderBoxStyle style, double measured) {
        double boxBasis = measured + style.HorizontalInsets;
        if (style.MaxWidth.HasValue) boxBasis = Math.Min(boxBasis, style.MaxWidth.Value + (style.BorderBox ? 0D : style.HorizontalInsets));
        if (style.MinWidth.HasValue) boxBasis = Math.Max(boxBasis, style.MinWidth.Value + (style.BorderBox ? 0D : style.HorizontalInsets));
        double outer = boxBasis + style.MarginLeft + style.MarginRight;
        return Math.Max(1D, outer);
    }

    private sealed class GridIntrinsicTextRun {
        internal GridIntrinsicTextRun(string text, HtmlRenderBoxStyle style, bool isForcedBreak = false, bool isReplaced = false, double replacedWidth = 0D) {
            Text = text;
            Style = style;
            IsForcedBreak = isForcedBreak;
            IsReplaced = isReplaced;
            ReplacedWidth = replacedWidth;
        }

        internal static GridIntrinsicTextRun ForcedBreak(HtmlRenderBoxStyle style) => new(string.Empty, style, isForcedBreak: true);
        internal static GridIntrinsicTextRun Replaced(double width, HtmlRenderBoxStyle style) => new(string.Empty, style, isReplaced: true, replacedWidth: width);

        internal string Text { get; }
        internal HtmlRenderBoxStyle Style { get; }
        internal bool IsForcedBreak { get; }
        internal bool IsReplaced { get; }
        internal double ReplacedWidth { get; }
    }

}
