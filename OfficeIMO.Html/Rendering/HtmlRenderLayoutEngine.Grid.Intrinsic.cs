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

        IReadOnlyList<IntrinsicTextRun> textRuns = ResolveInFlowIntrinsicTextRuns(item, availableSize);
        double measured;
        if (textRuns.Count == 0) {
            measured = 1D;
        } else {
            measured = MeasureMinContentRuns(textRuns);
        }
        measured = Math.Max(measured, ResolveDescendantReplacedGridContribution(item, availableSize));

        return ResolveGridMeasuredContribution(style, measured);
    }

    private double ResolveGridMaxContentContribution(FlexItem item, double availableSize) {
        HtmlRenderBoxStyle style = item.Style;
        if (TryResolveDefiniteGridContribution(item, availableSize, out double definite)) return definite;

        IReadOnlyList<IntrinsicTextRun> textRuns = ResolveInFlowIntrinsicTextRuns(item, availableSize);
        double measured = textRuns.Count == 0
            ? 1D
            : MeasureMaxContentRuns(textRuns);
        measured = Math.Max(measured, ResolveDescendantReplacedGridContribution(item, availableSize));
        return ResolveGridMeasuredContribution(style, measured);
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


}
