namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    /// <summary>Measures an atomic inline block once for both flex and grid intrinsic sizing.</summary>
    private GridIntrinsicContributions ResolveInlineBlockIntrinsicContributions(FlexItem item, double availableSize, int depth) {
        HtmlRenderBoxStyle style = item.Style;
        if (style.ExplicitWidthUsesPercentage) {
            // The percentage depends on the auto-sized ancestor being measured.
            style = style.Clone();
            style.ExplicitWidth = null;
            style.ExplicitWidthUsesPercentage = false;
            item.Style = style;
        }
        if (!style.HasIntrinsicWidths && TryResolveDefiniteGridContribution(item, availableSize, out double definite)) {
            return new GridIntrinsicContributions(definite, definite);
        }

        IReadOnlyList<IntrinsicTextRun> runs = ResolveInFlowIntrinsicTextRuns(item, availableSize, depth, includeDescendantInsets: style.HasIntrinsicWidths);
        double minimum = Math.Max(runs.Count == 0 ? 1D : MeasureMinContentRuns(runs),
            ResolveDescendantReplacedGridContribution(item, availableSize, minimum: true));
        double maximum = Math.Max(runs.Count == 0 ? 1D : MeasureMaxContentRuns(runs),
            ResolveDescendantReplacedGridContribution(item, availableSize));
        if (style.HasIntrinsicWidths) {
            (double low, double high) = ResolveOrdinaryIntrinsicContributions(style, minimum, maximum, availableSize);
            return new GridIntrinsicContributions(low, high);
        }
        return new GridIntrinsicContributions(
            ResolveGridMeasuredContribution(style, minimum),
            ResolveGridMeasuredContribution(style, maximum));
    }

    private readonly record struct GridIntrinsicContributions(double Minimum, double Maximum);
}
