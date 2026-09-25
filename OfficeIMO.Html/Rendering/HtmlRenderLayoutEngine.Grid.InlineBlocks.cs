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
        if (TryResolveDefiniteGridContribution(item, availableSize, out double definite)) {
            return new GridIntrinsicContributions(definite, definite);
        }

        IReadOnlyList<GridIntrinsicTextRun> runs = ResolveGridInFlowTextRuns(item, availableSize, depth);
        double replaced = ResolveDescendantReplacedGridContribution(item, availableSize);
        double minimum = Math.Max(runs.Count == 0 ? 1D : MeasureGridMinContentRuns(runs), replaced);
        double maximum = Math.Max(runs.Count == 0 ? 1D : MeasureGridMaxContentRuns(runs), replaced);
        return new GridIntrinsicContributions(
            ResolveGridMeasuredContribution(style, minimum),
            ResolveGridMeasuredContribution(style, maximum));
    }

    private readonly record struct GridIntrinsicContributions(double Minimum, double Maximum);
}
