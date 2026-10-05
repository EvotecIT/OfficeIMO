namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private ColumnFlowScope? columnBalanceMeasurementScope;

        /// <summary>Uses physical continuation capacity while accounting for active or preflight-only container padding exactly once.</summary>
        private double GetMeasuredContinuationFrameHeight() {
            ColumnFlowScope? scope = columnBalanceMeasurementScope ?? activeColumnFlow;
            double height = scope == null ? yStart - currentOpts.MarginBottom : scope.ParentYStart - scope.ParentOptions.MarginBottom;
            int activeDepth = columnBalanceMeasurementScope == null ? activeContainerScopes.Count : columnBalanceMeasurementScope.ContainerDepth;
            for (int index = 0; index < activeDepth; index++) height -= activeContainerScopes[index].Style.GetFragmentTopPadding(isContinuation: true);
            return Math.Max(0D, height - containerMeasurementTopPadding);
        }

        /// <summary>Includes the full physical page available after columns that start partway down a page.</summary>
        private double GetMaximumTableContinuationFrameHeight(PdfTableStyle style, double currentFrameHeight) {
            if (style.Position != null || activeColumnFlow == null) return currentFrameHeight;
            return Math.Max(currentFrameHeight, GetMeasuredContinuationFrameHeight());
        }
    }
}
