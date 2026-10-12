namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        /// <summary>Includes later column widths without discarding the current container's horizontal padding.</summary>
        private double GetMaximumFixedFlowWidth(double containerWidth) {
            if (activeColumnFlow is not { } scope) return containerWidth;
            return ResolveFixedFlowWidth(scope.Widths.Max(), containerWidth);
        }

        /// <summary>Finds the narrowest usable fit while preserving resumed-container padding.</summary>
        private double GetMinimumFixedFlowWidth(double containerWidth) {
            double minimum = containerWidth;
            if (activeColumnFlow is not { } scope) return minimum;
            foreach (double columnWidth in scope.Widths) {
                if (TryResolveFixedFlowWidth(columnWidth, containerWidth, out double available))
                    minimum = Math.Min(minimum, available);
            }
            return minimum;
        }

        private double ResolveFixedFlowWidth(double frameWidth, double containerWidth) {
            if (!TryResolveFixedFlowWidth(frameWidth, containerWidth, out double available))
                throw new ArgumentException("Container padding must leave positive content width.");
            return available;
        }

        /// <summary>Probes a hypothetical column without making an unusable container width a mandatory destination.</summary>
        private bool TryResolveFixedFlowWidth(double frameWidth, double containerWidth, out double available) {
            available = containerWidth;
            if (activeColumnFlow is not { } scope) return available > .001D;
            for (int index = scope.ContainerDepth; index < activeContainerScopes.Count; index++) {
                ContainerRenderScope container = activeContainerScopes[index];
                if (!TryResolveContainerFrame(container.Container, container.Style, 0D, frameWidth, out var frame)) return false;
                frameWidth = frame.ContentWidth;
            }
            available = frameWidth - Math.Max(0D, width - containerWidth);
            return available > .001D;
        }

        /// <summary>Finds a usable physical-column width whose measured content fits without rendering a candidate.</summary>
        private bool CanFitAnotherFixedFlowWidth(double containerWidth, Func<double, double?> measureHeight) {
            if (activeColumnFlow is not { } scope) return false;
            foreach (double columnWidth in scope.Widths) {
                if (!TryResolveFixedFlowWidth(columnWidth, containerWidth, out double candidateWidth)) continue;
                double? height = measureHeight(candidateWidth);
                if (height.HasValue && height.Value <= GetMaximumBlockContinuationHeight() + .001D) return true;
            }
            return false;
        }

        /// <summary>Rebinds the content width after a column or physical-page transition, including resumed containers.</summary>
        private void AdvanceFixedFlowFrame(ref double containerWidth) {
            double previousWidth = width;
            NewBlockFrame();
            containerWidth += width - previousWidth;
        }

        /// <summary>Places an unscaled object in a fitting frame, skipping narrower or unused partial columns.</summary>
        private double PlaceFixedFlowBlock(string name, double objectWidth, double objectHeight,
            double spacingBefore, double spacingAfter, ref double containerWidth) {
            double closingPadding = GetClosingContainerPadding();
            EnsureFixedFlowBlockFits(name, objectWidth, objectHeight + spacingAfter, GetMaximumFixedFlowWidth(containerWidth), closingPadding);
            double before = ResolveTopLevelSpacingBefore(spacingBefore);
            while (objectWidth > containerWidth + .001D || before + objectHeight + spacingAfter + closingPadding > y - currentOpts.MarginBottom + .001D) {
                AdvanceFixedFlowFrame(ref containerWidth);
                before = 0D;
                EnsureFixedFlowBlockFits(name, objectWidth, objectHeight + spacingAfter, GetMaximumFixedFlowWidth(containerWidth), closingPadding);
            }
            return before;
        }
    }
}
