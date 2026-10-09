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
                double available = ResolveFixedFlowWidth(columnWidth, containerWidth);
                if (available > 0D) minimum = Math.Min(minimum, available);
            }
            return minimum;
        }

        private double ResolveFixedFlowWidth(double frameWidth, double containerWidth) {
            if (activeColumnFlow is not { } scope) return containerWidth;
            for (int index = scope.ContainerDepth; index < activeContainerScopes.Count; index++)
                frameWidth = ResolveContainerFrame(activeContainerScopes[index].Container, activeContainerScopes[index].Style, 0D, frameWidth).ContentWidth;
            return frameWidth - Math.Max(0D, width - containerWidth);
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
