namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        /// <summary>Includes later column widths without discarding the current container's horizontal padding.</summary>
        private double GetMaximumFixedFlowWidth(double containerWidth) {
            if (activeColumnFlow is not { } scope) return containerWidth;
            double maximum = scope.Widths.Max();
            for (int index = scope.ContainerDepth; index < activeContainerScopes.Count; index++)
                maximum = ResolveContainerFrame(activeContainerScopes[index].Style, 0D, maximum).ContentWidth;
            return maximum - Math.Max(0D, width - containerWidth);
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
