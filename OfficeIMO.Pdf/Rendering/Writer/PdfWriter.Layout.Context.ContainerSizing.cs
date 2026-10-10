namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        /// <summary>Uses canonical frame geometry for a candidate after excluding widths that cannot contain its padding.</summary>
        private bool TryResolveContainerFrame(ContainerBlock? container, PdfPanelStyle style, double parentLeft, double parentWidth,
            out (double X, double Width, double ContentWidth) frame) {
            frame = default;
            if (container?.FrameTable == null && ResolveContainerOuterWidth(style, parentWidth) - 2D * style.PaddingX <= .001D)
                return false;
            frame = ResolveContainerFrame(container, style, parentLeft, parentWidth);
            return frame.ContentWidth > .001D;
        }

        /// <summary>Checks only candidate container widths; actual layout continues to use strict frame validation.</summary>
        private bool HasUsableContainerMeasurementWidths(IReadOnlyList<IPdfBlock> blocks, double frameWidth, bool wholeContainer) {
            foreach (IPdfBlock block in ExpandTransparentMeasurementBlocks(blocks)) {
                if (IsNonVisualFlowMarker(block)) continue;
                if (block is ContainerBlock child) {
                    PdfPanelStyle childStyle = ResolveContainerStyle(child);
                    if (!TryResolveContainerFrame(child, childStyle, 0D, frameWidth, out var childFrame) ||
                        !HasUsableContainerMeasurementWidths(child.Blocks, childFrame.ContentWidth, wholeContainer)) return false;
                } else if (block is FlowBlock flow && flow.StaticBlocks != null &&
                           !HasUsableContainerMeasurementWidths(flow.StaticBlocks, frameWidth, wholeContainer)) return false;
                if (!wholeContainer) break;
            }
            return true;
        }

        /// <summary>Allows a width-dependent first drawing or kept container to start in another valid physical-column width.</summary>
        private bool CanContainerFitAnotherFlowWidth(ContainerBlock container, PdfPanelStyle style, double parentWidth, bool wholeContainer) {
            return CanFitAnotherFixedFlowWidth(parentWidth, candidateWidth => {
                if (!TryResolveContainerFrame(container, style, currentOpts.MarginLeft, candidateWidth, out var frame) ||
                    !HasUsableContainerMeasurementWidths(container.Blocks, frame.ContentWidth, wholeContainer)) return null;
                if (wholeContainer) {
                    return MeasureWholeBlockAtFrameStart(container, currentOpts.MarginLeft, candidateWidth, currentOpts.DefaultFontSize);
                }
                double firstVisual = container.Blocks.Count == 0 ? 0D : MeasureWithContainerPaddingReservation(style, () =>
                    MeasureNextBlockFirstVisualHeight(container.Blocks[0], frame.X + style.PaddingX, frame.ContentWidth,
                        currentOpts.DefaultFontSize, allowTableFragments: true));
                return style.TopPadding + style.FragmentPaddingReservation + style.FragmentBottomInset + firstVisual;
            });
        }

        private static double ResolveContainerOuterWidth(PdfPanelStyle style, double parentWidth) =>
            style.MaxWidth.HasValue ? Math.Min(parentWidth, style.MaxWidth.Value) : parentWidth;
    }
}
