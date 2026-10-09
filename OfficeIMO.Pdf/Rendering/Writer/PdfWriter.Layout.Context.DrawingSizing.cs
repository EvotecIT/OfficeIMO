using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        /// <summary>Resolves drawing dimensions against the current frame without changing the retained scene.</summary>
        private static (double Width, double Height) ResolveDrawingFlowBox(DrawingBlock block, PdfDrawingStyle style, double frameWidth) {
            OfficeDrawing drawing = block.Drawing;
            if (!style.ConstrainToContentWidth || drawing.Width <= frameWidth) return (drawing.Width, drawing.Height);
            Guard.Positive(frameWidth, nameof(frameWidth));
            double scale = frameWidth / drawing.Width;
            return (frameWidth, drawing.Height * scale);
        }

        /// <summary>Uses a drawing transform so vector geometry, text, images, and embedded links share the proportional fit.</summary>
        private static OfficeDrawing ResolveDrawingFlowScene(DrawingBlock block, PdfDrawingStyle style, double frameWidth) {
            var box = ResolveDrawingFlowBox(block, style, frameWidth);
            if (box.Width == block.Drawing.Width) return block.Drawing;
            double scale = box.Width / block.Drawing.Width;
            return new OfficeDrawing(box.Width, box.Height)
                .AddEffectDrawing(block.Drawing, OfficeTransform.Scale(scale, scale));
        }

        /// <summary>Recomputes the fit after advancing to a column or page with a different available width.</summary>
        private (double Width, double Height, double SpacingBefore) PlaceDrawingFlowBlock(DrawingBlock block, PdfDrawingStyle style, ref double containerWidth) {
            double before = ResolveTopLevelSpacingBefore(style.SpacingBefore);
            while (true) {
                var box = ResolveDrawingFlowBox(block, style, containerWidth);
                double closingPadding = GetClosingContainerPadding();
                EnsureFixedFlowBlockFits("Drawing", box.Width, box.Height + style.SpacingAfter,
                    GetMaximumFixedFlowWidth(containerWidth), closingPadding);
                if (box.Width <= containerWidth + .001D &&
                    before + box.Height + style.SpacingAfter + closingPadding <= y - currentOpts.MarginBottom + .001D) {
                    return (box.Width, box.Height, before);
                }
                AdvanceFixedFlowFrame(ref containerWidth);
                before = 0D;
            }
        }

        /// <summary>Measures the kept drawing and following content again when their destination frame changes.</summary>
        private void KeepDrawingBlockWithNext(DrawingBlock block, PdfDrawingStyle style, IList<IPdfBlock> blocks, int blockIndex) {
            while (true) {
                var box = ResolveDrawingFlowBox(block, style, width);
                double needed = ResolveTopLevelSpacingBefore(style.SpacingBefore) + box.Height + style.SpacingAfter;
                EnsureFixedFlowBlockFits("Kept drawing", box.Width, box.Height + style.SpacingAfter, GetMaximumFixedFlowWidth(width));
                if (box.Width > width + .001D || ShouldAdvanceForBlockHeight(needed)) {
                    NewBlockFrame();
                    continue;
                }
                double nextHeight = MeasureKeepWithNextChainHeight(blocks, blockIndex + 1, currentOpts.MarginLeft,
                    width, currentOpts.DefaultFontSize, needed);
                double keepHeight = needed + nextHeight;
                if (nextHeight <= .001D || keepHeight > GetMaximumBlockContinuationHeight() + .001D || !ShouldAdvanceForBlockHeight(keepHeight)) break;
                NewBlockFrame();
            }
        }
    }
}
