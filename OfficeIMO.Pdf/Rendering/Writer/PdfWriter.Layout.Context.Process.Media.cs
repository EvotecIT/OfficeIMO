using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private void RenderHorizontalRuleFlowBlock(HorizontalRuleBlock hr, IPdfBlock? nextBlock, System.Collections.Generic.IList<IPdfBlock> blockList, int blockIndex) {
            PdfHorizontalRuleStyle ruleStyle = ResolveHorizontalRuleStyle(hr, currentOpts);
            ValidateHorizontalRule(ruleStyle);
            if (ruleStyle.KeepWithNext && nextBlock != null) {
                KeepFixedBlockWithNext(0D, ruleStyle.Thickness, ruleStyle.SpacingBefore, ruleStyle.SpacingAfter, blockList, blockIndex);
            }

            RenderHorizontalRuleBlock(hr, currentOpts.MarginLeft, width);
        }

        private void RenderShapeFlowBlock(ShapeBlock sbk, IPdfBlock? nextBlock, System.Collections.Generic.IList<IPdfBlock> blockList, int blockIndex) {
            PdfDrawingStyle shapeStyle = ResolveDrawingStyle(sbk, currentOpts);
            PdfDocument.ValidateDrawingStyle(shapeStyle, "Shape");
            if (shapeStyle.KeepWithNext && nextBlock != null) {
                KeepFixedBlockWithNext(sbk.Shape.Width, sbk.Shape.Height, shapeStyle.SpacingBefore, shapeStyle.SpacingAfter, blockList, blockIndex);
            }

            RenderShapeBlock(sbk, currentOpts.MarginLeft, width);
        }

        private void RenderDrawingFlowBlock(DrawingBlock dbk, IPdfBlock? nextBlock, System.Collections.Generic.IList<IPdfBlock> blockList, int blockIndex) {
            PdfDrawingStyle drawingStyle = ResolveDrawingStyle(dbk, currentOpts);
            PdfDocument.ValidateDrawingStyle(drawingStyle, "Drawing");
            if (drawingStyle.KeepWithNext && nextBlock != null) {
                KeepFixedBlockWithNext(dbk.Drawing.Width, dbk.Drawing.Height, drawingStyle.SpacingBefore, drawingStyle.SpacingAfter, blockList, blockIndex);
            }

            RenderDrawingBlock(dbk, currentOpts.MarginLeft, width);
        }

        private void RenderImageFlowBlock(ImageBlock ib, IPdfBlock? nextBlock, System.Collections.Generic.IList<IPdfBlock> blockList, int blockIndex) {
            double contentWidth = currentOpts.PageWidth - currentOpts.MarginLeft - currentOpts.MarginRight;
            PdfImageStyle imageStyle = ResolveImageStyle(ib, currentOpts);
            PdfDocument.ValidateImageStyleForBox(imageStyle, ib.Width, ib.Height, nameof(imageStyle.ClipPath));
            PdfDocument.ValidateImageFitDimensions(ib.Info, imageStyle.Fit, nameof(imageStyle.Fit));
            double imageSpacingBefore = ResolveTopLevelSpacingBefore(imageStyle.SpacingBefore);
            var imageBox = ResolveImageFlowBox(ib, imageStyle, contentWidth, imageSpacingBefore, imageStyle.SpacingAfter);
            double needed = imageSpacingBefore + imageBox.Height + imageStyle.SpacingAfter;
            EnsureFixedFlowBlockFits("Image", imageBox.Width, imageBox.Height + imageStyle.SpacingAfter, GetMaximumFixedFlowWidth(contentWidth));
            while (true) {
                bool advance = imageBox.Width > contentWidth + .001D || y - needed < currentOpts.MarginBottom - .001D;
                if (!advance && imageStyle.KeepWithNext && nextBlock != null) {
                    double nextHeight = MeasureKeepWithNextChainHeight(blockList, blockIndex + 1, currentOpts.MarginLeft, width, currentOpts.DefaultFontSize, needed);
                    double keepHeight = needed + nextHeight;
                    advance = nextHeight > .001D && keepHeight <= GetMaximumBlockContinuationHeight() + .001D && ShouldAdvanceForBlockHeight(keepHeight);
                }
                if (!advance) break;
                AdvanceFixedFlowFrame(ref contentWidth);
                imageSpacingBefore = 0D;
                imageBox = ResolveImageFlowBox(ib, imageStyle, contentWidth, imageSpacingBefore, imageStyle.SpacingAfter);
                needed = imageBox.Height + imageStyle.SpacingAfter;
                EnsureFixedFlowBlockFits("Image", imageBox.Width, needed, GetMaximumFixedFlowWidth(contentWidth));
            }
            if (imageSpacingBefore > 0) y -= imageSpacingBefore;
            EnsurePage();
            double xImg = GetAlignedObjectX(currentOpts.MarginLeft, contentWidth, imageBox.Width, imageStyle.Align);
            RecordFlowPlacement(y);
            PageImage pageImage = CreatePageImage(ib, imageStyle, xImg, y - imageBox.Height, imageBox.Width, imageBox.Height);
            currentPage!.Images.Add(pageImage);
            if (!string.IsNullOrWhiteSpace(pageImage.AlternativeText)) {
                int? markedContentId = RegisterFigureStructureElement(pageImage.AlternativeText!);
                pageImage.MarkedContentId = markedContentId;
                pageImage.StructElementIndex = FindStructElementIndex(currentPage, markedContentId, "Figure");
            }

            AddImageLinkAnnotation(ib, imageStyle, pageImage, xImg, y - imageBox.Height, imageBox.Width, imageBox.Height);
            if (currentOpts.Debug?.ShowFlowObjectBoxes == true) {
                pageImage.DebugBox = true;
            }

            pageDirty = true;
            y -= imageBox.Height + imageStyle.SpacingAfter;
        }


    }
}
