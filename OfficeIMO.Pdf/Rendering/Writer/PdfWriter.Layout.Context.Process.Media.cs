using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private void RenderHorizontalRuleFlowBlock(HorizontalRuleBlock hr, IPdfBlock? nextBlock, System.Collections.Generic.IList<IPdfBlock> blockList, int blockIndex) {
            PdfHorizontalRuleStyle ruleStyle = ResolveHorizontalRuleStyle(hr, currentOpts);
            ValidateHorizontalRule(ruleStyle);
            if (ruleStyle.KeepWithNext && nextBlock != null) {
                double needed = ruleStyle.SpacingBefore + ruleStyle.Thickness + ruleStyle.SpacingAfter;
                double nextHeight = MeasureKeepWithNextChainHeight(blockList, blockIndex + 1, currentOpts.MarginLeft, width, currentOpts.DefaultFontSize, needed);
                KeepFixedBlockWithNext(needed, nextHeight);
            }

            RenderHorizontalRuleBlock(hr, currentOpts.MarginLeft, width);
        }

        private void RenderShapeFlowBlock(ShapeBlock sbk, IPdfBlock? nextBlock, System.Collections.Generic.IList<IPdfBlock> blockList, int blockIndex) {
            PdfDrawingStyle shapeStyle = ResolveDrawingStyle(sbk, currentOpts);
            PdfDocument.ValidateDrawingStyle(shapeStyle, "Shape");
            if (shapeStyle.KeepWithNext && nextBlock != null) {
                double needed = shapeStyle.SpacingBefore + sbk.Shape.Height + shapeStyle.SpacingAfter;
                double nextHeight = MeasureKeepWithNextChainHeight(blockList, blockIndex + 1, currentOpts.MarginLeft, width, currentOpts.DefaultFontSize, needed);
                KeepFixedBlockWithNext(needed, nextHeight);
            }

            RenderShapeBlock(sbk, currentOpts.MarginLeft, width);
        }

        private void RenderDrawingFlowBlock(DrawingBlock dbk, IPdfBlock? nextBlock, System.Collections.Generic.IList<IPdfBlock> blockList, int blockIndex) {
            PdfDrawingStyle drawingStyle = ResolveDrawingStyle(dbk, currentOpts);
            PdfDocument.ValidateDrawingStyle(drawingStyle, "Drawing");
            if (drawingStyle.KeepWithNext && nextBlock != null) {
                double needed = drawingStyle.SpacingBefore + dbk.Drawing.Height + drawingStyle.SpacingAfter;
                double nextHeight = MeasureKeepWithNextChainHeight(blockList, blockIndex + 1, currentOpts.MarginLeft, width, currentOpts.DefaultFontSize, needed);
                KeepFixedBlockWithNext(needed, nextHeight);
            }

            RenderDrawingBlock(dbk, currentOpts.MarginLeft, width);
        }

        private void RenderImageFlowBlock(ImageBlock ib, IPdfBlock? nextBlock, System.Collections.Generic.IList<IPdfBlock> blockList, int blockIndex) {
            double xImg = currentOpts.MarginLeft;
            double contentWidth = currentOpts.PageWidth - currentOpts.MarginLeft - currentOpts.MarginRight;
            PdfImageStyle imageStyle = ResolveImageStyle(ib, currentOpts);
            PdfDocument.ValidateImageStyleForBox(imageStyle, ib.Width, ib.Height, nameof(imageStyle.ClipPath));
            PdfDocument.ValidateImageFitDimensions(ib.Info, imageStyle.Fit, nameof(imageStyle.Fit));
            double imageSpacingBefore = ResolveTopLevelSpacingBefore(imageStyle.SpacingBefore);
            var imageBox = ResolveImageFlowBox(ib, imageStyle, contentWidth, imageSpacingBefore, imageStyle.SpacingAfter);
            double needed = imageSpacingBefore + imageBox.Height + imageStyle.SpacingAfter;
            if (imageStyle.Align == PdfAlign.Center) xImg = currentOpts.MarginLeft + Math.Max(0, (contentWidth - imageBox.Width) / 2);
            else if (imageStyle.Align == PdfAlign.Right) xImg = currentOpts.MarginLeft + Math.Max(0, contentWidth - imageBox.Width);
            EnsureFixedFlowBlockFits("Image", imageBox.Width, needed, contentWidth);
            if (imageStyle.KeepWithNext && nextBlock != null) {
                double nextHeight = MeasureKeepWithNextChainHeight(blockList, blockIndex + 1, currentOpts.MarginLeft, width, currentOpts.DefaultFontSize, needed);
                double keepHeight = needed + nextHeight;
                double availableHeight = currentOpts.PageHeight - currentOpts.MarginTop - currentOpts.MarginBottom;
                if (nextHeight > 0.001 && keepHeight <= availableHeight + 0.001 && y < yStart - 0.001 && y - keepHeight < currentOpts.MarginBottom) {
                    NewPage();
                    imageSpacingBefore = 0D;
                    imageBox = ResolveImageFlowBox(ib, imageStyle, contentWidth, imageSpacingBefore, imageStyle.SpacingAfter);
                    needed = imageBox.Height + imageStyle.SpacingAfter;
                    if (imageStyle.Align == PdfAlign.Center) xImg = currentOpts.MarginLeft + Math.Max(0, (contentWidth - imageBox.Width) / 2);
                    else if (imageStyle.Align == PdfAlign.Right) xImg = currentOpts.MarginLeft + Math.Max(0, contentWidth - imageBox.Width);
                }
            }

            if (y - needed < currentOpts.MarginBottom) {
                NewPage();
                imageSpacingBefore = 0D;
            }
            if (imageSpacingBefore > 0) y -= imageSpacingBefore;
            EnsurePage();
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
