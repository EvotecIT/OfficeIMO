using OfficeIMO.Drawing;

namespace OfficeIMO.OneNote;

public static partial class OneNotePageRenderer {
    private sealed partial class RenderContext {
        private double RenderElementWithNegativeOffset(
            OneNoteElement element,
            double x,
            double y,
            double availableWidth,
            double availableHeight,
            bool forcePageBounds,
            bool? inheritedRightToLeft) {
            double renderWidth = Math.Max(1D, availableWidth);
            double contentWidth = Math.Max(renderWidth, MeasureElementWidthExtent(element, renderWidth));
            double contentHeight = Math.Max(1D, MeasureElementHeight(element, renderWidth));
            if (x + contentWidth <= 0D || y + contentHeight <= 0D) return contentHeight;

            double localWidth = Math.Max(contentWidth, _drawing.Width - x);
            double localHeight = Math.Max(contentHeight, _drawing.Height - y);
            var localDrawing = new OfficeDrawing(localWidth, localHeight);
            var localContext = new RenderContext(localDrawing, _options, _diagnostics, _pageRightToLeft, _imageCache);
            double used = localContext.RenderElement(
                element,
                0D,
                0D,
                renderWidth,
                availableHeight,
                forcePageBounds,
                inheritedRightToLeft);

            double clipX = Math.Max(0D, x);
            double clipY = Math.Max(0D, y);
            double clipWidth = _drawing.Width - clipX;
            double clipHeight = _drawing.Height - clipY;
            if (clipWidth > 0D && clipHeight > 0D) {
                _drawing.AddClippedDrawing(
                    localDrawing,
                    clipX,
                    clipY,
                    OfficeClipPath.Rectangle(clipWidth, clipHeight),
                    x - clipX,
                    y - clipY);
            }
            return used;
        }
    }
}
