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
            Dictionary<string, int> listIndices = SnapshotListNumbering();
            double renderWidth = Math.Max(1D, availableWidth);
            double contentWidth = Math.Max(renderWidth, MeasureElementWidthExtent(element, renderWidth));
            double contentHeight = Math.Max(1D, MeasureElementHeight(element, renderWidth));
            if (x + contentWidth <= 0D || y + contentHeight <= 0D) {
                AdvanceListNumberingForCulledElement(element);
                return contentHeight;
            }

            double localWidth = Math.Max(contentWidth, _drawing.Width - x);
            double localHeight = Math.Max(contentHeight, _drawing.Height - y);
            var localDrawing = new OfficeDrawing(localWidth, localHeight);
            var localContext = new RenderContext(localDrawing, _options, _diagnostics, _pageRightToLeft, _imageCache);
            localContext.RestoreListNumbering(listIndices);
            double used = localContext.RenderElement(
                element,
                0D,
                0D,
                renderWidth,
                availableHeight,
                forcePageBounds,
                inheritedRightToLeft);
            RestoreListNumbering(localContext.SnapshotListNumbering());

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

        private void AdvanceListNumberingForCulledElement(OneNoteElement element) {
            if (element is OneNoteOutline outline) {
                ResetListNumbering();
                foreach (OneNoteElement child in outline.Children) AdvanceListNumberingForCulledElement(child);
                return;
            }

            if (element is OneNoteParagraph paragraph) {
                if (paragraph.List?.Ordered == true) ResolveListIndex(paragraph.List, advanceListState: true);
                foreach (OneNoteElement child in paragraph.Children) AdvanceListNumberingForCulledElement(child);
                return;
            }

            if (element is OneNoteTable table) {
                foreach (OneNoteTableRow row in table.Rows) {
                    foreach (OneNoteTableCell cell in row.Cells) {
                        foreach (OneNoteElement child in cell.Content) AdvanceListNumberingForCulledElement(child);
                    }
                }
            }
        }
    }
}
