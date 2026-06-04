using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private void ProcessBlocks(System.Collections.Generic.IEnumerable<IPdfBlock> sequence) {
            var blockList = sequence as System.Collections.Generic.IList<IPdfBlock> ?? sequence.ToList();
            for (int blockIndex = 0; blockIndex < blockList.Count; blockIndex++) {
                var block = blockList[blockIndex];
                IPdfBlock? nextBlock = blockIndex + 1 < blockList.Count ? blockList[blockIndex + 1] : null;
                if (block is PageBlock pageBlock) {
                    FlushPage(pageDirty || HasCurrentPageNonContentObjects());
                    optionsStack.Push(pageBlock.Options);
                    pageGroupStack.Push(currentPageGroupId);
                    currentOpts = pageBlock.Options;
                    currentPageGroupId = nextPageGroupId++;
                    currentPage = null;
                    StartPage(currentOpts);
                    ProcessBlocks(pageBlock.Blocks);
                    FlushPage(force: true);
                    optionsStack.Pop();
                    currentPageGroupId = pageGroupStack.Pop();
                    currentOpts = optionsStack.Peek();
                    currentPage = null;
                    continue;
                }

                EnsurePage();

                if (block is PageBreakBlock) { NewPage(); continue; }
                if (block is BookmarkBlock bookmark) { AddNamedDestination(bookmark, y); continue; }
                if (block is SpacerBlock spacer) { ConsumeSpacer(spacer.Height); continue; }
                if (block is HeadingBlock heading) { RenderHeadingFlowBlock(heading, nextBlock); continue; }
                if (block is RichParagraphBlock paragraph) { RenderRichParagraphFlowBlock(paragraph, nextBlock); continue; }
                if (block is BulletListBlock bulletList) { RenderBulletListFlowBlock(bulletList, nextBlock); continue; }
                if (block is NumberedListBlock numberedList) { RenderNumberedListFlowBlock(numberedList, nextBlock); continue; }
                if (block is TableBlock table) { RenderTableFlowBlock(table, nextBlock); continue; }
                if (block is HorizontalRuleBlock horizontalRule) { RenderHorizontalRuleFlowBlock(horizontalRule, nextBlock); continue; }
                if (block is TextFieldBlock textField) { RenderTextFieldBlock(textField, currentOpts.MarginLeft, width); continue; }
                if (block is CheckBoxBlock checkBox) { RenderCheckBoxBlock(checkBox, currentOpts.MarginLeft, width); continue; }
                if (block is ChoiceFieldBlock choice) { RenderChoiceFieldBlock(choice, currentOpts.MarginLeft, width); continue; }
                if (block is RadioButtonGroupBlock radioButtonGroup) { RenderRadioButtonGroupBlock(radioButtonGroup, currentOpts.MarginLeft, width); continue; }
                if (block is PdfCanvasBlock canvas) { RenderCanvasBlock(canvas); continue; }
                if (block is ShapeBlock shape) { RenderShapeFlowBlock(shape, nextBlock); continue; }
                if (block is DrawingBlock drawing) { RenderDrawingFlowBlock(drawing, nextBlock); continue; }
                if (block is RowBlock row) { RenderRowFlowBlock(row, nextBlock); continue; }
                if (block is ImageBlock image) { RenderImageFlowBlock(image, nextBlock); continue; }
                if (block is PanelParagraphBlock panel) { RenderPanelFlowBlock(panel, nextBlock); continue; }
            }
        }

    }
}
