namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private void RenderMultiColumnBlock(MultiColumnBlock columns) {
            PdfMultiColumnOptions options = columns.Options;
            double totalGap = options.Gap * (options.ColumnCount - 1);
            if (totalGap >= width) throw new ArgumentException("Multi-column gaps must leave positive column widths.");
            double columnWidth = (width - totalGap) / options.ColumnCount;
            ValidateColumnBlocks(columns.Blocks);
            int blockIndex = 0;
            while (blockIndex < columns.Blocks.Count) {
                while (blockIndex < columns.Blocks.Count && columns.Blocks[blockIndex] is ColumnBreakBlock) blockIndex++;
                if (blockIndex >= columns.Blocks.Count) break;
                double availableHeight = y - currentOpts.MarginBottom;
                if (availableHeight < currentOpts.DefaultFontSize * 1.4D) {
                    NewPage();
                    availableHeight = y - currentOpts.MarginBottom;
                }

                double remainingHeight = 0D;
                for (int i = blockIndex; i < columns.Blocks.Count; i++) {
                    if (columns.Blocks[i] is not ColumnBreakBlock) remainingHeight += MeasureColumnBlock(columns.Blocks[i], columnWidth);
                }
                double target = options.BalanceLastPage && remainingHeight <= availableHeight * options.ColumnCount
                    ? Math.Max(currentOpts.DefaultFontSize * 1.4D, remainingHeight / options.ColumnCount)
                    : availableHeight;

                var row = new RowBlock();
                row.SetGap(options.Gap);
                row.SetStyle(new PdfRowStyle {
                    Gap = options.Gap,
                    ColumnSeparatorColor = options.SeparatorColor,
                    ColumnSeparatorWidth = options.SeparatorWidth
                });
                double widthPercent = 100D / options.ColumnCount;
                for (int columnIndex = 0; columnIndex < options.ColumnCount; columnIndex++) {
                    var column = new RowColumn(widthPercent);
                    double consumed = 0D;
                    while (blockIndex < columns.Blocks.Count) {
                        IPdfBlock block = columns.Blocks[blockIndex];
                        if (block is ColumnBreakBlock) {
                            blockIndex++;
                            break;
                        }

                        double blockHeight = MeasureColumnBlock(block, columnWidth);
                        if (column.Blocks.Count > 0 && consumed + blockHeight > target + 0.001D) break;
                        column.AddBlock(block);
                        consumed += blockHeight;
                        blockIndex++;
                        if (consumed >= target - 0.001D) break;
                    }
                    row.AddColumn(column);
                }

                RenderRowFlowBlock(row, nextBlock: null, new List<IPdfBlock> { row }, 0);
                if (blockIndex < columns.Blocks.Count) NewPage();
            }
        }

        private double MeasureColumnBlock(IPdfBlock block, double columnWidth) =>
            MeasureKeepWithNextBlockHeight(block, currentOpts.MarginLeft, columnWidth, currentOpts.DefaultFontSize);

        private static void ValidateColumnBlocks(IReadOnlyList<IPdfBlock> blocks) {
            foreach (IPdfBlock block in blocks) {
                if (block is HeadingBlock or RichParagraphBlock or BulletListBlock or NumberedListBlock or
                    PanelParagraphBlock or TableBlock or HorizontalRuleBlock or ImageBlock or ShapeBlock or DrawingBlock or
                    TextFieldBlock or CheckBoxBlock or ChoiceFieldBlock or RadioButtonGroupBlock or BookmarkBlock or SpacerBlock or ColumnBreakBlock) {
                    continue;
                }

                throw new NotSupportedException("Automatic multi-column flow does not support nested block type " + block.GetType().Name + ". Use separate Columns blocks around that content.");
            }
        }

        private void RenderContainerBlock(ContainerBlock container) {
            ValidateColumnBlocks(container.Blocks);
            PanelStyle style = container.Style;
            double outerWidth = style.MaxWidth.HasValue ? Math.Min(width, style.MaxWidth.Value) : width;
            ValidatePanelStyle(style, outerWidth);
            double outerX = style.Align switch {
                PdfAlign.Center => currentOpts.MarginLeft + (width - outerWidth) / 2D,
                PdfAlign.Right => currentOpts.MarginLeft + width - outerWidth,
                _ => currentOpts.MarginLeft
            };
            double contentWidth = outerWidth - 2D * style.PaddingX;
            if (contentWidth <= 0.001D) throw new ArgumentException("Container padding must leave positive content width.");

            var row = new RowBlock();
            row.SetGap(0D);
            row.SetStyle(new PdfRowStyle {
                Gap = 0D,
                SpacingBefore = style.SpacingBefore,
                SpacingAfter = style.SpacingAfter,
                KeepTogether = true,
                KeepWithNext = style.KeepWithNext
            });

            double leadingBlank = outerX - currentOpts.MarginLeft + style.PaddingX;
            double trailingBlank = width - leadingBlank - contentWidth;
            if (leadingBlank > 0.001D) row.AddColumn(new RowColumn(leadingBlank / width * 100D));
            var contentColumn = new RowColumn(contentWidth / width * 100D);
            if (style.PaddingY > 0D) contentColumn.AddBlock(new SpacerBlock(style.PaddingY));
            foreach (IPdfBlock block in container.Blocks) contentColumn.AddBlock(block);
            if (style.PaddingY > 0D) contentColumn.AddBlock(new SpacerBlock(style.PaddingY));
            row.AddColumn(contentColumn);
            if (trailingBlank > 0.001D) row.AddColumn(new RowColumn(trailingBlank / width * 100D));

            double rowHeight = MeasureRowBlockHeight(row, currentOpts.MarginLeft, width, currentOpts.DefaultFontSize, firstVisualOnly: false);
            double fullPageHeight = currentOpts.PageHeight - currentOpts.MarginTop - currentOpts.MarginBottom;
            if (rowHeight > fullPageHeight + 0.001D) throw new ArgumentException("Container height exceeds the available page content height.");
            if (y < yStart - 0.001D && y - rowHeight < currentOpts.MarginBottom) NewPage();

            double spacingBefore = ResolveTopLevelSpacingBefore(style.SpacingBefore);
            double outerTop = y - spacingBefore;
            int insertionIndex = sb.Length;
            RenderRowFlowBlock(row, nextBlock: null, new List<IPdfBlock> { row }, 0);
            double outerBottom = y + style.SpacingAfter;
            var decoration = new StringBuilder();
            if (style.Background.HasValue) DrawRowFill(decoration, style.Background.Value, outerX, outerBottom, outerWidth, outerTop - outerBottom, emitGeneratedStructure);
            DrawPanelBorder(decoration, style, outerX, outerBottom, outerWidth, outerTop - outerBottom, emitGeneratedStructure);
            if (decoration.Length > 0) {
                sb.Insert(insertionIndex, decoration.ToString());
                pageDirty = true;
            }
        }
    }
}
