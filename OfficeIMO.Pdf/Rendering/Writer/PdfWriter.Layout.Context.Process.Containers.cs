namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private void RenderMultiColumnBlock(MultiColumnBlock columns) {
            PdfMultiColumnOptions options = columns.Options;
            double totalGap = options.Gap * (options.ColumnCount - 1);
            if (totalGap >= width) throw new ArgumentException("Multi-column gaps must leave positive column widths.");
            double columnWidth = (width - totalGap) / options.ColumnCount;
            ValidateColumnBlocks(columns.Blocks);
            var pendingBlocks = columns.Blocks.ToList();
            int blockIndex = 0;
            while (blockIndex < pendingBlocks.Count) {
                while (blockIndex < pendingBlocks.Count && pendingBlocks[blockIndex] is ColumnBreakBlock) blockIndex++;
                if (blockIndex >= pendingBlocks.Count) break;
                double availableHeight = y - currentOpts.MarginBottom;
                if (availableHeight < currentOpts.DefaultFontSize * 1.4D) {
                    NewPage();
                    availableHeight = y - currentOpts.MarginBottom;
                }

                double remainingHeight = 0D;
                for (int i = blockIndex; i < pendingBlocks.Count; i++) {
                    if (pendingBlocks[i] is not ColumnBreakBlock) remainingHeight += MeasureColumnBlock(pendingBlocks[i], columnWidth);
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
                    while (blockIndex < pendingBlocks.Count) {
                        IPdfBlock block = pendingBlocks[blockIndex];
                        if (block is ColumnBreakBlock) {
                            blockIndex++;
                            break;
                        }

                        double blockHeight = MeasureColumnBlock(block, columnWidth);
                        double remainingTarget = target - consumed;
                        if (options.BalanceParagraphLines &&
                            block is RichParagraphBlock paragraph &&
                            blockHeight > remainingTarget + 0.001D &&
                            TrySplitColumnParagraph(paragraph, columnWidth, remainingTarget, out RichParagraphBlock? first, out RichParagraphBlock? remainder)) {
                            column.AddBlock(first!);
                            consumed += MeasureColumnBlock(first!, columnWidth);
                            pendingBlocks[blockIndex] = remainder!;
                            break;
                        }

                        if (column.Blocks.Count > 0 && consumed + blockHeight > target + 0.001D) break;
                        column.AddBlock(block);
                        consumed += blockHeight;
                        blockIndex++;
                        if (consumed >= target - 0.001D) break;
                    }
                    row.AddColumn(column);
                }

                RenderRowFlowBlock(row, nextBlock: null, new List<IPdfBlock> { row }, 0);
                if (blockIndex < pendingBlocks.Count) NewPage();
            }
        }

        private bool TrySplitColumnParagraph(
            RichParagraphBlock paragraph,
            double columnWidth,
            double availableHeight,
            out RichParagraphBlock? first,
            out RichParagraphBlock? remainder) {
            first = null;
            remainder = null;
            PdfParagraphStyle? sourceStyle = EffectiveParagraphStyle(paragraph);
            if (availableHeight <= 0.001D ||
                sourceStyle?.KeepTogether == true ||
                paragraph.Runs.Any(run => run.Text.Contains('\t'))) {
                return false;
            }

            double fontSize = currentOpts.DefaultFontSize;
            double leading = GetParagraphLeading(sourceStyle, fontSize);
            var textFrame = GetParagraphTextFrame(sourceStyle, currentOpts.MarginLeft, columnWidth);
            var wrapped = WrapRichRunsCoreWithFirstLineOrigin(
                paragraph.Runs,
                textFrame.Width,
                fontSize,
                ChooseNormal(currentOpts.DefaultFont),
                leading,
                textFrame.FirstLineWidth,
                textFrame.FirstLineX - textFrame.X,
                GetParagraphTabStopWidth(sourceStyle),
                currentOpts,
                GetParagraphTabStops(sourceStyle));
            if (wrapped.Lines.Count < 2) {
                return false;
            }

            double remainingHeight = Math.Max(0D, availableHeight - GetParagraphSpacingBefore(sourceStyle));
            int take = 0;
            double height = 0D;
            for (int index = 0; index < wrapped.LineHeights.Count; index++) {
                if (height + wrapped.LineHeights[index] > remainingHeight + 0.001D) {
                    break;
                }

                height += wrapped.LineHeights[index];
                take++;
            }

            int minimumOrphanLines = ResolveMinimumOrphanLines(sourceStyle ?? new PdfParagraphStyle());
            int minimumWidowLines = ResolveMinimumWidowLines(sourceStyle ?? new PdfParagraphStyle());
            if (take < Math.Max(1, minimumOrphanLines) || wrapped.Lines.Count - take < Math.Max(1, minimumWidowLines)) {
                return false;
            }

            PdfParagraphStyle firstStyle = sourceStyle?.Clone() ?? new PdfParagraphStyle();
            firstStyle.KeepTogether = false;
            firstStyle.KeepWithNext = false;
            firstStyle.WidowControl = false;
            firstStyle.MinimumOrphanLines = 0;
            firstStyle.MinimumWidowLines = 0;
            firstStyle.SpacingAfter = 0D;

            PdfParagraphStyle remainderStyle = sourceStyle?.Clone() ?? new PdfParagraphStyle();
            remainderStyle.KeepTogether = false;
            remainderStyle.WidowControl = false;
            remainderStyle.MinimumOrphanLines = 0;
            remainderStyle.MinimumWidowLines = 0;
            remainderStyle.SpacingBefore = 0D;
            remainderStyle.FirstLineIndent = 0D;

            first = new RichParagraphBlock(BuildTextRunsFromWrappedLines(wrapped.Lines, 0, take), paragraph.Align, paragraph.DefaultColor, firstStyle);
            remainder = new RichParagraphBlock(BuildTextRunsFromWrappedLines(wrapped.Lines, take, wrapped.Lines.Count - take), paragraph.Align, paragraph.DefaultColor, remainderStyle);
            return true;
        }

        private static List<TextRun> BuildTextRunsFromWrappedLines(
            IReadOnlyList<List<RichSeg>> lines,
            int start,
            int count) {
            var runs = new List<TextRun>();
            for (int lineIndex = 0; lineIndex < count; lineIndex++) {
                IReadOnlyList<RichSeg> line = lines[start + lineIndex];
                for (int segmentIndex = 0; segmentIndex < line.Count; segmentIndex++) {
                    RichSeg segment = line[segmentIndex];
                    if (segment.InlineElement != null) {
                        if (segment.LeadingSpace) {
                            runs.Add(BuildTextRunFromWrappedSegment(" ", segment));
                        }

                        runs.Add(TextRun.Inline(segment.InlineElement));
                        continue;
                    }

                    string text = (segment.LeadingSpace ? " " : string.Empty) + segment.Text;
                    if (text.Length == 0) {
                        continue;
                    }

                    runs.Add(BuildTextRunFromWrappedSegment(text, segment));
                }

                if (lineIndex + 1 < count) {
                    if (line.Count == 0 || line[line.Count - 1].EndsWithHardBreak) {
                        runs.Add(TextRun.LineBreak());
                    } else if (line[line.Count - 1].EndsWithTextSeparator) {
                        runs.Add(BuildTextRunFromWrappedSegment(" ", line[line.Count - 1].WithoutLink()));
                    }
                }
            }

            return runs;
        }

        private static TextRun BuildTextRunFromWrappedSegment(string text, RichSeg segment) =>
            new TextRun(
                text,
                segment.Bold,
                segment.Underline,
                segment.Color,
                segment.Italic,
                segment.Strike,
                segment.FontSize,
                segment.Font,
                segment.Uri,
                segment.Contents,
                segment.Baseline,
                segment.DestinationName,
                backgroundColor: segment.BackgroundColor,
                fontFamily: segment.NamedFont?.FamilyName);

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
                KeepTogether = style.KeepTogether,
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
            if (style.KeepTogether && rowHeight > fullPageHeight + 0.001D) {
                throw new ArgumentException("Container height exceeds the available page content height while KeepTogether is enabled.");
            }

            void DecorateFragment(
                StringBuilder content,
                int insertionIndex,
                double top,
                double bottom,
                bool isFirstFragment,
                bool isLastFragment) {
                var decoration = new StringBuilder();
                if (style.Background.HasValue) {
                    DrawRowFill(decoration, style.Background.Value, outerX, bottom, outerWidth, top - bottom, emitGeneratedStructure);
                }

                DrawPanelBorder(decoration, style, outerX, bottom, outerWidth, top - bottom, emitGeneratedStructure);
                if (decoration.Length > 0) {
                    content.Insert(insertionIndex, decoration.ToString());
                    pageDirty = true;
                }
            }

            RenderRowFlowBlock(row, nextBlock: null, new List<IPdfBlock> { row }, 0, DecorateFragment);
        }
    }
}
