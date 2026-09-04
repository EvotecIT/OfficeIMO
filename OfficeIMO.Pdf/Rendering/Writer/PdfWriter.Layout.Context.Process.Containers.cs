namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private void RenderMultiColumnBlock(MultiColumnBlock columns) {
            PdfMultiColumnOptions options = columns.Options;
            double totalGap = options.Gap * (options.ColumnCount - 1);
            if (totalGap >= width) throw new ArgumentException("Multi-column gaps must leave positive column widths.");
            double columnWidth = (width - totalGap) / options.ColumnCount;
            ValidateMultiColumnBlocks(columns.Blocks);
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
                    var column = new RowColumn(PdfColumnWidth.Percent(widthPercent));
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

        private static List<PdfTextRun> BuildTextRunsFromWrappedLines(
            IReadOnlyList<List<RichSeg>> lines,
            int start,
            int count) {
            var runs = new List<PdfTextRun>();
            for (int lineIndex = 0; lineIndex < count; lineIndex++) {
                IReadOnlyList<RichSeg> line = lines[start + lineIndex];
                for (int segmentIndex = 0; segmentIndex < line.Count; segmentIndex++) {
                    RichSeg segment = line[segmentIndex];
                    if (segment.InlineElement != null) {
                        if (segment.LeadingSpace) {
                            runs.Add(BuildTextRunFromWrappedSegment(" ", segment));
                        }

                        runs.Add(PdfTextRun.Inline(segment.InlineElement));
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
                        runs.Add(PdfTextRun.LineBreak());
                    } else if (line[line.Count - 1].EndsWithTextSeparator) {
                        runs.Add(BuildTextRunFromWrappedSegment(" ", line[line.Count - 1].WithoutLink()));
                    }
                }
            }

            return runs;
        }

        private static PdfTextRun BuildTextRunFromWrappedSegment(string text, RichSeg segment) =>
            new PdfTextRun(
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
                fontFamily: segment.NamedFont?.FamilyName,
                underlineStyle: segment.UnderlineStyle,
                strikeStyle: segment.StrikeStyle);

        private double MeasureColumnBlock(IPdfBlock block, double columnWidth) =>
            MeasureKeepWithNextBlockHeight(block, currentOpts.MarginLeft, columnWidth, currentOpts.DefaultFontSize);

        private static void ValidateMultiColumnBlocks(IReadOnlyList<IPdfBlock> blocks) {
            foreach (IPdfBlock block in blocks) {
                if (PdfFlowNestingRules.IsColumnFlowPrimitive(block) || block is ColumnBreakBlock) {
                    continue;
                }

                throw new NotSupportedException("Automatic multi-column flow does not support nested block type " + block.GetType().Name + ". Use separate Columns blocks around that content.");
            }
        }

        private void RenderContainerBlock(
            ContainerBlock container,
            IPdfBlock? nextBlock,
            System.Collections.Generic.IList<IPdfBlock> blockList,
            int blockIndex) {
            PdfPanelStyle style = container.Style;
            double parentLeft = currentOpts.MarginLeft;
            double parentWidth = width;
            double outerWidth = style.MaxWidth.HasValue ? Math.Min(parentWidth, style.MaxWidth.Value) : parentWidth;
            ValidatePanelStyle(style, outerWidth);
            double outerX = style.Align switch {
                PdfAlign.Center => parentLeft + (parentWidth - outerWidth) / 2D,
                PdfAlign.Right => parentLeft + parentWidth - outerWidth,
                _ => parentLeft
            };
            double contentWidth = outerWidth - 2D * style.PaddingX;
            if (contentWidth <= 0.001D) {
                throw new ArgumentException("Container padding must leave positive content width.");
            }

            double spacingBefore = ResolveTopLevelSpacingBefore(style.SpacingBefore);
            double firstVisualHeight = container.Blocks.Count == 0
                ? 0D
                : MeasureNextBlockFirstVisualHeight(container.Blocks[0], outerX + style.PaddingX, contentWidth, currentOpts.DefaultFontSize);
            double minimumStartHeight = spacingBefore + style.PaddingY + firstVisualHeight;
            if (y < yStart - 0.001D && y - minimumStartHeight < currentOpts.MarginBottom) {
                NewPage();
                spacingBefore = ResolveTopLevelSpacingBefore(style.SpacingBefore);
            }

            if (style.KeepTogether) {
                double? keepHeight = MeasureWholeBlockHeight(container, parentLeft, parentWidth, currentOpts.DefaultFontSize);
                double? fullPageKeepHeight = MeasureWholeBlockAtFrameStart(container, parentLeft, parentWidth, currentOpts.DefaultFontSize);
                if (!keepHeight.HasValue || !fullPageKeepHeight.HasValue) {
                    throw new NotSupportedException("KeepTogether requires element content whose height can be determined before rendering. Remove KeepTogether or move dynamic, multi-column, deferred-table, table-of-contents, or explicit page-boundary content outside the element.");
                }

                double fullPageHeight = GetCurrentFramePageStartY() - currentOpts.MarginBottom;
                if (fullPageKeepHeight.Value > fullPageHeight + 0.001D) {
                    throw new ArgumentException("Container height exceeds the available page content height while KeepTogether is enabled.");
                }

                if (y < yStart - 0.001D && y - keepHeight.Value < currentOpts.MarginBottom) {
                    NewPage();
                    spacingBefore = ResolveTopLevelSpacingBefore(style.SpacingBefore);
                }
            } else if (style.KeepWithNext && nextBlock != null) {
                double? elementHeight = MeasureWholeBlockHeight(container, parentLeft, parentWidth, currentOpts.DefaultFontSize);
                if (!elementHeight.HasValue) {
                    throw new NotSupportedException("KeepWithNext requires element content whose height can be determined before rendering. Remove KeepWithNext or move dynamic, multi-column, deferred-table, table-of-contents, canvas, or explicit page-boundary content outside the element.");
                }

                double nextHeight = MeasureKeepWithNextChainHeight(blockList, blockIndex + 1, parentLeft, parentWidth, currentOpts.DefaultFontSize, elementHeight.Value);
                double keepHeight = elementHeight.Value + nextHeight;
                double fullPageHeight = GetCurrentFramePageStartY() - currentOpts.MarginBottom;
                if (nextHeight > 0.001D && keepHeight <= fullPageHeight + 0.001D && y < yStart - 0.001D && y - keepHeight < currentOpts.MarginBottom) {
                    NewPage();
                    spacingBefore = ResolveTopLevelSpacingBefore(style.SpacingBefore);
                }
            }

            y -= spacingBefore;
            PdfOptions parentOptions = currentOpts;
            double parentYStart = yStart;
            PdfOptions pageOptions = currentPage!.Options;
            var nestedOptions = currentOpts.Clone();
            nestedOptions.MarginLeft = outerX + style.PaddingX;
            nestedOptions.MarginRight = nestedOptions.PageWidth - (outerX + outerWidth - style.PaddingX);
            nestedOptions.Validate();

            var scope = new ContainerRenderScope(style, outerX, outerWidth, pageOptions);
            activeContainerScopes.Add(scope);
            currentOpts = nestedOptions;
            width = contentWidth;
            BeginContainerFragment(scope);
            try {
                ProcessBlocks(container.Blocks);
                double bottomPadding = Math.Min(style.PaddingY, Math.Max(0D, y - currentOpts.MarginBottom));
                y -= bottomPadding;
                FinalizeContainerFragment(scope);
            } finally {
                activeContainerScopes.RemoveAt(activeContainerScopes.Count - 1);
                currentOpts = parentOptions;
                width = parentWidth;
                yStart = parentYStart;
                if (currentPage != null) {
                    currentPage.Options = pageOptions;
                }
            }

            if (style.SpacingAfter > 0D) {
                ConsumeSpacer(style.SpacingAfter);
            }
        }

        private void PrepareActiveContainerScopesForPageBreak() {
            if (activeContainerScopes.Count == 0 || currentPage == null) {
                return;
            }

            for (int index = activeContainerScopes.Count - 1; index >= 0; index--) {
                ContainerRenderScope scope = activeContainerScopes[index];
                double bottomPadding = Math.Min(scope.Style.PaddingY, Math.Max(0D, y - currentOpts.MarginBottom));
                y -= bottomPadding;
                FinalizeContainerFragment(scope);
            }
        }

        private void ResumeActiveContainerScopesOnNewPage() {
            if (activeContainerScopes.Count == 0 || currentPage == null) {
                return;
            }

            currentPage.Options = activeContainerScopes[0].PageOptions;
            for (int index = 0; index < activeContainerScopes.Count; index++) {
                BeginContainerFragment(activeContainerScopes[index]);
            }
        }

        private void BeginContainerFragment(ContainerRenderScope scope) {
            scope.InsertionIndex = sb.Length;
            scope.FragmentTop = y;
            y -= Math.Min(scope.Style.PaddingY, Math.Max(0D, y - currentOpts.MarginBottom));
        }

        private void FinalizeContainerFragment(ContainerRenderScope scope) {
            double fragmentHeight = scope.FragmentTop - y;
            if (fragmentHeight <= 0.001D) {
                return;
            }

            var decoration = new StringBuilder();
            if (scope.Style.Background.HasValue) {
                DrawRowFill(decoration, scope.Style.Background.Value, scope.OuterX, y, scope.OuterWidth, fragmentHeight, emitGeneratedStructure);
            }

            DrawPanelBorder(decoration, scope.Style, scope.OuterX, y, scope.OuterWidth, fragmentHeight, emitGeneratedStructure);
            if (decoration.Length > 0) {
                sb.Insert(scope.InsertionIndex, decoration.ToString());
                pageDirty = true;
            }
        }

        private sealed class ContainerRenderScope {
            public ContainerRenderScope(PdfPanelStyle style, double outerX, double outerWidth, PdfOptions pageOptions) {
                Style = style;
                OuterX = outerX;
                OuterWidth = outerWidth;
                PageOptions = pageOptions;
            }

            public PdfPanelStyle Style { get; }
            public double OuterX { get; }
            public double OuterWidth { get; }
            public PdfOptions PageOptions { get; }
            public int InsertionIndex { get; set; }
            public double FragmentTop { get; set; }
        }
    }
}
