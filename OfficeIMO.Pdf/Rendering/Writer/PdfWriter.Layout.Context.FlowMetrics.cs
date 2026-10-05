using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private void EnsureFixedFlowBlockFits(string blockName, double blockWidth, double blockHeight, double availableWidth, double reservedHeight = 0D) {
            if (blockWidth > availableWidth + 0.001) {
                throw new ArgumentException(blockName + " width exceeds the available page content width.");
            }

            double availableHeight = GetMaximumBlockContinuationHeight() - reservedHeight;
            if (blockHeight > availableHeight + 0.001) {
                throw new ArgumentException(blockName + " height exceeds the available page content height.");
            }
        }

        private (double Width, double Height) ResolveImageFlowBox(ImageBlock image, PdfImageStyle style, double frameWidth, double spacingBefore, double spacingAfter, double reservedHeight = 0D) {
            double imageWidth = image.Width;
            double imageHeight = image.Height;
            if (!style.ScaleDownToFit) {
                return (imageWidth, imageHeight);
            }

            double availableHeight = GetMaximumBlockContinuationHeight() - reservedHeight - imageMeasurementReservedHeight
                - activeContainerScopes.Sum(scope => scope.Style.FragmentPaddingReservation) - spacingBefore - spacingAfter;
            double scale = 1D;
            if (imageWidth > frameWidth) {
                scale = Math.Min(scale, frameWidth / imageWidth);
            }

            if (availableHeight > 0D && imageHeight * scale > availableHeight) {
                scale = Math.Min(scale, availableHeight / imageHeight);
            }

            if (scale >= 1D) {
                return (imageWidth, imageHeight);
            }

            if (scale <= 0D || double.IsNaN(scale) || double.IsInfinity(scale)) {
                return (imageWidth, imageHeight);
            }

            return (imageWidth * scale, imageHeight * scale);
        }

        private static void ValidateHorizontalRule(PdfHorizontalRuleStyle rule) {
            if (rule.Thickness <= 0 || double.IsNaN(rule.Thickness) || double.IsInfinity(rule.Thickness)) {
                throw new ArgumentException("Horizontal rule thickness must be a positive finite value.");
            }

            if (rule.SpacingBefore < 0 || double.IsNaN(rule.SpacingBefore) || double.IsInfinity(rule.SpacingBefore)) {
                throw new ArgumentException("Horizontal rule spacing before must be a non-negative finite value.");
            }

            if (rule.SpacingAfter < 0 || double.IsNaN(rule.SpacingAfter) || double.IsInfinity(rule.SpacingAfter)) {
                throw new ArgumentException("Horizontal rule spacing after must be a non-negative finite value.");
            }
        }

        private static void ValidatePanelStyle(PdfPanelStyle style, double panelWidth) {
            Guard.LeftCenterRightAlign(style.Align, nameof(style.Align), "Panel box");

            if (style.BorderWidth < 0 || double.IsNaN(style.BorderWidth) || double.IsInfinity(style.BorderWidth)) {
                throw new ArgumentException("Panel border width must be a non-negative finite value.");
            }

            if (style.PaddingX < 0 || double.IsNaN(style.PaddingX) || double.IsInfinity(style.PaddingX)) {
                throw new ArgumentException("Panel horizontal padding must be a non-negative finite value.");
            }

            if (style.PaddingY < 0 || double.IsNaN(style.PaddingY) || double.IsInfinity(style.PaddingY)) {
                throw new ArgumentException("Panel vertical padding must be a non-negative finite value.");
            }

            if (style.MaxWidth.HasValue && (style.MaxWidth.Value <= 0 || double.IsNaN(style.MaxWidth.Value) || double.IsInfinity(style.MaxWidth.Value))) {
                throw new ArgumentException("Panel maximum width must be a positive finite value.");
            }

            if (style.SpacingBefore < 0 || double.IsNaN(style.SpacingBefore) || double.IsInfinity(style.SpacingBefore)) {
                throw new ArgumentException("Panel spacing before must be a non-negative finite value.");
            }

            if (style.SpacingAfter < 0 || double.IsNaN(style.SpacingAfter) || double.IsInfinity(style.SpacingAfter)) {
                throw new ArgumentException("Panel spacing after must be a non-negative finite value.");
            }

            if (panelWidth - 2 * style.PaddingX <= 0) {
                throw new ArgumentException("Panel horizontal padding must leave a positive text width.");
            }
        }

        private void EnsurePanelSegmentCanFitLine(double topPadding, double lineHeight) {
            double availableHeight = currentOpts.PageHeight - currentOpts.MarginTop - currentOpts.MarginBottom;
            if (topPadding + lineHeight > availableHeight + 0.001D) {
                throw new ArgumentException("Panel vertical padding and first line height exceed the available page content height.");
            }
        }

        private PdfParagraphStyle? EffectiveParagraphStyle(RichParagraphBlock paragraph) => paragraph.Style ?? currentOpts.DefaultParagraphStyleSnapshot;

        private double MeasureNextParagraphFirstVisualHeight(RichParagraphBlock paragraph, double frameX, double frameWidth, double fontSize,
            bool suppressSpacingBefore = false) {
            PdfParagraphStyle? paragraphStyle = EffectiveParagraphStyle(paragraph);
            fontSize = paragraphStyle?.FontSize ?? currentOpts.DefaultFontSize;
            double leading = GetParagraphLeading(paragraphStyle, fontSize);
            double spacingBefore = suppressSpacingBefore ? 0D : GetParagraphSpacingBefore(paragraphStyle);
            var textFrame = GetParagraphTextFrame(paragraphStyle, frameX, frameWidth);
            var wrap = WrapRichRunsCoreWithFirstLineOrigin(paragraph.Runs, textFrame.Width, fontSize, ChooseNormal(currentOpts.DefaultFont), leading, textFrame.FirstLineWidth, textFrame.FirstLineX - textFrame.X, GetParagraphTabStopWidth(paragraphStyle), currentOpts, paragraphStyle?.TabStops.ToArray(), lineSpacing: paragraphStyle?.LineSpacing);
            if (wrap.LineHeights.Count == 0) {
                return spacingBefore;
            }

            int linesToReserve = 1;
            if (paragraphStyle?.KeepTogether == true) {
                linesToReserve = wrap.LineHeights.Count;
            } else if (paragraphStyle != null) {
                int minimumOrphanLines = ResolveMinimumOrphanLines(paragraphStyle);
                if (minimumOrphanLines > 1 && wrap.LineHeights.Count > 1) {
                    linesToReserve = Math.Min(minimumOrphanLines, wrap.LineHeights.Count);
                }
            }

            double height = spacingBefore;
            for (int i = 0; i < linesToReserve; i++) {
                height += GetRichLineHeight(wrap.LineHeights, i, leading);
            }

            return height;
        }

        private const int MaxKeepWithNextChainBlocks = 256;

        private double MeasureKeepWithNextChainHeight(System.Collections.Generic.IList<IPdfBlock> blocks, int startIndex, double frameX, double frameWidth, double fontSize, double precedingHeight) {
            double savedY = y;
            var savedFloats = floatingTables.ToArray();
            try {
                y -= precedingHeight;
                double startY = y;
                double deepest = y;
                int inspectedBlocks = 0;
                for (int blockIndex = startIndex; blockIndex < blocks.Count; blockIndex++) {
                    IPdfBlock block = blocks[blockIndex];
                    if (IsNonVisualFlowMarker(block)) {
                        continue;
                    }
                    if (inspectedBlocks >= MaxKeepWithNextChainBlocks) {
                        throw new NotSupportedException("KeepWithNext chains cannot contain more than " + MaxKeepWithNextChainBlocks + " visual blocks.");
                    }
                    inspectedBlocks++;

                    bool keepWithNext = KeepsWithNext(block);
                    if (block is TableBlock floating && TryMeasureFloatingTable(floating, frameWidth, fontSize, out double floatingBottom)) {
                        deepest = Math.Min(deepest, floatingBottom);
                        if (!keepWithNext) break;
                        continue;
                    }
                    double blockHeight;
                    if (keepWithNext) {
                        double? measured = MeasureWholeBlockHeight(block, frameX, frameWidth, fontSize);
                        if (!measured.HasValue) {
                            throw new NotSupportedException("KeepWithNext requires content whose height can be determined before rendering. Remove KeepWithNext or move dynamic, multi-column, deferred-table, table-of-contents, canvas, or explicit page-boundary content outside the element.");
                        }

                        blockHeight = measured.Value;
                    } else {
                        blockHeight = MeasureNextBlockFirstVisualHeight(block, frameX, frameWidth, fontSize);
                    }

                    if (floatingTables.Count > savedFloats.Length) {
                        if (block is RichParagraphBlock paragraph)
                            blockHeight = MeasureFloatingParagraph(paragraph, frameX, frameWidth, fontSize, firstVisualOnly: !keepWithNext);
                        else AvoidFloatingBlock(blockHeight);
                    }
                    y -= blockHeight;
                    deepest = Math.Min(deepest, y);
                    if (!keepWithNext) {
                        break;
                    }
                }

                return startY - deepest;
            } finally {
                y = savedY;
                floatingTables.Clear(); floatingTables.AddRange(savedFloats);
            }
        }

        private static bool IsNonVisualFlowMarker(IPdfBlock block) =>
            block is BookmarkBlock;

        private double MeasureKeepWithNextBlockHeight(IPdfBlock block, double frameX, double frameWidth, double fontSize) {
            if (block is HeadingBlock heading) {
                return MeasureHeadingBlockHeight(heading, frameWidth);
            }

            if (block is RichParagraphBlock paragraph) {
                return MeasureParagraphBlockHeight(paragraph, frameX, frameWidth, fontSize);
            }

            if (block is PdfListBlock list) {
                return MeasureListBlockHeight(list, frameWidth, fontSize);
            }

            if (block is TableBlock table) {
                return MeasureTableBlockHeight(table, frameWidth, fontSize, firstVisualOnly: false);
            }

            if (block is HorizontalRuleBlock rule) {
                PdfHorizontalRuleStyle style = ResolveHorizontalRuleStyle(rule, currentOpts);
                return ResolveTopLevelSpacingBefore(style.SpacingBefore) + style.Thickness + style.SpacingAfter;
            }

            if (block is ImageBlock image) {
                return MeasureImageBlockHeight(image, frameWidth);
            }

            if (block is ShapeBlock shape) {
                return MeasureShapeBlockHeight(shape);
            }

            if (block is DrawingBlock drawing) {
                return MeasureDrawingBlockHeight(drawing);
            }


            if (block is RowBlock row) {
                return MeasureRowBlockHeight(row, frameX, frameWidth, fontSize, firstVisualOnly: false);
            }

            return MeasureNextBlockFirstVisualHeight(block, frameX, frameWidth, fontSize);
        }

        private double MeasureHeadingBlockHeight(HeadingBlock heading, double frameWidth) {
            PdfHeadingStyle? headingStyle = ResolveHeadingStyle(heading, currentOpts);
            double headingSize = GetHeadingFontSize(heading, headingStyle);
            double headingLeading = GetHeadingLeading(headingStyle, headingSize);
            double spacingBefore = y < GetCurrentFramePageStartY() - 0.001D || headingStyle?.ApplySpacingBeforeAtTop == true
                ? headingStyle?.SpacingBefore ?? 0D
                : 0D;
            double spacingAfter = GetHeadingSpacingAfter(headingStyle, headingLeading);
            PdfColor? headingColor = heading.Color ?? headingStyle?.Color;
            System.Collections.Generic.IReadOnlyList<PdfTextRun> headingRuns = CreateHeadingTextRuns(heading, headingStyle, headingColor);
            var wrap = WrapRichRunsWithSpacing(headingRuns, frameWidth, headingSize, ChooseNormal(currentOpts.DefaultFont), headingLeading, null, DefaultParagraphTabStopWidth, currentOpts, headingStyle?.LineSpacing);
            return spacingBefore + MeasureRichLinesHeight(wrap.LineHeights, wrap.Lines.Count, headingLeading) + spacingAfter;
        }

        private double MeasureParagraphBlockHeight(RichParagraphBlock paragraph, double frameX, double frameWidth, double fontSize) {
            PdfParagraphStyle? paragraphStyle = EffectiveParagraphStyle(paragraph);
            fontSize = paragraphStyle?.FontSize ?? currentOpts.DefaultFontSize;
            double leading = GetParagraphLeading(paragraphStyle, fontSize);
            double spacingBefore = ResolveTopLevelSpacingBefore(GetParagraphSpacingBefore(paragraphStyle));
            double spacingAfter = GetParagraphSpacingAfter(paragraphStyle, leading);
            var textFrame = GetParagraphTextFrame(paragraphStyle, frameX, frameWidth);
            var wrap = WrapRichRunsCoreWithFirstLineOrigin(paragraph.Runs, textFrame.Width, fontSize, ChooseNormal(currentOpts.DefaultFont), leading, textFrame.FirstLineWidth, textFrame.FirstLineX - textFrame.X, GetParagraphTabStopWidth(paragraphStyle), currentOpts, paragraphStyle?.TabStops.ToArray(), lineSpacing: paragraphStyle?.LineSpacing);
            return spacingBefore + wrap.LineHeights.Sum() + spacingAfter;
        }

        private double MeasureListBlockHeight(PdfListBlock list, double frameWidth, double fontSize) {
            PreparedListLayout prepared = PrepareListLayout(list, frameWidth, fontSize, topLevelSpacing: true);
            return MeasurePreparedListHeight(prepared);
        }

        private bool KeepsWithNext(IPdfBlock block) {
            if (block is HeadingBlock heading) {
                return ResolveHeadingStyle(heading, currentOpts)?.KeepWithNext ?? true;
            }

            if (block is RichParagraphBlock paragraph) {
                return EffectiveParagraphStyle(paragraph)?.KeepWithNext == true;
            }

            if (block is PdfListBlock list) {
                return ResolveListStyle(list, currentOpts)?.KeepWithNext == true;
            }

            if (block is TableBlock table) {
                return (table.Style ?? currentOpts.DefaultTableStyleSnapshot ?? TableStyles.Light()).KeepWithNext;
            }

            if (block is HorizontalRuleBlock rule) {
                return ResolveHorizontalRuleStyle(rule, currentOpts).KeepWithNext;
            }

            if (block is ImageBlock image) {
                return ResolveImageStyle(image, currentOpts).KeepWithNext;
            }

            if (block is ShapeBlock shape) {
                return ResolveDrawingStyle(shape, currentOpts).KeepWithNext;
            }

            if (block is DrawingBlock drawing) {
                return ResolveDrawingStyle(drawing, currentOpts).KeepWithNext;
            }


            if (block is RowBlock row) {
                return (row.StyleSnapshot ?? currentOpts.DefaultRowStyleSnapshot)?.KeepWithNext == true;
            }

            if (block is ContainerBlock container) {
                return ResolveContainerStyle(container).KeepWithNext;
            }

            return false;
        }

        private double MeasureNextBlockFirstVisualHeight(IPdfBlock block, double frameX, double frameWidth, double fontSize,
            bool allowTableFragments = false, bool suppressParagraphSpacingBefore = false) {
            if (block is SemanticBlock semantic) {
                return MeasureFirstNestedVisualHeight(semantic.Blocks, frameX, frameWidth, fontSize, allowTableFragments, suppressParagraphSpacingBefore);
            }

            if (block is LayerBlock layer) {
                return MeasureFirstNestedVisualHeight(layer.Blocks, frameX, frameWidth, fontSize, allowTableFragments, suppressParagraphSpacingBefore);
            }

            if (block is SectionBlock section) {
                if (section.Options.StartOnNewPage) {
                    return 0D;
                }

                if (section.Options.IncludeHeading) {
                    var sectionHeading = new HeadingBlock(section.Options.Level, section.Title, PdfAlign.Left, color: null, style: section.Options.HeadingStyle);
                    return MeasureNextBlockFirstVisualHeight(sectionHeading, frameX, frameWidth, fontSize);
                }

                return MeasureFirstNestedVisualHeight(section.Blocks, frameX, frameWidth, fontSize, allowTableFragments, suppressParagraphSpacingBefore);
            }

            if (block is ContainerBlock container) {
                PdfPanelStyle style = ResolveContainerStyle(container);
                double outerWidth = style.MaxWidth.HasValue ? Math.Min(frameWidth, style.MaxWidth.Value) : frameWidth;
                ValidatePanelStyle(style, outerWidth);
                double contentWidth = outerWidth - 2D * style.PaddingX;
                if (contentWidth <= 0.001D) {
                    throw new ArgumentException("Container padding must leave positive content width.");
                }

                return ResolveTopLevelSpacingBefore(style.SpacingBefore) + style.PaddingY +
                       MeasureWithContainerPaddingReservation(style.FragmentPaddingReservation, () =>
                           MeasureFirstNestedVisualHeight(container.Blocks, frameX + style.PaddingX, contentWidth, fontSize, allowTableFragments, suppressParagraphSpacingBefore));
            }

            if (block is FlowBlock flow && !flow.IsReplayable && flow.Options.ShowIf == null && flow.StaticBlocks != null) {
                return MeasureFirstNestedVisualHeight(flow.StaticBlocks, frameX, frameWidth, fontSize, allowTableFragments, suppressParagraphSpacingBefore);
            }

            if (block is MultiColumnBlock columns) {
                double totalGap = columns.Options.Gap * (columns.Options.ColumnCount - 1);
                if (totalGap >= frameWidth) {
                    return 0D;
                }

                double columnWidth = (frameWidth - totalGap) / columns.Options.ColumnCount;
                return MeasureFirstNestedVisualHeight(columns.Blocks, frameX, columnWidth, fontSize);
            }

            if (block is RichParagraphBlock paragraph) {
                return MeasureNextParagraphFirstVisualHeight(paragraph, frameX, frameWidth, fontSize, suppressParagraphSpacingBefore);
            }

            if (block is HeadingBlock heading) {
                PdfHeadingStyle? headingStyle = ResolveHeadingStyle(heading, currentOpts);
                double headingSize = GetHeadingFontSize(heading, headingStyle);
                double headingLeading = GetHeadingLeading(headingStyle, headingSize);
                double spacingBefore = y < GetCurrentFramePageStartY() - 0.001D || headingStyle?.ApplySpacingBeforeAtTop == true
                    ? headingStyle?.SpacingBefore ?? 0D
                    : 0D;
                return spacingBefore + headingLeading;
            }

            if (block is SpacerBlock spacer) {
                return spacer.Height;
            }

            if (block is PdfListBlock list) {
                PdfListStyle? listStyle = ResolveListStyle(list, currentOpts);
                double size = GetListFontSize(listStyle, fontSize);
                double markerSize = GetListMarkerFontSize(listStyle, size);
                double leading = Math.Max(GetListLeading(listStyle, size), GetListLeading(listStyle, markerSize));
                string? firstItem = list.Items.Count > 0 ? list.Items[0] : null;
                if (firstItem == null) {
                    return listStyle?.SpacingBefore ?? 0D;
                }

                return (listStyle?.SpacingBefore ?? 0D) + leading;
            }


            if (block is TableBlock table) {
                return MeasureTableBlockHeight(table, frameWidth, fontSize, firstVisualOnly: true, allowTableFragments);
            }

            if (block is DeferredTableBlock deferredTable) {
                return MeasureDeferredTableFirstVisualHeight(deferredTable, frameWidth, fontSize);
            }

            if (block is HorizontalRuleBlock rule) {
                PdfHorizontalRuleStyle style = ResolveHorizontalRuleStyle(rule, currentOpts);
                return ResolveTopLevelSpacingBefore(style.SpacingBefore) + style.Thickness + style.SpacingAfter;
            }

            if (block is TextFieldBlock textField) {
                return ResolveTopLevelSpacingBefore(textField.SpacingBefore) + textField.Height + textField.SpacingAfter;
            }

            if (block is TextAnnotationBlock textAnnotation)
                return ResolveTopLevelSpacingBefore(textAnnotation.SpacingBefore) + textAnnotation.Height + textAnnotation.SpacingAfter;
            if (block is FreeTextAnnotationBlock freeTextAnnotation)
                return ResolveTopLevelSpacingBefore(freeTextAnnotation.SpacingBefore) + freeTextAnnotation.Height + freeTextAnnotation.SpacingAfter;
            if (block is HighlightAnnotationBlock highlightAnnotation)
                return ResolveTopLevelSpacingBefore(highlightAnnotation.SpacingBefore) + highlightAnnotation.Height + highlightAnnotation.SpacingAfter;

            if (block is CheckBoxBlock checkBox) {
                return ResolveTopLevelSpacingBefore(checkBox.SpacingBefore) + checkBox.Size + checkBox.SpacingAfter;
            }

            if (block is ChoiceFieldBlock choiceField) {
                return ResolveTopLevelSpacingBefore(choiceField.SpacingBefore) + choiceField.Height + choiceField.SpacingAfter;
            }

            if (block is RadioButtonGroupBlock radioButtonGroup) {
                return ResolveTopLevelSpacingBefore(radioButtonGroup.SpacingBefore) + radioButtonGroup.Height + radioButtonGroup.SpacingAfter;
            }

            if (block is ImageBlock image) {
                return MeasureImageBlockHeight(image, frameWidth);
            }

            if (block is ShapeBlock shape) {
                return MeasureShapeBlockHeight(shape);
            }

            if (block is DrawingBlock drawing) {
                return MeasureDrawingBlockHeight(drawing);
            }

            if (block is RowBlock row) {
                return MeasureRowBlockHeight(row, frameX, frameWidth, fontSize, firstVisualOnly: true);
            }

            return 0D;
        }

        private double MeasureFirstNestedVisualHeight(IReadOnlyList<IPdfBlock> blocks, double frameX, double frameWidth, double fontSize,
            bool allowTableFragments = false, bool suppressParagraphSpacingBefore = false) {
            for (int index = 0; index < blocks.Count; index++) {
                IPdfBlock block = blocks[index];
                if (block is BookmarkBlock || block is ColumnBreakBlock) {
                    continue;
                }

                return MeasureNextBlockFirstVisualHeight(block, frameX, frameWidth, fontSize, allowTableFragments, suppressParagraphSpacingBefore);
            }

            return 0D;
        }

        private double MeasureDeferredTableFirstVisualHeight(DeferredTableBlock table, double frameWidth, double fontSize) {
            PdfTableStyle style = table.Style ?? currentOpts.DefaultTableStyleSnapshot ?? TableStyles.Light();
            using System.Collections.Generic.IEnumerator<DeferredTableBatch> batches = table.CreateBatches(style).GetEnumerator();
            return batches.MoveNext()
                ? MeasureTableBlockHeight(batches.Current.Table, frameWidth, fontSize, firstVisualOnly: true)
                : 0D;
        }

        private double MeasureImageBlockHeight(ImageBlock image, double frameWidth) {
            PdfImageStyle style = ResolveImageStyle(image, currentOpts);
            double spacingBefore = ResolveTopLevelSpacingBefore(style.SpacingBefore);
            var box = ResolveImageFlowBox(image, style, frameWidth, spacingBefore, style.SpacingAfter);
            return spacingBefore + box.Height + style.SpacingAfter;
        }

        private double MeasureShapeBlockHeight(ShapeBlock shape) {
            PdfDrawingStyle style = ResolveDrawingStyle(shape, currentOpts);
            return ResolveTopLevelSpacingBefore(style.SpacingBefore) + shape.Shape.Height + style.SpacingAfter;
        }

        private double MeasureDrawingBlockHeight(DrawingBlock drawing) {
            PdfDrawingStyle style = ResolveDrawingStyle(drawing, currentOpts);
            return ResolveTopLevelSpacingBefore(style.SpacingBefore) + drawing.Drawing.Height + style.SpacingAfter;
        }


        private double MeasureRowBlockHeight(RowBlock row, double frameX, double frameWidth, double fontSize, bool firstVisualOnly) {
            int columns = row.Columns.Count;
            PdfRowStyle? rowStyle = row.StyleSnapshot ?? currentOpts.DefaultRowStyleSnapshot;
            double spacingBefore = ResolveTopLevelSpacingBefore(rowStyle?.SpacingBefore ?? 0D);
            if (columns == 0) {
                double spacingAfter = firstVisualOnly ? 0D : rowStyle?.SpacingAfter ?? 0D;
                return spacingBefore + spacingAfter;
            }

            double rowGap = row.GapOverride ?? rowStyle?.Gap ?? PdfRowStyle.DefaultGap;
            double totalGap = rowGap * Math.Max(0, columns - 1);
            if (totalGap >= frameWidth) {
                return spacingBefore;
            }

            double columnAreaWidth = frameWidth - totalGap;
            double[] columnWidths = ResolveRowColumnWidths(row, columnAreaWidth);
            double tallestFirstVisual = 0D;
            for (int columnIndex = 0; columnIndex < columns; columnIndex++) {
                RowColumn column = row.Columns[columnIndex];
                double columnWidth = columnWidths[columnIndex];
                if (firstVisualOnly && column.Blocks.Count > 0) {
                    tallestFirstVisual = Math.Max(tallestFirstVisual, MeasureNextBlockFirstVisualHeight(column.Blocks[0], frameX, columnWidth, fontSize));
                }
            }

            if (firstVisualOnly) {
                return spacingBefore + tallestFirstVisual;
            }

            var columnItems = BuildRowColumnItems(row, columnWidths);
            double contentHeight = 0D;
            foreach (var items in columnItems) {
                contentHeight = Math.Max(contentHeight, MeasureRowKeepTogetherHeight(items));
            }

            return spacingBefore + contentHeight + (rowStyle?.SpacingAfter ?? 0D);
        }

        private double MeasureTableBlockHeight(TableBlock table, double frameWidth, double fontSize, bool firstVisualOnly, bool allowTableFragments = false) {
            PdfTableStyle style = table.Style ?? currentOpts.DefaultTableStyleSnapshot ?? TableStyles.Light();
            int columns = GetTableColumnCount(table);
            if (columns == 0 || table.Rows.Count == 0) {
                return ResolveTopLevelSpacingBefore(style.SpacingBefore) + (firstVisualOnly ? 0D : style.SpacingAfter);
            }

            double columnGap = GetTableCellSpacing(style);
            double rowGap = columnGap;
            ValidateTableRoleRowCounts(style, table.Rows.Count);
            int headerRowCount = style.HeaderRowCount;
            int footerRowCount = style.FooterRowCount;
            int footerStartRowIndex = table.Rows.Count - footerRowCount;
            ValidateTableCellStyleCoordinates(style, table, columns);
            ValidateTableColumnStyleBounds(style, columns);
            ValidateTableRowStyleBounds(style, table.Rows.Count);
            ValidateTableRowSpansWithinRoleBoundaries(table, columns, headerRowCount, footerStartRowIndex);
            double tableFontSize = GetTableBodyFontSize(style, fontSize);
            TableColumnLayout columnLayout = ResolveTableColumnLayout(table, currentOpts, style, columns, frameWidth, tableFontSize, headerRowCount, footerStartRowIndex);

            PreparedFlowTableRows prepared = PrepareFlowTableRows(table, style, columns, columnLayout.Widths,
                columnGap, rowGap, headerRowCount, footerStartRowIndex, fallbackFontSize: fontSize);
            double[] rowHeights = prepared.Heights;
            double captionHeight = 0D;
            if (!string.IsNullOrWhiteSpace(style.Caption)) {
                double captionSize = style.CaptionFontSize ?? fontSize;
                double captionLeading = captionSize * 1.25D;
                var captionRuns = new[] { PdfTextRun.Normal(style.Caption!, style.CaptionColor, captionSize) };
                var captionWrap = WrapRichRunsCore(captionRuns, columnLayout.Width, captionSize, ChooseNormal(currentOpts.DefaultFont), captionLeading, null, DefaultParagraphTabStopWidth, currentOpts);
                captionHeight = MeasureRichLinesHeight(captionWrap.LineHeights, captionWrap.Lines.Count, captionLeading) + style.CaptionSpacingAfter;
            }

            int measuredRowCount = firstVisualOnly ? 1 : rowHeights.Length;
            double rowHeight = firstVisualOnly && allowTableFragments
                ? MeasureTableFirstFragmentHeight(table, style, prepared, columns, columnLayout.Widths, columnGap)
                : GetTableRowsHeight(rowHeights, 0, measuredRowCount, rowGap);
            double tableHeight = (style.Position == null ? ResolveTopLevelSpacingBefore(style.SpacingBefore) : 0D) + captionHeight + rowHeight;
            return firstVisualOnly || style.Position != null ? tableHeight : tableHeight + style.SpacingAfter;
        }

        private void ConsumeSpacer(double height) {
            double remaining = height;
            while (remaining > 0.001D) {
                double available = y - currentOpts.MarginBottom;
                if (available <= 0.5D) {
                    NewPage();
                    continue;
                }

                // Empty pages are discarded by FlushPage. Jump over complete
                // blank spacer pages instead of allocating and discarding each
                // one; count them against the generated-page limit regardless.
                if (activeColumnFlow == null && remaining > available &&
                    System.Math.Abs(y - yStart) <= 0.001D && !pageDirty &&
                    !HasCurrentPageNonContentObjects() && sb.Length == 0 &&
                    activeLayers.Count == 0 && activeContainerScopes.Count == 0 &&
                    activeFloatingFlowCaptures.Count == 0 && pendingFloatingBookmarks.Count == 0 &&
                    behindTextCanvases.Count == 0 && !HasFloatingTables) {
                    double completePages = System.Math.Floor((remaining - 0.001D) / available);
                    if (completePages >= 2D) {
                        if (completePages > long.MaxValue - startedPageCount)
                            throw new InvalidDataException("PDF spacer traversed too many generated pages.");
                        long skippedPages = (long)completePages;
                        startedPageCount += skippedPages - 1;
                        remaining = System.Math.Max(0D, remaining - skippedPages * available);
                        NewPage();
                        continue;
                    }
                }

                RecordFlowPlacement(y);
                double consumed = Math.Min(remaining, available);
                y -= consumed;
                remaining -= consumed;
                if (remaining > 0.001D) {
                    NewPage();
                }
            }
        }

        private void RenderHorizontalRuleBlock(HorizontalRuleBlock block, double containerX, double containerWidth) {
            double frameMarginLeft = currentOpts.MarginLeft;
            PdfHorizontalRuleStyle ruleStyle = ResolveHorizontalRuleStyle(block, currentOpts);
            ValidateHorizontalRule(ruleStyle);
            double spacingBefore = PlaceFixedFlowBlock("Horizontal rule", 0D, ruleStyle.Thickness,
                ruleStyle.SpacingBefore, ruleStyle.SpacingAfter, ref containerWidth);
            if (spacingBefore > 0) y -= spacingBefore;
            containerX += currentOpts.MarginLeft - frameMarginLeft;
            RecordFlowPlacement(y);
            double yLine = y - ruleStyle.Thickness * 0.5;
            DrawHLine(sb, ruleStyle.Color, ruleStyle.Thickness, containerX, containerX + containerWidth, yLine, emitGeneratedStructure);
            pageDirty = true;
            y -= ruleStyle.Thickness + ruleStyle.SpacingAfter;
        }

    }
}
