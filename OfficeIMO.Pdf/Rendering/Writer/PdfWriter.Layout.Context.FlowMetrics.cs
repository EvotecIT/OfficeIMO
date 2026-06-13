using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private void EnsureFixedFlowBlockFits(string blockName, double blockWidth, double blockHeight, double availableWidth) {
            if (blockWidth > availableWidth + 0.001) {
                throw new ArgumentException(blockName + " width exceeds the available page content width.");
            }

            double availableHeight = currentOpts.PageHeight - currentOpts.MarginTop - currentOpts.MarginBottom;
            if (blockHeight > availableHeight + 0.001) {
                throw new ArgumentException(blockName + " height exceeds the available page content height.");
            }
        }

        private (double Width, double Height) ResolveImageFlowBox(ImageBlock image, PdfImageStyle style, double frameWidth, double spacingBefore, double spacingAfter) {
            double imageWidth = image.Width;
            double imageHeight = image.Height;
            if (!style.ScaleDownToFit) {
                return (imageWidth, imageHeight);
            }

            double availableHeight = currentOpts.PageHeight - currentOpts.MarginTop - currentOpts.MarginBottom - spacingBefore - spacingAfter;
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

        private static void ValidatePanelStyle(PanelStyle style, double panelWidth) {
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

        private void RenderListItem(System.Collections.Generic.IReadOnlyList<TextRun> runs, System.Collections.Generic.List<System.Collections.Generic.List<RichSeg>> lines, System.Collections.Generic.List<double> lineHeights, string marker, double markerX, double markerWidth, PdfAlign markerAlign, double textX, double textWidth, PdfAlign textAlign, PdfColor? color, double size, double leading, double spacingBefore, double spacingAfter, string? bookmarkName, ref int? listStructureElementIndex, ref LayoutResult.Page? listStructurePage) {
            int lineIndex = 0;
            bool firstSegment = true;
            var listFont = ChooseNormal(currentOpts.DefaultFont);
            PageStructElement? listItemElement = null;
            spacingBefore = ResolveTopLevelSpacingBefore(spacingBefore);
            if (spacingBefore > 0) {
                if (y - spacingBefore < currentOpts.MarginBottom) {
                    NewPage();
                    spacingBefore = 0D;
                }

                if (spacingBefore > 0) y -= spacingBefore;
            }

            while (lineIndex < lines.Count) {
                double available = y - currentOpts.MarginBottom;
                double firstLineHeight = GetRichLineHeight(lineHeights, lineIndex, leading);
                if (available < firstLineHeight) {
                    NewPage();
                    available = y - currentOpts.MarginBottom;
                    if (available < firstLineHeight) {
                        break;
                    }
                }

                int take = 0;
                double heightSum = 0;
                for (int k = lineIndex; k < lines.Count; k++) {
                    double lineHeight = GetRichLineHeight(lineHeights, k, leading);
                    if (heightSum + lineHeight > available) {
                        break;
                    }

                    heightSum += lineHeight;
                    take++;
                }

                if (take == 0) {
                    NewPage();
                    continue;
                }

                var segmentLines = new System.Collections.Generic.List<System.Collections.Generic.List<RichSeg>>(take);
                var segmentHeights = new System.Collections.Generic.List<double>(take);
                for (int k = 0; k < take; k++) {
                    segmentLines.Add(lines[lineIndex + k]);
                    segmentHeights.Add(GetRichLineHeight(lineHeights, lineIndex + k, leading));
                }

                double baselineY = FirstTextBaselineFromTop(listFont, size, y);
                int? listElementIndex = firstSegment ? EnsurePageStructureContainer("L", ref listStructureElementIndex, ref listStructurePage) : null;
                int? listItemElementIndex = firstSegment ? RegisterStructureContainer("LI", listElementIndex) : null;
                if (firstSegment) {
                    if (listItemElementIndex.HasValue && currentPage != null) {
                        listItemElement = currentPage.StructElements[listItemElementIndex.Value];
                    }

                    if (!string.IsNullOrEmpty(bookmarkName)) {
                        AddNamedDestinationName(bookmarkName!, y);
                    }

                    var markerLines = new System.Collections.Generic.List<string>(1) { marker };
                    int? labelMarkedContentId = RegisterTextStructureElement("Lbl", listItemElementIndex);
                    WriteLinesInternal("F1", size, leading, markerX, markerWidth, baselineY, markerLines, markerAlign, color, applyBaselineTweak: true, structureType: "Lbl", markedContentId: labelMarkedContentId);
                }

                pageDirty = true;
                int? bodyMarkedContentId = firstSegment || listItemElement == null
                    ? RegisterTextStructureElement("LBody", listItemElementIndex)
                    : RegisterTextStructureElement("LBody", listItemElement);
                WriteRichParagraph(sb, new RichParagraphBlock(runs, textAlign, color), segmentLines, segmentHeights, currentOpts, baselineY, size, leading, currentPage!.Annotations, textX, textWidth, structureType: "LBody", markedContentId: bodyMarkedContentId, structurePage: currentPage);
                MarkRichFonts(runs);
                y -= heightSum;
                lineIndex += take;
                firstSegment = false;
                if (lineIndex < lines.Count) {
                    NewPage();
                } else {
                    y -= spacingAfter;
                }
            }
        }

        private static double MeasureListKeepTogetherHeight(System.Collections.Generic.IReadOnlyList<TableCellTextLayout> itemLayouts, double leading, double spacingBefore, double itemSpacing, double spacingAfter) {
            double total = 0D;
            for (int itemIndex = 0; itemIndex < itemLayouts.Count; itemIndex++) {
                total += itemIndex == 0 ? spacingBefore : 0D;
                total += MeasureRichLinesHeight(itemLayouts[itemIndex].LineHeights, itemLayouts[itemIndex].LineCount, leading);
                total += itemIndex == itemLayouts.Count - 1 ? spacingAfter : itemSpacing;
            }

            return total;
        }

        private PdfParagraphStyle? EffectiveParagraphStyle(RichParagraphBlock paragraph) => paragraph.Style ?? currentOpts.DefaultParagraphStyleSnapshot;

        private double MeasureNextParagraphFirstLineHeight(RichParagraphBlock paragraph, double frameX, double frameWidth, double fontSize) {
            PdfParagraphStyle? paragraphStyle = EffectiveParagraphStyle(paragraph);
            double leading = GetParagraphLeading(paragraphStyle, fontSize);
            double spacingBefore = GetParagraphSpacingBefore(paragraphStyle);
            var textFrame = GetParagraphTextFrame(paragraphStyle, frameX, frameWidth);
            var wrap = WrapRichRunsCore(paragraph.Runs, textFrame.Width, fontSize, ChooseNormal(currentOpts.DefaultFont), leading, textFrame.FirstLineWidth, GetParagraphTabStopWidth(paragraphStyle), currentOpts, paragraphStyle?.TabStops.ToArray());
            return wrap.LineHeights.Count == 0 ? spacingBefore : spacingBefore + wrap.LineHeights[0];
        }

        private double MeasureNextBlockFirstVisualHeight(IPdfBlock block, double frameX, double frameWidth, double fontSize) {
            if (block is RichParagraphBlock paragraph) {
                return MeasureNextParagraphFirstLineHeight(paragraph, frameX, frameWidth, fontSize);
            }

            if (block is HeadingBlock heading) {
                PdfHeadingStyle? headingStyle = ResolveHeadingStyle(heading, currentOpts);
                double headingSize = GetHeadingFontSize(heading, headingStyle);
                double headingLeading = GetHeadingLeading(headingStyle, headingSize);
                return (headingStyle?.SpacingBefore ?? 0D) + headingLeading;
            }

            if (block is SpacerBlock spacer) {
                return spacer.Height;
            }

            if (block is BulletListBlock bullets) {
                PdfListStyle? listStyle = ResolveListStyle(bullets, currentOpts);
                double size = GetListFontSize(listStyle, fontSize);
                double leading = GetListLeading(listStyle, size);
                string? firstItem = bullets.Items.Count > 0 ? bullets.Items[0] : null;
                if (firstItem == null) {
                    return listStyle?.SpacingBefore ?? 0D;
                }

                return (listStyle?.SpacingBefore ?? 0D) + leading;
            }

            if (block is NumberedListBlock numbered) {
                PdfListStyle? listStyle = ResolveListStyle(numbered, currentOpts);
                double size = GetListFontSize(listStyle, fontSize);
                double leading = GetListLeading(listStyle, size);
                string? firstItem = numbered.Items.Count > 0 ? numbered.Items[0] : null;
                if (firstItem == null) {
                    return listStyle?.SpacingBefore ?? 0D;
                }

                return (listStyle?.SpacingBefore ?? 0D) + leading;
            }

            if (block is PanelParagraphBlock panel) {
                PanelStyle panelStyle = ResolvePanelStyle(panel, currentOpts);
                double innerWidth = panelStyle.MaxWidth.HasValue ? Math.Min(frameWidth, panelStyle.MaxWidth.Value) : frameWidth;
                ValidatePanelStyle(panelStyle, innerWidth);
                double size = fontSize;
                double leading = size * 1.4;
                double textWidth = innerWidth - 2 * panelStyle.PaddingX;
                var wrap = WrapRichRunsCore(panel.Runs, textWidth, size, ChooseNormal(currentOpts.DefaultFont), leading, null, DefaultParagraphTabStopWidth, currentOpts);
                double firstLineHeight = wrap.LineHeights.Count == 0 ? 0D : wrap.LineHeights[0];
                return panelStyle.SpacingBefore + panelStyle.PaddingY + firstLineHeight + panelStyle.PaddingY;
            }

            if (block is TableBlock table) {
                PdfTableStyle style = table.Style ?? currentOpts.DefaultTableStyleSnapshot ?? TableStyles.Light();
                int columns = GetTableColumnCount(table);
                if (columns == 0) {
                    return style.SpacingBefore;
                }

                double padLeft = GetTableCellPaddingLeft(style);
                double padRight = GetTableCellPaddingRight(style);
                double padTop = GetTableCellPaddingTop(style);
                double padBottom = GetTableCellPaddingBottom(style);
                double columnGap = GetTableCellSpacing(style);
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
                double tableWidth = columnLayout.Width;
                double rowSize = GetTableRowFontSize(style, 0, headerRowCount, footerStartRowIndex, fontSize);
                double rowLeading = GetTableLeading(style, rowSize);
                bool rowUsesBold = GetTableRowBold(style, 0, headerRowCount, footerStartRowIndex);
                int maxLines = 1;
                var firstRowCells = GetTableCellLayouts(table, 0, columns);
                for (int cellIndex = 0; cellIndex < firstRowCells.Count; cellIndex++) {
                    TableCellLayout cell = firstRowCells[cellIndex];
                    double cellWidth = GetTableCellWidth(columnLayout.Widths, cell.Column, cell.ColumnSpan, columnGap);
                    double innerWidth = Math.Max(1D, cellWidth - GetTableCellPaddingLeft(style, 0, cell.Column) - GetTableCellPaddingRight(style, 0, cell.Column));
                    var lines = WrapSimpleTextForOptions(cell.Text, innerWidth, GetTableRowFont(currentOpts, rowUsesBold), rowSize, currentOpts);
                    maxLines = Math.Max(maxLines, lines.Count);
                }

                    double firstRowHeight = Math.Max(maxLines * rowLeading + GetTableRowMaxPaddingTop(table, style, 0, columns) + GetTableRowMaxPaddingBottom(table, style, 0, columns), GetTableRowMinHeight(style, 0));
                double captionHeight = 0D;
                if (!string.IsNullOrWhiteSpace(style.Caption)) {
                    double captionSize = style.CaptionFontSize ?? fontSize;
                    double captionLeading = captionSize * 1.25D;
                    var captionRuns = new[] { TextRun.Normal(style.Caption!, style.CaptionColor, captionSize) };
                    var captionWrap = WrapRichRunsCore(captionRuns, tableWidth, captionSize, ChooseNormal(currentOpts.DefaultFont), captionLeading, null, DefaultParagraphTabStopWidth, currentOpts);
                    captionHeight = MeasureRichLinesHeight(captionWrap.LineHeights, captionWrap.Lines.Count, captionLeading) + style.CaptionSpacingAfter;
                }

                return style.SpacingBefore + captionHeight + firstRowHeight;
            }

            if (block is HorizontalRuleBlock rule) {
                PdfHorizontalRuleStyle style = ResolveHorizontalRuleStyle(rule, currentOpts);
                return style.SpacingBefore + style.Thickness + style.SpacingAfter;
            }

            if (block is TextFieldBlock textField) {
                return textField.SpacingBefore + textField.Height + textField.SpacingAfter;
            }

            if (block is CheckBoxBlock checkBox) {
                return checkBox.SpacingBefore + checkBox.Size + checkBox.SpacingAfter;
            }

            if (block is ChoiceFieldBlock choiceField) {
                return choiceField.SpacingBefore + choiceField.Height + choiceField.SpacingAfter;
            }

            if (block is RadioButtonGroupBlock radioButtonGroup) {
                return radioButtonGroup.SpacingBefore + radioButtonGroup.Height + radioButtonGroup.SpacingAfter;
            }

            if (block is ImageBlock image) {
                PdfImageStyle style = ResolveImageStyle(image, currentOpts);
                var box = ResolveImageFlowBox(image, style, frameWidth, style.SpacingBefore, style.SpacingAfter);
                return style.SpacingBefore + box.Height + style.SpacingAfter;
            }

            if (block is ShapeBlock shape) {
                PdfDrawingStyle style = ResolveDrawingStyle(shape, currentOpts);
                return style.SpacingBefore + shape.Shape.Height + style.SpacingAfter;
            }

            if (block is DrawingBlock drawing) {
                PdfDrawingStyle style = ResolveDrawingStyle(drawing, currentOpts);
                return style.SpacingBefore + drawing.Drawing.Height + style.SpacingAfter;
            }

            if (block is RowBlock row) {
                int columns = row.Columns.Count;
                if (columns == 0) {
                    return 0D;
                }

                PdfRowStyle? rowStyle = row.StyleSnapshot ?? currentOpts.DefaultRowStyleSnapshot;
                    double rowGap = row.GapOverride ?? rowStyle?.Gap ?? PdfRowStyle.DefaultGap;
                double totalGap = rowGap * Math.Max(0, columns - 1);
                if (totalGap >= frameWidth) {
                    return rowStyle?.SpacingBefore ?? 0D;
                }

                double columnAreaWidth = frameWidth - totalGap;
                double tallestFirstVisual = 0D;
                for (int columnIndex = 0; columnIndex < columns; columnIndex++) {
                    RowColumn column = row.Columns[columnIndex];
                    if (column.Blocks.Count == 0) {
                        continue;
                    }

                    double columnWidth = Math.Max(0D, columnAreaWidth * (column.WidthPercent / 100D));
                    tallestFirstVisual = Math.Max(tallestFirstVisual, MeasureNextBlockFirstVisualHeight(column.Blocks[0], frameX, columnWidth, fontSize));
                }

                return (rowStyle?.SpacingBefore ?? 0D) + tallestFirstVisual;
            }

            return 0D;
        }

        private void ConsumeSpacer(double height) {
            double remaining = height;
            while (remaining > 0.001D) {
                double available = y - currentOpts.MarginBottom;
                if (available <= 0.5D) {
                    NewPage();
                    continue;
                }

                double consumed = Math.Min(remaining, available);
                y -= consumed;
                remaining -= consumed;
                if (remaining > 0.001D) {
                    NewPage();
                }
            }
        }

        private void RenderHorizontalRuleBlock(HorizontalRuleBlock block, double containerX, double containerWidth) {
            PdfHorizontalRuleStyle ruleStyle = ResolveHorizontalRuleStyle(block, currentOpts);
            ValidateHorizontalRule(ruleStyle);
            double spacingBefore = ResolveTopLevelSpacingBefore(ruleStyle.SpacingBefore);
            double needed = spacingBefore + ruleStyle.Thickness + ruleStyle.SpacingAfter;
            EnsureFixedFlowBlockFits("Horizontal rule", containerWidth, needed, containerWidth);
            if (y - needed < currentOpts.MarginBottom) {
                NewPage();
                spacingBefore = 0D;
            }
            if (spacingBefore > 0) y -= spacingBefore;
            double yLine = y - ruleStyle.Thickness * 0.5;
            DrawHLine(sb, ruleStyle.Color, ruleStyle.Thickness, containerX, containerX + containerWidth, yLine, emitGeneratedStructure);
            pageDirty = true;
            y -= ruleStyle.Thickness + ruleStyle.SpacingAfter;
        }

    }
}

