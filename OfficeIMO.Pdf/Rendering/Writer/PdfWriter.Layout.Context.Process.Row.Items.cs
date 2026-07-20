using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private System.Collections.Generic.List<System.Collections.Generic.List<ColItem>> BuildRowColumnItems(RowBlock rb, double[] colWs) {
            int ncols = rb.Columns.Count;
        var colItems = new System.Collections.Generic.List<System.Collections.Generic.List<ColItem>>(ncols);
        for (int i = 0; i < ncols; i++) {
            var items = new System.Collections.Generic.List<ColItem>();
            int nextListGroupId = 1;
            foreach (var cb in rb.Columns[i].Blocks) {
                if (cb is HeadingBlock hb2) {
                    PdfHeadingStyle? headingStyle = ResolveHeadingStyle(hb2, currentOpts);
                    double size = GetHeadingFontSize(hb2, headingStyle);
                    double leading = GetHeadingLeading(headingStyle, size);
                    PdfColor? headingColor = hb2.Color ?? headingStyle?.Color;
                    System.Collections.Generic.IReadOnlyList<TextRun> headingRuns = CreateHeadingTextRuns(hb2, headingStyle, headingColor);
                    var wrap = WrapRichRunsCore(headingRuns, colWs[i], size, ChooseNormal(currentOpts.DefaultFont), leading, null, DefaultParagraphTabStopWidth, currentOpts);
                    items.Add(new ColHead {
                        Block = hb2,
                        Runs = headingRuns,
                        Lines = wrap.Lines,
                        Heights = wrap.LineHeights,
                        Leading = leading,
                        Size = size,
                        SpacingBefore = headingStyle?.SpacingBefore ?? 0D,
                        SpacingAfter = GetHeadingSpacingAfter(headingStyle, leading),
                        Bold = GetHeadingBold(headingStyle),
                        ApplySpacingBeforeAtTop = headingStyle?.ApplySpacingBeforeAtTop ?? false,
                        KeepWithNext = headingStyle?.KeepWithNext ?? true,
                        Color = headingColor
                    });
                } else if (cb is RichParagraphBlock rpb2) {
                    double size = currentOpts.DefaultFontSize;
                    PdfParagraphStyle? paragraphStyle = EffectiveParagraphStyle(rpb2);
                    double leading = GetParagraphLeading(paragraphStyle, size);
                    var textFrame = GetParagraphTextFrame(paragraphStyle, 0, colWs[i]);
                    var wrap = WrapRichRunsCoreWithFirstLineOrigin(rpb2.Runs, textFrame.Width, size, ChooseNormal(currentOpts.DefaultFont), leading, textFrame.FirstLineWidth, textFrame.FirstLineX - textFrame.X, GetParagraphTabStopWidth(paragraphStyle), currentOpts, paragraphStyle?.TabStops.ToArray());
                    items.Add(new ColPar { Block = rpb2, Lines = wrap.Lines, Heights = wrap.LineHeights, Leading = leading, Size = size, XOffset = textFrame.X, TextWidth = textFrame.Width, FirstLineXOffset = textFrame.FirstLineX, FirstLineTextWidth = textFrame.FirstLineWidth });
                } else if (cb is BulletListBlock bl2) {
                    PdfListStyle? listStyle = ResolveListStyle(bl2, currentOpts);
                    double size = GetListFontSize(listStyle, currentOpts.DefaultFontSize);
                    double markerSize = GetListMarkerFontSize(listStyle, size);
                    double leading = Math.Max(GetListLeading(listStyle, size), GetListLeading(listStyle, markerSize));
                    var baseFont = ChooseNormal(currentOpts.DefaultFont);
                    PdfStandardFont markerFont = GetListMarkerFont(listStyle, currentOpts.DefaultFont);
                    PdfNamedFontFace? markerNamedFont = GetListMarkerNamedFont(listStyle, currentOpts);
                    const string bulletGlyph = "•";
                    double estimatedBulletWidth = bl2.RichItems.Count == 0
                        ? EstimateSimpleTextWidthForOptions(bulletGlyph, markerFont, markerNamedFont, markerSize, currentOpts)
                        : bl2.RichItems.Max(item => EstimateSimpleTextWidthForOptions(item.Marker ?? bulletGlyph, markerFont, markerNamedFont, markerSize, currentOpts));
                    double bulletWidth = GetListMarkerWidth(listStyle, estimatedBulletWidth);
                    double spaceAdvance = EstimateSimpleTextWidthForOptions(" ", markerFont, markerNamedFont, markerSize, currentOpts);
                    double markerGap = GetListMarkerGap(listStyle, spaceAdvance);
                    double indent = bulletWidth + markerGap;
                    double listLeftIndent = listStyle?.LeftIndent ?? 0D;
                    double rawTextWidth = colWs[i] - listLeftIndent - indent;
                    double availableWidth = Math.Max(rawTextWidth, EstimateSimpleTextWidthForOptions("WW", baseFont, size, currentOpts));
                    double alignmentWidth = Math.Max(0, rawTextWidth);
                    double itemSpacing = GetListItemSpacing(listStyle, leading);
                    int listGroupId = nextListGroupId++;
                    var listItems = new System.Collections.Generic.List<ColListItem>(bl2.RichItems.Count);
                    for (int itemIndex = 0; itemIndex < bl2.RichItems.Count; itemIndex++) {
                        var item = bl2.RichItems[itemIndex];
                        string marker = item.Marker ?? bulletGlyph;
                        var layout = CreateListItemTextLayout(item, availableWidth, baseFont, size, leading, currentOpts);
                        double firstLineWidth = layout.Lines.Count > 0 ? MeasureRichLineWidth(layout.Lines[0], currentOpts) : 0;
                        double firstLineDx = 0;
                        if (bl2.Align == PdfAlign.Center) firstLineDx = Math.Max(0, (alignmentWidth - firstLineWidth) / 2);
                        else if (bl2.Align == PdfAlign.Right) firstLineDx = Math.Max(0, alignmentWidth - firstLineWidth);
                        double spacingBefore = itemIndex == 0 ? listStyle?.SpacingBefore ?? 0D : 0D;
                        double spacingAfter = itemIndex == bl2.RichItems.Count - 1 ? listStyle?.GetSpacingAfter(itemSpacing) ?? itemSpacing : itemSpacing;
                        PdfColor? listColor = bl2.Color ?? listStyle?.Color;
                        listItems.Add(new ColListItem { Runs = item.Runs, Lines = layout.Lines, Heights = layout.LineHeights, Marker = marker, MarkerFont = markerFont, MarkerNamedFont = markerNamedFont, MarkerSize = markerSize, MarkerColor = listStyle?.MarkerColor ?? listColor, MarkerXOffset = listLeftIndent + firstLineDx, MarkerWidth = bulletWidth, MarkerAlign = GetBulletMarkerAlign(listStyle), TextXOffset = listLeftIndent + indent, TextWidth = alignmentWidth, TextAlign = bl2.Align, Color = listColor, Leading = leading, Size = size, SpacingBefore = spacingBefore, SpacingAfter = spacingAfter, BookmarkName = item.BookmarkName, ListGroupId = listGroupId });
                    }

                    if ((listStyle?.KeepTogether == true || listStyle?.KeepWithNext == true) && listItems.Count > 0) {
                        double listGroupHeight = 0D;
                        foreach (var listItem in listItems) {
                            listGroupHeight += listItem.SpacingBefore + MeasureRichLinesHeight(listItem.Heights, listItem.Lines.Count, listItem.Leading) + listItem.SpacingAfter;
                        }

                        if (listStyle?.KeepTogether == true) {
                            listItems[0].IsFirstInKeepGroup = true;
                            foreach (var listItem in listItems) {
                                listItem.KeepTogether = true;
                                listItem.KeepGroupHeight = listGroupHeight;
                            }
                        }

                        if (listStyle?.KeepWithNext == true) {
                            listItems[0].IsFirstInKeepWithNextGroup = true;
                            foreach (var listItem in listItems) {
                                listItem.KeepWithNext = true;
                                listItem.KeepWithNextGroupItemCount = listItems.Count;
                                listItem.KeepWithNextGroupHeight = listGroupHeight;
                            }
                        }
                    }

                    items.AddRange(listItems);
                } else if (cb is NumberedListBlock nl2) {
                    PdfListStyle? listStyle = ResolveListStyle(nl2, currentOpts);
                    double size = GetListFontSize(listStyle, currentOpts.DefaultFontSize);
                    double markerSize = GetListMarkerFontSize(listStyle, size);
                    double leading = Math.Max(GetListLeading(listStyle, size), GetListLeading(listStyle, markerSize));
                    var baseFont = ChooseNormal(currentOpts.DefaultFont);
                    PdfStandardFont markerFont = GetListMarkerFont(listStyle, currentOpts.DefaultFont);
                    PdfNamedFontFace? markerNamedFont = GetListMarkerNamedFont(listStyle, currentOpts);
                    int lastNumber = nl2.StartNumber + Math.Max(0, nl2.RichItems.Count - 1);
                    string widestMarker = lastNumber.ToString(CultureInfo.InvariantCulture) + ".";
                    double estimatedMarkerWidth = nl2.RichItems.Count == 0
                        ? EstimateSimpleTextWidthForOptions(widestMarker, markerFont, markerNamedFont, markerSize, currentOpts)
                        : nl2.RichItems
                            .Select((item, itemIndex) => item.Marker ?? ((nl2.StartNumber + itemIndex).ToString(CultureInfo.InvariantCulture) + "."))
                            .Max(marker => EstimateSimpleTextWidthForOptions(marker, markerFont, markerNamedFont, markerSize, currentOpts));
                    double markerWidth = GetListMarkerWidth(listStyle, estimatedMarkerWidth);
                    double spaceAdvance = EstimateSimpleTextWidthForOptions(" ", markerFont, markerNamedFont, markerSize, currentOpts);
                    double markerGap = GetListMarkerGap(listStyle, spaceAdvance);
                    double indent = markerWidth + markerGap;
                    double listLeftIndent = listStyle?.LeftIndent ?? 0D;
                    double rawTextWidth = colWs[i] - listLeftIndent - indent;
                    double availableWidth = Math.Max(rawTextWidth, EstimateSimpleTextWidthForOptions("WW", baseFont, size, currentOpts));
                    double alignmentWidth = Math.Max(0, rawTextWidth);
                    double itemSpacing = GetListItemSpacing(listStyle, leading);
                    int listGroupId = nextListGroupId++;
                    var listItems = new System.Collections.Generic.List<ColListItem>(nl2.RichItems.Count);
                    for (int itemIndex = 0; itemIndex < nl2.RichItems.Count; itemIndex++) {
                        var item = nl2.RichItems[itemIndex];
                        string marker = item.Marker ?? ((nl2.StartNumber + itemIndex).ToString(CultureInfo.InvariantCulture) + ".");
                        var layout = CreateListItemTextLayout(item, availableWidth, baseFont, size, leading, currentOpts);
                        double firstLineWidth = layout.Lines.Count > 0 ? MeasureRichLineWidth(layout.Lines[0], currentOpts) : 0;
                        double firstLineDx = 0;
                        if (nl2.Align == PdfAlign.Center) firstLineDx = Math.Max(0, (alignmentWidth - firstLineWidth) / 2);
                        else if (nl2.Align == PdfAlign.Right) firstLineDx = Math.Max(0, alignmentWidth - firstLineWidth);
                        double spacingBefore = itemIndex == 0 ? listStyle?.SpacingBefore ?? 0D : 0D;
                        double spacingAfter = itemIndex == nl2.RichItems.Count - 1 ? listStyle?.GetSpacingAfter(itemSpacing) ?? itemSpacing : itemSpacing;
                        PdfColor? listColor = nl2.Color ?? listStyle?.Color;
                        listItems.Add(new ColListItem { Runs = item.Runs, Lines = layout.Lines, Heights = layout.LineHeights, Marker = marker, MarkerFont = markerFont, MarkerNamedFont = markerNamedFont, MarkerSize = markerSize, MarkerColor = listStyle?.MarkerColor ?? listColor, MarkerXOffset = listLeftIndent + firstLineDx, MarkerWidth = markerWidth, MarkerAlign = GetNumberedMarkerAlign(listStyle), TextXOffset = listLeftIndent + indent, TextWidth = alignmentWidth, TextAlign = nl2.Align, Color = listColor, Leading = leading, Size = size, SpacingBefore = spacingBefore, SpacingAfter = spacingAfter, BookmarkName = item.BookmarkName, ListGroupId = listGroupId });
                    }

                    if ((listStyle?.KeepTogether == true || listStyle?.KeepWithNext == true) && listItems.Count > 0) {
                        double listGroupHeight = 0D;
                        foreach (var listItem in listItems) {
                            listGroupHeight += listItem.SpacingBefore + MeasureRichLinesHeight(listItem.Heights, listItem.Lines.Count, listItem.Leading) + listItem.SpacingAfter;
                        }

                        if (listStyle?.KeepTogether == true) {
                            listItems[0].IsFirstInKeepGroup = true;
                            foreach (var listItem in listItems) {
                                listItem.KeepTogether = true;
                                listItem.KeepGroupHeight = listGroupHeight;
                            }
                        }

                        if (listStyle?.KeepWithNext == true) {
                            listItems[0].IsFirstInKeepWithNextGroup = true;
                            foreach (var listItem in listItems) {
                                listItem.KeepWithNext = true;
                                listItem.KeepWithNextGroupItemCount = listItems.Count;
                                listItem.KeepWithNextGroupHeight = listGroupHeight;
                            }
                        }
                    }

                    items.AddRange(listItems);
                } else if (cb is PanelParagraphBlock ppb2) {
                    double size = currentOpts.DefaultFontSize;
                    double leading = size * 1.4;
                    var panelFont = ChooseNormal(currentOpts.DefaultFont);
                    double firstBaselineOffset = GetAscenderForOptions(panelFont, size, currentOpts);
                    PanelStyle panelStyle = ResolvePanelStyle(ppb2, currentOpts);
                    double innerWidth = panelStyle.MaxWidth.HasValue ? Math.Min(colWs[i], panelStyle.MaxWidth.Value) : colWs[i];
                    ValidatePanelStyle(panelStyle, innerWidth);
                    double textWidthAvail = innerWidth - 2 * panelStyle.PaddingX;
                    var wrap = WrapRichRunsCore(ppb2.Runs, textWidthAvail, size, panelFont, leading, null, DefaultParagraphTabStopWidth, currentOpts);
                    double xOffset = 0;
                    if (panelStyle.Align == PdfAlign.Center) xOffset = Math.Max(0, (colWs[i] - innerWidth) / 2);
                    else if (panelStyle.Align == PdfAlign.Right) xOffset = Math.Max(0, colWs[i] - innerWidth);
                    items.Add(new ColPanel { Block = ppb2, Style = panelStyle, Lines = wrap.Lines, Heights = wrap.LineHeights, Leading = leading, Size = size, FirstBaselineOffset = firstBaselineOffset, XOffset = xOffset, PanelWidth = innerWidth, TextWidth = textWidthAvail });
                } else if (cb is TableBlock tb2) {
                    PdfTableStyle style = tb2.Style ?? currentOpts.DefaultTableStyleSnapshot ?? TableStyles.Light();
                    int cols = GetTableColumnCount(tb2);
                    if (cols == 0) {
                        continue;
                    }

                    double padLeft = GetTableCellPaddingLeft(style);
                    double padRight = GetTableCellPaddingRight(style);
                    double padTop = GetTableCellPaddingTop(style);
                    double padBottom = GetTableCellPaddingBottom(style);
                    double cellSpacing = GetTableCellSpacing(style);
                    double columnGap = cellSpacing;
                    double tableRowGap = cellSpacing;
                    double size = currentOpts.DefaultFontSize;
                    ValidateTableRoleRowCounts(style, tb2.Rows.Count);
                    int headerRowCount = style.HeaderRowCount;
                    int repeatHeaderRowCount = GetTableRepeatHeaderRowCount(style);
                    int footerRowCount = style.FooterRowCount;
                    int footerStartRowIndex = tb2.Rows.Count - footerRowCount;
                    ValidateTableCellStyleCoordinates(style, tb2, cols);
                    ValidateTableColumnStyleBounds(style, cols);
                    ValidateTableRowStyleBounds(style, tb2.Rows.Count);
                    ValidateTableRowSpansWithinRoleBoundaries(tb2, cols, headerRowCount, footerStartRowIndex);
                    PreparedTableColumns preparedColumns = PrepareTableColumns(tb2, style, colWs[i], size, headerRowCount, footerStartRowIndex);
                    double[] colPixel = preparedColumns.ColumnWidths;
                    double tableWidth = preparedColumns.TableWidth;
                    ValidateTableCellTextWidths(tb2, style, cols, colPixel, columnGap);

                    var rowLines = new TableCellTextLayout[tb2.Rows.Count][];
                    var rowLineCounts = new int[tb2.Rows.Count];
                    var rowHeights = new double[tb2.Rows.Count];
                    var rowLeadings = new double[tb2.Rows.Count];
                    var rowSizes = new double[tb2.Rows.Count];
                    var rowBold = new bool[tb2.Rows.Count];
                    for (int ri = 0; ri < tb2.Rows.Count; ri++) {
                        bool rowUsesBold = GetTableRowBold(style, ri, headerRowCount, footerStartRowIndex);
                        double originalRowSize = GetTableRowFontSize(style, ri, headerRowCount, footerStartRowIndex, currentOpts.DefaultFontSize);
                        double rowSize = ResolveTableRowShrinkFontSize(tb2, style, ri, cols, colPixel, columnGap, originalRowSize, rowUsesBold, currentOpts);
                        double runFontSizeScale = GetTableRunFontSizeScale(tb2, style, ri, cols, colPixel, columnGap, originalRowSize, rowSize, rowUsesBold, currentOpts);
                        double rowLeading = GetTableLeading(style, rowSize);
                        rowSizes[ri] = rowSize;
                        rowLeadings[ri] = rowLeading;
                        rowBold[ri] = rowUsesBold;
                        rowLines[ri] = new TableCellTextLayout[cols];
                        int maxLines = 1;
                        double maxRequiredHeight = rowLeading + GetTableRowMaxPaddingTop(tb2, style, ri, cols) + GetTableRowMaxPaddingBottom(tb2, style, ri, cols);
                        for (int ci = 0; ci < cols; ci++) {
                            rowLines[ri][ci] = new TableCellTextLayout(new System.Collections.Generic.List<System.Collections.Generic.List<RichSeg>> { new() }, new System.Collections.Generic.List<double> { rowLeading });
                        }

                        var cells = GetTableCellLayouts(tb2, ri, cols);
                        for (int cellIndex = 0; cellIndex < cells.Count; cellIndex++) {
                            TableCellLayout cell = cells[cellIndex];
                            var cellFont = GetTableRowFont(currentOpts, rowUsesBold);
                            double cellWidth = GetTableCellWidth(colPixel, cell.Column, cell.ColumnSpan, columnGap);
                            double innerWidth = Math.Max(1, cellWidth - GetTableCellPaddingLeft(style, ri, cell.Column) - GetTableCellPaddingRight(style, ri, cell.Column));
                            TableCellTextLayout lines = CreateTableCellTextLayout(cell, innerWidth, cellFont, rowSize, rowLeading, currentOpts, runFontSizeScale, style.MinimumShrinkFontSize ?? 6D);
                            rowLines[ri][cell.Column] = lines;
                            if (cell.RowSpan <= 1) {
                                maxLines = Math.Max(maxLines, lines.LineCount);
                                maxRequiredHeight = Math.Max(maxRequiredHeight, MeasureTableCellContentHeight(cell, lines, 0, lines.LineCount, rowLeading, innerWidth) + GetTableCellPaddingTop(style, ri, cell.Column) + GetTableCellPaddingBottom(style, ri, cell.Column));
                            }
                        }

                        rowLineCounts[ri] = maxLines;
                        rowHeights[ri] = ResolveTableRowHeight(style, ri, maxRequiredHeight);
                    }
                    ApplyTableRowSpanHeights(tb2, style, cols, colPixel, rowLines, rowHeights, rowLeadings, columnGap, tableRowGap);

                    System.Collections.Generic.IReadOnlyList<TextRun>? captionRuns = null;
                    System.Collections.Generic.List<System.Collections.Generic.List<RichSeg>>? captionLines = null;
                    System.Collections.Generic.List<double>? captionLineHeights = null;
                    double captionLeading = 0;
                    double captionHeight = 0;
                    if (!string.IsNullOrWhiteSpace(style.Caption)) {
                        double captionSize = style.CaptionFontSize ?? size;
                        captionLeading = captionSize * 1.25;
                        var captionFont = ChooseNormal(currentOpts.DefaultFont);
                        captionRuns = new[] { TextRun.Normal(style.Caption!, style.CaptionColor, captionSize) };
                        var captionWrap = WrapRichRunsCore(captionRuns, tableWidth, captionSize, captionFont, captionLeading, null, DefaultParagraphTabStopWidth, currentOpts);
                        captionLines = captionWrap.Lines;
                        captionLineHeights = captionWrap.LineHeights;
                        captionHeight = MeasureRichLinesHeight(captionLineHeights, captionLines.Count, captionLeading) + style.CaptionSpacingAfter;
                        double firstRowHeight = rowHeights.Length > 0 ? rowHeights[0] : 0;
                        double maxContentHeightForCaption = currentOpts.PageHeight - currentOpts.MarginTop - currentOpts.MarginBottom;
                        if (captionHeight + firstRowHeight > maxContentHeightForCaption + 0.001) {
                            throw new ArgumentException("Table caption and first row exceed the available page content height.");
                        }
                    }

                    items.Add(new ColTable { Block = tb2, Style = style, Columns = cols, ColumnWidths = colPixel, RowLines = rowLines, RowLineCounts = rowLineCounts, RowHeights = rowHeights, RowLeadings = rowLeadings, RowSizes = rowSizes, RowBold = rowBold, Width = tableWidth, Size = size, HeaderRowCount = headerRowCount, RepeatHeaderRowCount = repeatHeaderRowCount, FooterStartRowIndex = footerStartRowIndex, CaptionRuns = captionRuns, CaptionLines = captionLines, CaptionLineHeights = captionLineHeights, CaptionLeading = captionLeading, CaptionHeight = captionHeight });
                } else if (cb is HorizontalRuleBlock hr2) {
                    items.Add(new ColRule { Block = hr2 });
                } else if (cb is ImageBlock ib2) {
                    PdfImageStyle imageStyle = ResolveImageStyle(ib2, currentOpts);
                    double spacingBefore = imageStyle.SpacingBefore;
                    var imageBox = ResolveImageFlowBox(ib2, imageStyle, colWs[i], spacingBefore, imageStyle.SpacingAfter);
                    items.Add(new ColImg { Block = ib2, Style = imageStyle, Width = imageBox.Width, Height = imageBox.Height });
                } else if (cb is ShapeBlock sb2) {
                    items.Add(new ColShape { Block = sb2 });
                } else if (cb is DrawingBlock db2) {
                    items.Add(new ColDrawing { Block = db2 });
                } else if (cb is TextFieldBlock || cb is CheckBoxBlock || cb is ChoiceFieldBlock || cb is RadioButtonGroupBlock) {
                    items.Add(new ColForm { Block = cb });
                } else if (cb is BookmarkBlock bookmark2) {
                    items.Add(new ColBookmark { Block = bookmark2 });
                } else if (cb is SpacerBlock spacer2) {
                    items.Add(new ColSpacer { Block = spacer2 });
                }
            }
            colItems.Add(items);
        }
            return colItems;
        }

        private double MeasureRowKeepTogetherHeight(System.Collections.Generic.List<ColItem> items) {
            double total = 0D;
            foreach (var item in items) {
                total += MeasureColItemFullHeight(item, total);
            }

            return total;
        }

        private double MeasureColKeepWithNextChainHeight(System.Collections.Generic.List<ColItem> items, int startIndex) {
            double total = 0D;
            for (int itemIndex = startIndex; itemIndex < items.Count; itemIndex++) {
                ColItem item = items[itemIndex];
                if (item is ColBookmark) {
                    continue;
                }

                bool keepWithNext = ColItemKeepsWithNext(item);
                if (keepWithNext && item is ColListItem listItem && listItem.IsFirstInKeepWithNextGroup) {
                    double spacingBefore = ResolveColumnSpacingBefore(listItem.SpacingBefore, total + 1D);
                    total += listItem.KeepWithNextGroupHeight - listItem.SpacingBefore + spacingBefore;
                    itemIndex += Math.Max(0, listItem.KeepWithNextGroupItemCount - 1);
                } else {
                    total += keepWithNext
                        ? MeasureColItemFullHeight(item, total + 1D)
                        : MeasureColItemFirstVisualHeight(item);
                }

                if (!keepWithNext) {
                    break;
                }
            }

            return total;
        }

        private bool ColItemKeepsWithNext(ColItem item) {
            if (item is ColPar paragraph) {
                return EffectiveParagraphStyle(paragraph.Block)?.KeepWithNext == true;
            }

            if (item is ColHead heading) {
                return heading.KeepWithNext;
            }

            if (item is ColListItem listItem) {
                return listItem.KeepWithNext && listItem.IsFirstInKeepWithNextGroup;
            }

            if (item is ColPanel panel) {
                return panel.Style.KeepWithNext;
            }

            if (item is ColTable table) {
                return table.Style.KeepWithNext;
            }

            if (item is ColRule rule) {
                return ResolveHorizontalRuleStyle(rule.Block, currentOpts).KeepWithNext;
            }

            if (item is ColImg image) {
                return image.Style.KeepWithNext;
            }

            if (item is ColShape shape) {
                return ResolveDrawingStyle(shape.Block, currentOpts).KeepWithNext;
            }

            if (item is ColDrawing drawing) {
                return ResolveDrawingStyle(drawing.Block, currentOpts).KeepWithNext;
            }

            return false;
        }

        private double MeasureColItemFullHeight(ColItem item, double consumedBefore) {
            if (item is ColPar paragraph) {
                PdfParagraphStyle? paragraphStyle = EffectiveParagraphStyle(paragraph.Block);
                return ResolveColumnSpacingBefore(GetParagraphSpacingBefore(paragraphStyle), consumedBefore) + paragraph.Heights.Sum() + GetParagraphSpacingAfter(paragraphStyle, paragraph.Leading);
            }

            if (item is ColHead heading) {
                return ResolveColumnSpacingBefore(heading.SpacingBefore, consumedBefore) + MeasureRichLinesHeight(heading.Heights, heading.Lines.Count, heading.Leading) + heading.SpacingAfter;
            }

            if (item is ColListItem listItem) {
                return ResolveColumnSpacingBefore(listItem.SpacingBefore, consumedBefore) + MeasureRichLinesHeight(listItem.Heights, listItem.Lines.Count, listItem.Leading) + listItem.SpacingAfter;
            }

            if (item is ColPanel panel) {
                return ResolveColumnSpacingBefore(panel.Style.SpacingBefore, consumedBefore) + panel.Style.PaddingY + panel.Heights.Sum() + panel.Style.PaddingY + panel.Style.SpacingAfter;
            }

            if (item is ColTable table) {
                return ResolveColumnSpacingBefore(table.Style.SpacingBefore, consumedBefore) + table.CaptionHeight + GetTableRowsHeight(table.RowHeights, 0, table.RowHeights.Length, GetTableCellSpacing(table.Style)) + table.Style.SpacingAfter;
            }

            if (item is ColRule rule) {
                PdfHorizontalRuleStyle ruleStyle = ResolveHorizontalRuleStyle(rule.Block, currentOpts);
                ValidateHorizontalRule(ruleStyle);
                return ResolveColumnSpacingBefore(ruleStyle.SpacingBefore, consumedBefore) + ruleStyle.Thickness + ruleStyle.SpacingAfter;
            }

            if (item is ColImg image) {
                PdfImageStyle imageStyle = image.Style;
                return ResolveColumnSpacingBefore(imageStyle.SpacingBefore, consumedBefore) + image.Height + imageStyle.SpacingAfter;
            }

            if (item is ColShape shape) {
                PdfDrawingStyle shapeStyle = ResolveDrawingStyle(shape.Block, currentOpts);
                return ResolveColumnSpacingBefore(shapeStyle.SpacingBefore, consumedBefore) + shape.Block.Shape.Height + shapeStyle.SpacingAfter;
            }

            if (item is ColDrawing drawing) {
                PdfDrawingStyle drawingStyle = ResolveDrawingStyle(drawing.Block, currentOpts);
                return ResolveColumnSpacingBefore(drawingStyle.SpacingBefore, consumedBefore) + drawing.Block.Drawing.Height + drawingStyle.SpacingAfter;
            }

            if (item is ColForm form) {
                return ResolveColumnSpacingBefore(GetFormFieldSpacingBefore(form.Block), consumedBefore) + GetFormFieldHeight(form.Block) + GetFormFieldSpacingAfter(form.Block);
            }

            if (item is ColSpacer spacerItem) {
                return spacerItem.Block.Height;
            }

            return 0D;
        }

        private double MeasureColItemFirstVisualHeight(ColItem item) {
            if (item is ColPar paragraph) {
                PdfParagraphStyle? paragraphStyle = EffectiveParagraphStyle(paragraph.Block);
                return GetParagraphSpacingBefore(paragraphStyle) + (paragraph.Heights.Count == 0 ? 0D : paragraph.Heights[0]);
            }

            if (item is ColHead heading) {
                return heading.SpacingBefore + (heading.Lines.Count == 0 ? 0D : GetRichLineHeight(heading.Heights, 0, heading.Leading));
            }

            if (item is ColListItem listItem) {
                return listItem.SpacingBefore + (listItem.Lines.Count == 0 ? 0D : GetRichLineHeight(listItem.Heights, 0, listItem.Leading));
            }

            if (item is ColPanel panel) {
                return panel.Style.SpacingBefore + panel.Style.PaddingY + (panel.Heights.Count == 0 ? 0D : panel.Heights[0]) + panel.Style.PaddingY;
            }

            if (item is ColTable table) {
                double firstRowHeight = table.RowHeights.Length == 0 ? 0D : table.RowHeights[0];
                return table.Style.SpacingBefore + table.CaptionHeight + firstRowHeight;
            }

            if (item is ColRule rule) {
                PdfHorizontalRuleStyle ruleStyle = ResolveHorizontalRuleStyle(rule.Block, currentOpts);
                return ruleStyle.SpacingBefore + ruleStyle.Thickness + ruleStyle.SpacingAfter;
            }

            if (item is ColImg image) {
                PdfImageStyle imageStyle = image.Style;
                return imageStyle.SpacingBefore + image.Height + imageStyle.SpacingAfter;
            }

            if (item is ColShape shape) {
                PdfDrawingStyle shapeStyle = ResolveDrawingStyle(shape.Block, currentOpts);
                return shapeStyle.SpacingBefore + shape.Block.Shape.Height + shapeStyle.SpacingAfter;
            }

            if (item is ColDrawing drawing) {
                PdfDrawingStyle drawingStyle = ResolveDrawingStyle(drawing.Block, currentOpts);
                return drawingStyle.SpacingBefore + drawing.Block.Drawing.Height + drawingStyle.SpacingAfter;
            }

            if (item is ColForm form) {
                return GetFormFieldSpacingBefore(form.Block) + GetFormFieldHeight(form.Block) + GetFormFieldSpacingAfter(form.Block);
            }

            if (item is ColSpacer spacerItem) {
                return spacerItem.Block.Height;
            }

            return 0D;
        }

    }
}
