using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private delegate void RowFragmentDecorator(
            StringBuilder content,
            int insertionIndex,
            double top,
            double bottom,
            bool isFirstFragment,
            bool isLastFragment);

        private void RenderRowFlowBlock(
            RowBlock rb,
            IPdfBlock? nextBlock,
            System.Collections.Generic.IList<IPdfBlock> blockList,
            int blockIndex,
            RowFragmentDecorator? fragmentDecorator = null) {
            double contentWidth = currentOpts.PageWidth - currentOpts.MarginLeft - currentOpts.MarginRight;
            int ncols = rb.Columns.Count;
            PdfRowStyle? rowStyle = rb.StyleSnapshot ?? currentOpts.DefaultRowStyleSnapshot;
            double rowGap = rb.GapOverride ?? rowStyle?.Gap ?? PdfRowStyle.DefaultGap;
            double rowSpacingBefore = ResolveTopLevelSpacingBefore(rowStyle?.SpacingBefore ?? 0D);
            double rowSpacingAfter = rowStyle?.SpacingAfter ?? 0D;
            double totalGap = rowGap * Math.Max(0, ncols - 1);
            if (totalGap >= contentWidth) {
                throw new ArgumentException("Row column gaps must be smaller than the available page content width.");
            }

            double columnAreaWidth = contentWidth - totalGap;
            double[] colXs = new double[ncols];
            double[] colWs = new double[ncols];
            double xAcc = currentOpts.MarginLeft;
            for (int i = 0; i < ncols; i++) { double wCol = Math.Max(0, columnAreaWidth * (rb.Columns[i].WidthPercent / 100.0)); colXs[i] = xAcc; colWs[i] = wCol; xAcc += wCol + rowGap; }

            void DrawRowColumnSeparators(double topY, double bottomY) {
                if (ncols <= 1 || rowStyle?.ColumnSeparatorColor == null || rowStyle.ColumnSeparatorWidth <= 0D || topY - bottomY <= 0.001D) {
                    return;
                }

                for (int boundary = 0; boundary < ncols - 1; boundary++) {
                    double separatorX = colXs[boundary] + colWs[boundary] + (rowGap / 2D);
                    DrawVLine(sb, rowStyle.ColumnSeparatorColor.Value, rowStyle.ColumnSeparatorWidth, separatorX, topY, bottomY, emitGeneratedStructure);
                }

                pageDirty = true;
            }

            var colStates = CreateRowColumnStates(ncols);
            var colItems = BuildRowColumnItems(rb, colWs);
            var columnListStructureElementIndexes = new int?[ncols];
            var columnListStructurePages = new LayoutResult.Page?[ncols];
            var columnActiveListGroupIds = new int[ncols];
            for (int i = 0; i < ncols; i++) {
                columnActiveListGroupIds[i] = -1;
            }

            static System.Collections.Generic.List<(int idx, int line, int subline)> CreateRowColumnStates(int columnCount) {
                var states = new System.Collections.Generic.List<(int idx, int line, int subline)>(columnCount);
                for (int i = 0; i < columnCount; i++) {
                    states.Add((0, 0, 0));
                }

                return states;
            }

            double? rowContentHeightCache = null;
            double GetRowContentHeight() {
                if (rowContentHeightCache.HasValue) {
                    return rowContentHeightCache.Value;
                }

                double measuredHeight = 0D;
                foreach (var items in colItems) {
                    measuredHeight = Math.Max(measuredHeight, MeasureRowKeepTogetherHeight(items));
                }

                rowContentHeightCache = measuredHeight;
                return measuredHeight;
            }

            if (rowStyle?.KeepTogether == true) {
                double rowContentHeight = GetRowContentHeight();
                double rowKeepHeight = rowSpacingBefore + rowContentHeight + rowSpacingAfter;
                double availableHeight = currentOpts.PageHeight - currentOpts.MarginTop - currentOpts.MarginBottom;
                if (rowKeepHeight > availableHeight + 0.001) {
                    throw new ArgumentException("Row height exceeds the available page content height.");
                }

                if (y < yStart - 0.001 && y - rowKeepHeight < currentOpts.MarginBottom) {
                    NewPage();
                    rowSpacingBefore = 0D;
                }
            }

            if (rowStyle?.KeepWithNext == true && nextBlock != null) {
                double rowContentHeight = GetRowContentHeight();
                double rowHeight = rowSpacingBefore + rowContentHeight + rowSpacingAfter;
                double nextHeight = MeasureKeepWithNextChainHeight(blockList, blockIndex + 1, currentOpts.MarginLeft, width, currentOpts.DefaultFontSize);
                double keepHeight = rowHeight + nextHeight;
                double availableHeight = currentOpts.PageHeight - currentOpts.MarginTop - currentOpts.MarginBottom;
                if (nextHeight > 0.001 && rowHeight <= availableHeight + 0.001 && keepHeight <= availableHeight + 0.001 && y < yStart - 0.001 && y - keepHeight < currentOpts.MarginBottom) {
                    NewPage();
                    rowSpacingBefore = 0D;
                }
            }

            if (rowSpacingBefore > 0) {
                if (y - rowSpacingBefore < currentOpts.MarginBottom) {
                    NewPage();
                    rowSpacingBefore = 0D;
                }

                if (rowSpacingBefore > 0) y -= rowSpacingBefore;
            }

            bool AnyRemaining() {
                for (int i = 0; i < ncols; i++) if (colStates[i].idx < colItems[i].Count) return true; return false;
            }

            int rowColumnFlowGuard = 0;
            bool isFirstFragment = true;
            while (AnyRemaining()) {
                rowColumnFlowGuard++;
                if (rowColumnFlowGuard > 10000) {
                    throw new InvalidOperationException("Row column layout did not make forward progress.");
                }

                double avail = y - currentOpts.MarginBottom;
                if (avail <= 0.5) { NewPage(); avail = y - currentOpts.MarginBottom; }

                int fragmentInsertionIndex = sb.Length;
                double maxConsumed = 0;
                bool anyColumnAdvanced = false;
                for (int ci = 0; ci < ncols; ci++) {
                    var items = colItems[ci];
                    var (idx, line, subline) = colStates[ci];
                    var startState = (idx, line, subline);
                    double xCol = colXs[ci];
                    double wCol = colWs[ci];
                    double yCol = y;
                    double consumed = 0;
                    double remain = avail;
                    while (idx < items.Count && remain > 0.1) {
                        var it = items[idx];
                        if (it is ColListItem currentListItem) {
                            if (columnActiveListGroupIds[ci] != currentListItem.ListGroupId) {
                                columnActiveListGroupIds[ci] = currentListItem.ListGroupId;
                                columnListStructureElementIndexes[ci] = null;
                                columnListStructurePages[ci] = null;
                            }
                        } else {
                            columnActiveListGroupIds[ci] = -1;
                            columnListStructureElementIndexes[ci] = null;
                            columnListStructurePages[ci] = null;
                        }

                        if (it is ColPar par) {
                            var pblock = par.Block;
                            var lines = par.Lines;
                            var heights = par.Heights;
                            double leading = par.Leading;
                            double size = par.Size;
                            PdfParagraphStyle? paragraphStyle = EffectiveParagraphStyle(pblock);
                            double spacingBefore = line == 0 && consumed > 0.001 ? GetParagraphSpacingBefore(paragraphStyle) : 0;
                            double spacingAfter = GetParagraphSpacingAfter(paragraphStyle, leading);
                            if (paragraphStyle?.KeepWithNext == true && line == 0 && idx + 1 < items.Count) {
                                double nextHeight = MeasureColKeepWithNextChainHeight(items, idx + 1);
                                double keepHeight = spacingBefore + heights.Sum() + spacingAfter + nextHeight;
                                double availableHeight = currentOpts.PageHeight - currentOpts.MarginTop - currentOpts.MarginBottom;
                                if (nextHeight > 0.001 && keepHeight <= availableHeight + 0.001 && keepHeight > remain + 0.001) {
                                    if (consumed > 0) break;
                                    remain = 0;
                                    break;
                                }
                            }

                            if (paragraphStyle?.KeepTogether == true && line == 0) {
                                double paragraphHeight = spacingBefore + heights.Sum() + spacingAfter;
                                double availableHeight = currentOpts.PageHeight - currentOpts.MarginTop - currentOpts.MarginBottom;
                                if (paragraphHeight > availableHeight + 0.001) {
                                    throw new ArgumentException("Paragraph height exceeds the available page content height.");
                                }

                                if (paragraphHeight > remain + 0.001) {
                                    if (consumed > 0) break;
                                    remain = 0;
                                    break;
                                }
                            }

                            double availableForLines = remain - spacingBefore;
                            if (availableForLines < 0) {
                                if (consumed > 0) break;
                                remain = 0;
                                break;
                            }

                            int start = line;
                            int take = 0; double hsum = 0;
                            for (int li2 = start; li2 < lines.Count; li2++) {
                                double hAdd = heights[li2];
                                if (hsum + hAdd + (li2 == lines.Count - 1 ? spacingAfter : 0) > availableForLines) break;
                                hsum += hAdd; take++;
                            }

                            if (TryApplyWidowControl(paragraphStyle, lines.Count, start, ref take, ref hsum, heights, consumed > 0 || y < yStart - 0.001)) {
                                break;
                            }

                            if (take == 0) break;
                            if (spacingBefore > 0) {
                                yCol -= spacingBefore;
                                remain -= spacingBefore;
                                consumed += spacingBefore;
                            }

                            var sliceLines = new System.Collections.Generic.List<System.Collections.Generic.List<RichSeg>>();
                            var sliceHeights = new System.Collections.Generic.List<double>();
                            for (int k = 0; k < take; k++) { sliceLines.Add(lines[start + k]); sliceHeights.Add(heights[start + k]); }
                            pageDirty = true;
                            var paragraphFont = ChooseNormal(currentOpts.DefaultFont);
                            int? markedContentId = RegisterTextStructureElement("P");
                            WriteRichParagraph(sb, pblock, sliceLines, sliceHeights, currentOpts, FirstTextBaselineFromTop(paragraphFont, size, yCol), size, leading, currentPage!.Annotations, xCol + par.XOffset, par.TextWidth, start == 0 ? xCol + par.FirstLineXOffset : null, start == 0 ? par.FirstLineTextWidth : null, "P", markedContentId, currentPage);
                            MarkRichFonts(pblock.Runs);
                            yCol -= hsum; remain -= hsum; consumed += hsum; line += take;
                            if (line >= lines.Count) { double space = spacingAfter; if (space <= remain) { yCol -= space; remain -= space; consumed += space; } idx++; line = 0; }
                        } else if (it is ColHead ch) {
                            var hb2 = ch.Block;
                            var lines = ch.Lines;
                            var heights = ch.Heights;
                            double leading = ch.Leading;
                            double size = ch.Size;
                            double spacingBefore = (consumed > 0.001 || ch.ApplySpacingBeforeAtTop) ? ch.SpacingBefore : 0D;
                            double textHeight = MeasureRichLinesHeight(heights, lines.Count, leading);
                            double needed = spacingBefore + textHeight + ch.SpacingAfter;
                            if (ch.KeepWithNext && idx + 1 < items.Count) {
                                double nextHeight = MeasureColKeepWithNextChainHeight(items, idx + 1);
                                double keepHeight = needed + nextHeight;
                                double availableHeight = currentOpts.PageHeight - currentOpts.MarginTop - currentOpts.MarginBottom;
                                if (nextHeight > 0.001 && keepHeight <= availableHeight + 0.001 && keepHeight > remain + 0.001) {
                                    if (consumed > 0) break;
                                    remain = 0;
                                    break;
                                }
                            }

                            if (needed > remain && consumed > 0) break;
                            if (needed > remain && consumed == 0) { remain = 0; break; }
                            if (spacingBefore > 0) {
                                yCol -= spacingBefore;
                                remain -= spacingBefore;
                                consumed += spacingBefore;
                            }

                            if (currentOpts.CreateOutlineFromHeadings) {
                                currentPage!.Bookmarks.Add(new PageBookmark { Level = hb2.Level, Title = hb2.Text, Y = yCol });
                            }
                            var headingFont = ch.Bold ? ChooseBold(ChooseNormal(currentOpts.DefaultFont)) : ChooseNormal(currentOpts.DefaultFont);
                            double firstBaseline = FirstTextBaselineFromTop(headingFont, size, yCol);
                            string structureType = "H" + hb2.Level.ToString(CultureInfo.InvariantCulture);
                            bool hasLinkTarget = !string.IsNullOrEmpty(hb2.LinkUri) || !string.IsNullOrEmpty(hb2.LinkDestinationName);
                            int? linkStructElementIndex = null;
                            string markedStructureType = structureType;
                            int? markedContentId;
                            if (hasLinkTarget && emitGeneratedStructure && currentPage != null) {
                                int? headingElementIndex = RegisterStructureContainer(structureType);
                                linkStructElementIndex = currentPage.StructElements.Count;
                                markedStructureType = "Link";
                                markedContentId = RegisterTextStructureElement(markedStructureType, headingElementIndex);
                            } else {
                                markedContentId = RegisterTextStructureElement(structureType);
                            }

                            AddHeadingLinkAnnotations(hb2, lines, headingFont, size, leading, xCol, wCol, firstBaseline, linkStructElementIndex);
                            WriteRichParagraph(sb, new RichParagraphBlock(ch.Runs, hb2.Align, ch.Color), lines, heights, currentOpts, firstBaseline, size, leading, currentPage!.Annotations, xCol, wCol, structureType: markedStructureType, markedContentId: markedContentId, structurePage: currentPage);
                            MarkRichFonts(ch.Runs);
                            if (ch.Bold) {
                                currentPage!.UsedBold = true;
                                usedBold = true;
                            }
                            double consumedHeight = textHeight + ch.SpacingAfter;
                            yCol -= consumedHeight; remain -= consumedHeight; consumed += consumedHeight; idx++;
                        } else if (it is ColListItem listItem) {
                            var lines = listItem.Lines;
                            double leading = listItem.Leading;
                            double spacingBefore = line == 0 ? ResolveColumnSpacingBefore(listItem.SpacingBefore, consumed) : 0D;
                            if (line == 0 && listItem.KeepTogether && listItem.IsFirstInKeepGroup) {
                                double keepGroupHeight = listItem.KeepGroupHeight - listItem.SpacingBefore + spacingBefore;
                                double availableHeight = currentOpts.PageHeight - currentOpts.MarginTop - currentOpts.MarginBottom;
                                if (keepGroupHeight > availableHeight + 0.001) {
                                    throw new ArgumentException("List height exceeds the available page content height.");
                                }

                                if (keepGroupHeight > remain + 0.001) {
                                    if (consumed > 0) break;
                                    remain = 0;
                                    break;
                                }
                            }

                            if (line == 0 && listItem.KeepWithNext && listItem.IsFirstInKeepWithNextGroup) {
                                int nextItemIndex = idx + listItem.KeepWithNextGroupItemCount;
                                if (nextItemIndex < items.Count) {
                                    double nextHeight = MeasureColKeepWithNextChainHeight(items, nextItemIndex);
                                    double keepHeight = listItem.KeepWithNextGroupHeight - listItem.SpacingBefore + spacingBefore + nextHeight;
                                    double availableHeight = currentOpts.PageHeight - currentOpts.MarginTop - currentOpts.MarginBottom;
                                    if (nextHeight > 0.001 && keepHeight <= availableHeight + 0.001 && keepHeight > remain + 0.001) {
                                        if (consumed > 0) break;
                                        remain = 0;
                                        break;
                                    }
                                }
                            }

                            if (line == 0 && spacingBefore > 0) {
                                if (spacingBefore > remain && consumed > 0) break;
                                if (spacingBefore > remain && consumed == 0) { remain = 0; break; }
                                yCol -= spacingBefore;
                                remain -= spacingBefore;
                                consumed += spacingBefore;
                            }

                            double availableForLines = remain;
                            int start = line;
                            int take = 0;
                            double hsum = 0;
                            for (int li2 = start; li2 < lines.Count; li2++) {
                                double lineHeight = GetRichLineHeight(listItem.Heights, li2, leading);
                                if (hsum + lineHeight > availableForLines) break;
                                hsum += lineHeight;
                                take++;
                            }
                            if (take == 0) break;

                            var sliceLines = new System.Collections.Generic.List<System.Collections.Generic.List<RichSeg>>(take);
                            var sliceHeights = new System.Collections.Generic.List<double>(take);
                            for (int k = 0; k < take; k++) {
                                sliceLines.Add(lines[start + k]);
                                sliceHeights.Add(GetRichLineHeight(listItem.Heights, start + k, leading));
                            }

                            pageDirty = true;
                            var listFont = ChooseNormal(currentOpts.DefaultFont);
                            double baselineY = FirstTextBaselineFromTop(listFont, listItem.Size, yCol);
                            int? listElementIndex = line == 0 || listItem.StructureElement == null
                                ? EnsurePageStructureContainer("L", ref columnListStructureElementIndexes[ci], ref columnListStructurePages[ci])
                                : null;
                            int? listItemElementIndex = line == 0 || listItem.StructureElement == null
                                ? RegisterStructureContainer("LI", listElementIndex)
                                : null;
                            if (listItemElementIndex.HasValue && currentPage != null) {
                                listItem.StructureElement = currentPage.StructElements[listItemElementIndex.Value];
                            }

                            if (line == 0) {
                                if (!string.IsNullOrEmpty(listItem.BookmarkName)) {
                                    AddNamedDestinationName(listItem.BookmarkName!, yCol);
                                }

                                var markerLines = new System.Collections.Generic.List<string>(1) { listItem.Marker };
                                int? labelMarkedContentId = RegisterTextStructureElement("Lbl", listItemElementIndex);
                                if (listItem.MarkerNamedFont.HasValue) {
                                    currentPage!.UsedNamedFonts.Add(listItem.MarkerNamedFont.Value);
                                } else {
                                    MarkSimpleFont(listItem.MarkerFont);
                                }

                                WriteLinesInternal(
                                    GetFontResourceName(listItem.MarkerFont, listItem.MarkerNamedFont, ChooseNormal(currentOpts.DefaultFont)),
                                    listItem.MarkerSize,
                                    leading,
                                    xCol + listItem.MarkerXOffset,
                                    listItem.MarkerWidth,
                                    baselineY,
                                    markerLines,
                                    listItem.MarkerAlign,
                                    listItem.MarkerColor ?? listItem.Color,
                                    applyBaselineTweak: true,
                                    structureType: "Lbl",
                                    markedContentId: labelMarkedContentId,
                                    namedFont: listItem.MarkerNamedFont);
                            }

                            int? bodyMarkedContentId = line == 0 || listItem.StructureElement == null
                                ? RegisterTextStructureElement("LBody", listItemElementIndex)
                                : RegisterTextStructureElement("LBody", listItem.StructureElement);
                            WriteRichParagraph(sb, new RichParagraphBlock(listItem.Runs, listItem.TextAlign, listItem.Color), sliceLines, sliceHeights, currentOpts, baselineY, listItem.Size, leading, currentPage!.Annotations, xCol + listItem.TextXOffset, listItem.TextWidth, structureType: "LBody", markedContentId: bodyMarkedContentId, structurePage: currentPage);
                            MarkRichFonts(listItem.Runs);
                            yCol -= hsum;
                            remain -= hsum;
                            consumed += hsum;
                            line += take;
                            if (line >= lines.Count) {
                                double space = listItem.SpacingAfter;
                                if (space <= remain) {
                                    yCol -= space;
                                    remain -= space;
                                    consumed += space;
                                }

                                idx++;
                                line = 0;
                            }
                        } else if (it is ColPanel panel) {
                            var pblock = panel.Block;
                            var panelStyle = panel.Style;
                            var lines = panel.Lines;
                            var heights = panel.Heights;
                            double xPanel = xCol + panel.XOffset;
                            double spacingBefore = line == 0 ? ResolveColumnSpacingBefore(panelStyle.SpacingBefore, consumed) : 0D;
                            if (line == 0 && panelStyle.KeepWithNext && idx + 1 < items.Count) {
                                double nextHeight = MeasureColKeepWithNextChainHeight(items, idx + 1);
                                double panelHeight = spacingBefore + panelStyle.PaddingY + heights.Sum() + panelStyle.PaddingY + panelStyle.SpacingAfter;
                                double keepHeight = panelHeight + nextHeight;
                                double availableHeight = currentOpts.PageHeight - currentOpts.MarginTop - currentOpts.MarginBottom;
                                if (nextHeight > 0.001 && keepHeight <= availableHeight + 0.001 && keepHeight > remain + 0.001) {
                                    if (consumed > 0) break;
                                    remain = 0;
                                    break;
                                }
                            }

                            if (line == 0 && spacingBefore > 0) {
                                if (spacingBefore > remain && consumed > 0) break;
                                if (spacingBefore > remain && consumed == 0) { remain = 0; break; }
                                yCol -= spacingBefore;
                                remain -= spacingBefore;
                                consumed += spacingBefore;
                            }

                            double keepTogetherTextHeight = heights.Sum();
                            double keepTogetherPanelHeight = panelStyle.PaddingY + keepTogetherTextHeight + panelStyle.PaddingY;
                            double keepTogetherAvailableHeight = currentOpts.PageHeight - currentOpts.MarginTop - currentOpts.MarginBottom;
                            if (panelStyle.KeepTogether && panelStyle.PaddingY + panelStyle.PaddingY > keepTogetherAvailableHeight + 0.001D) {
                                throw new ArgumentException("Panel vertical padding and first line height exceed the available page content height.");
                            }

                            bool keepPanelTogether = panelStyle.KeepTogether && keepTogetherPanelHeight <= keepTogetherAvailableHeight + 0.001;
                            if (keepPanelTogether) {
                                double panelHeight = keepTogetherPanelHeight;

                                if (panelHeight > remain && consumed > 0) break;
                                if (panelHeight > remain && consumed == 0) { remain = 0; break; }

                                double panelTop = yCol;
                                double panelBottom = yCol - panelHeight;
                                if (panelStyle.Background.HasValue) { pageDirty = true; DrawRowFill(sb, panelStyle.Background.Value, xPanel, panelBottom, panel.PanelWidth, panelTop - panelBottom, emitGeneratedStructure); }
                                if (DrawPanelBorder(sb, panelStyle, xPanel, panelBottom, panel.PanelWidth, panelTop - panelBottom, emitGeneratedStructure)) { pageDirty = true; }
                                pageDirty = true;
                                int? panelMarkedContentId = RegisterTextStructureElement("P");
                                WriteRichParagraph(sb, new RichParagraphBlock(pblock.Runs, pblock.Align, pblock.DefaultColor), lines, heights, currentOpts, panelTop - panelStyle.PaddingY - panel.FirstBaselineOffset, panel.Size, panel.Leading, currentPage!.Annotations, xPanel + panelStyle.PaddingX, panel.TextWidth, structureType: "P", markedContentId: panelMarkedContentId, structurePage: currentPage);
                                MarkRichFonts(pblock.Runs);

                                yCol = panelBottom;
                                remain -= panelHeight;
                                consumed += panelHeight;
                                if (panelStyle.SpacingAfter > 0 && panelStyle.SpacingAfter <= remain) {
                                    yCol -= panelStyle.SpacingAfter;
                                    remain -= panelStyle.SpacingAfter;
                                    consumed += panelStyle.SpacingAfter;
                                }
                                idx++;
                                line = 0;
                            } else {
                                int start = line;
                                double topPad = start == 0 ? panelStyle.PaddingY : 0;
                                double minLine = heights[start];
                                if (remain < topPad + minLine) {
                                    EnsurePanelSegmentCanFitLine(topPad, minLine);
                                    if (consumed > 0) break;
                                    remain = 0;
                                    break;
                                }

                                double roomForText = remain - topPad - panelStyle.PaddingY;
                                if (roomForText < minLine) {
                                    if (start == lines.Count - 1) {
                                        EnsurePanelSegmentCanFitLine(topPad + panelStyle.PaddingY, minLine);
                                        if (consumed > 0) break;
                                        remain = 0;
                                        break;
                                    }

                                    roomForText = remain - topPad;
                                }

                                int take = 0;
                                double hsum = 0;
                                for (int k = start; k < lines.Count; k++) {
                                    double h = heights[k];
                                    if (hsum + h > roomForText) break;
                                    hsum += h;
                                    take++;
                                }

                                if (take == 0) {
                                    EnsurePanelSegmentCanFitLine(topPad, minLine);
                                    break;
                                }

                                bool lastSeg = start + take >= lines.Count;
                                if (lastSeg && topPad + hsum + panelStyle.PaddingY > remain + 0.001D) {
                                    if (take > 1) {
                                        take--;
                                        hsum -= heights[start + take];
                                        lastSeg = false;
                                    } else {
                                        EnsurePanelSegmentCanFitLine(topPad + panelStyle.PaddingY, minLine);
                                        if (consumed > 0) break;
                                        remain = 0;
                                        break;
                                    }
                                }

                                double panelTop = yCol;
                                double usedBottomPad = lastSeg ? panelStyle.PaddingY : Math.Max(0, remain - (topPad + hsum));
                                double panelBottom = yCol - (topPad + hsum + usedBottomPad);
                                if (panelStyle.Background.HasValue) { pageDirty = true; DrawRowFill(sb, panelStyle.Background.Value, xPanel, panelBottom, panel.PanelWidth, panelTop - panelBottom, emitGeneratedStructure); }
                                if (DrawPanelBorder(sb, panelStyle, xPanel, panelBottom, panel.PanelWidth, panelTop - panelBottom, emitGeneratedStructure)) { pageDirty = true; }

                                var sliceLines = new System.Collections.Generic.List<System.Collections.Generic.List<RichSeg>>();
                                var sliceHeights = new System.Collections.Generic.List<double>();
                                for (int k = 0; k < take; k++) {
                                    sliceLines.Add(lines[start + k]);
                                    sliceHeights.Add(heights[start + k]);
                                }

                                pageDirty = true;
                                int? panelMarkedContentId = RegisterTextStructureElement("P");
                                WriteRichParagraph(sb, new RichParagraphBlock(pblock.Runs, pblock.Align, pblock.DefaultColor), sliceLines, sliceHeights, currentOpts, panelTop - topPad - panel.FirstBaselineOffset, panel.Size, panel.Leading, currentPage!.Annotations, xPanel + panelStyle.PaddingX, panel.TextWidth, structureType: "P", markedContentId: panelMarkedContentId, structurePage: currentPage);
                                MarkRichFonts(pblock.Runs);

                                double segmentHeight = panelTop - panelBottom;
                                yCol = panelBottom;
                                remain -= segmentHeight;
                                consumed += segmentHeight;
                                line += take;
                                if (line >= lines.Count) {
                                    if (panelStyle.SpacingAfter > 0 && panelStyle.SpacingAfter <= remain) {
                                        yCol -= panelStyle.SpacingAfter;
                                        remain -= panelStyle.SpacingAfter;
                                        consumed += panelStyle.SpacingAfter;
                                    }
                                    idx++;
                                    line = 0;
                                } else {
                                    break;
                                }
                            }
                        } else if (it is ColTable table) {
                            var tbColumn = table.Block;
                            var tableStyle = table.Style;
                            double padLeft = GetTableCellPaddingLeft(tableStyle);
                            double padRight = GetTableCellPaddingRight(tableStyle);
                            double padTop = GetTableCellPaddingTop(tableStyle);
                            double padBottom = GetTableCellPaddingBottom(tableStyle);
                            double columnGap = GetTableCellSpacing(tableStyle);
                            double columnTableRowGap = columnGap;
                            double xTable = ResolveTableX(tbColumn.Align, tableStyle, xCol, wCol, table.Width);

                            double maxContentHeight = currentOpts.PageHeight - currentOpts.MarginTop - currentOpts.MarginBottom;
                            double tableSpacingBefore = line == 0 && consumed > 0.001 ? tableStyle.SpacingBefore : 0D;
                            if (line == 0 && tableStyle.KeepTogether) {
                                double keepHeight = tableSpacingBefore + table.CaptionHeight + GetTableRowsHeight(table.RowHeights, 0, table.RowHeights.Length, columnTableRowGap) + tableStyle.SpacingAfter;
                                if (keepHeight > maxContentHeight + 0.001) {
                                    throw new ArgumentException("Table height exceeds the available page content height.");
                                }

                                if (keepHeight > remain + 0.001) {
                                    if (consumed > 0) break;
                                    remain = 0;
                                    break;
                                }
                            }

                            if (line == 0 && tableStyle.KeepWithNext && idx + 1 < items.Count) {
                                double tableHeight = tableSpacingBefore + table.CaptionHeight + GetTableRowsHeight(table.RowHeights, 0, table.RowHeights.Length, columnTableRowGap) + tableStyle.SpacingAfter;
                                double nextHeight = MeasureColKeepWithNextChainHeight(items, idx + 1);
                                double keepHeight = tableHeight + nextHeight;
                                if (nextHeight > 0.001 && tableHeight <= maxContentHeight + 0.001 && keepHeight <= maxContentHeight + 0.001 && keepHeight > remain + 0.001) {
                                    if (consumed > 0) break;
                                    remain = 0;
                                    break;
                                }
                            }

                            if (line == 0 && consumed > 0.001) {
                                int minimumFirstPageBodyRows = Math.Min(
                                    tableStyle.MinimumBodyRowsOnFirstPage,
                                    Math.Max(0, table.FooterStartRowIndex - table.HeaderRowCount));
                                if (minimumFirstPageBodyRows > 0) {
                                    int firstPageRowCount = table.HeaderRowCount + minimumFirstPageBodyRows;
                                    double firstPageGroupHeight =
                                        tableSpacingBefore +
                                        table.CaptionHeight +
                                        GetTableRowsHeight(table.RowHeights, 0, firstPageRowCount, columnTableRowGap);
                                    if (firstPageGroupHeight <= maxContentHeight + 0.001 &&
                                        firstPageGroupHeight > remain + 0.001) {
                                        break;
                                    }
                                }
                            }

                            if (line == 0 && tableSpacingBefore > 0) {
                                if (tableSpacingBefore > remain && consumed > 0) break;
                                if (tableSpacingBefore > remain && consumed == 0) { remain = 0; break; }
                                yCol -= tableSpacingBefore;
                                remain -= tableSpacingBefore;
                                consumed += tableSpacingBefore;
                            }

                            int? tableStructureElementIndex = null;
                            LayoutResult.Page? tableStructurePage = null;
                            int? EnsureTableStructureElement() {
                                if (!emitGeneratedStructure || currentPage == null) {
                                    return null;
                                }

                                if (!ReferenceEquals(tableStructurePage, currentPage)) {
                                    tableStructurePage = currentPage;
                                    tableStructureElementIndex = RegisterStructureContainer("Table", alternativeText: tableStyle.AlternativeText);
                                }

                                return tableStructureElementIndex;
                            }

                            if (line == 0 && table.CaptionRuns != null && table.CaptionLines != null && table.CaptionLineHeights != null) {
                                double firstRowHeight = table.RowHeights.Length > 0 ? table.RowHeights[0] : 0;
                                double neededWithFirstRow = table.CaptionHeight + firstRowHeight;
                                if (neededWithFirstRow > maxContentHeight + 0.001) {
                                    throw new ArgumentException("Table caption and first row exceed the available page content height.");
                                }
                                if (neededWithFirstRow > remain && consumed > 0) break;
                                if (neededWithFirstRow > remain && consumed == 0) { remain = 0; break; }

                                double captionSize = tableStyle.CaptionFontSize ?? table.Size;
                                var captionFont = ChooseNormal(currentOpts.DefaultFont);
                                pageDirty = true;
                                int? captionMarkedContentId = RegisterTextStructureElement("Caption", EnsureTableStructureElement());
                                MarkRichFonts(table.CaptionRuns);
                                WriteRichParagraph(sb, new RichParagraphBlock(table.CaptionRuns, tableStyle.CaptionAlign, tableStyle.CaptionColor), table.CaptionLines, table.CaptionLineHeights, currentOpts, FirstTextBaselineFromTop(captionFont, captionSize, yCol), captionSize, table.CaptionLeading, currentPage!.Annotations, xTable, table.Width, structureType: "Caption", markedContentId: captionMarkedContentId, structurePage: currentPage);
                                yCol -= table.CaptionHeight;
                                remain -= table.CaptionHeight;
                                consumed += table.CaptionHeight;
                            }

                            double repeatHeaderHeight = 0;
                            for (int headerIndex = 0; headerIndex < table.RepeatHeaderRowCount; headerIndex++) {
                                repeatHeaderHeight += table.RowHeights[headerIndex] + GetTableRowGapAfter(headerIndex, tbColumn.Rows.Count, columnTableRowGap);
                            }

                            bool HasRepeatableHeader() =>
                                table.RepeatHeaderRowCount > 0 &&
                                tbColumn.Rows.Count > table.HeaderRowCount;

                            bool AtContinuationPageTop() =>
                                Math.Abs(yCol - yStart) <= 0.001;

                            double MeasureColumnTableRowSegmentHeight(int rowIndex, int startLine, int lineCount, bool suppressCellObjects) {
                                double rowLeading = table.RowLeadings[rowIndex];
                                double rowPadTop = GetTableRowMaxPaddingTop(tbColumn, tableStyle, rowIndex, table.Columns);
                                double rowPadBottom = GetTableRowMaxPaddingBottom(tbColumn, tableStyle, rowIndex, table.Columns);
                                double segmentHeight = rowLeading + rowPadTop + rowPadBottom;
                                var cells = GetTableCellLayouts(tbColumn, rowIndex, table.Columns);
                                for (int cellIndex = 0; cellIndex < cells.Count; cellIndex++) {
                                    TableCellLayout cell = cells[cellIndex];
                                    double cellWidth = GetTableCellWidth(table.ColumnWidths, cell.Column, cell.ColumnSpan, columnGap);
                                    double cellPadLeft = GetTableCellPaddingLeft(tableStyle, rowIndex, cell.Column);
                                    double cellPadRight = GetTableCellPaddingRight(tableStyle, rowIndex, cell.Column);
                                    double innerW = cellWidth - cellPadLeft - cellPadRight;
                                    TableCellTextLayout lines = table.RowLines[rowIndex][cell.Column];
                                    int sourceStartLine = startLine;
                                    int visibleLineCount = Math.Max(0, Math.Min(lineCount, lines.LineCount - sourceStartLine));
                                    bool includeObjects = !suppressCellObjects && sourceStartLine == 0;
                                    double cellContentHeight = MeasureTableCellContentHeight(cell, lines, sourceStartLine, visibleLineCount, rowLeading, innerW, includeObjects) +
                                        GetTableCellPaddingTop(tableStyle, rowIndex, cell.Column) +
                                        GetTableCellPaddingBottom(tableStyle, rowIndex, cell.Column);
                                    segmentHeight = Math.Max(segmentHeight, cellContentHeight);
                                }

                                return segmentHeight;
                            }

                            int GetColumnTableRowSegmentLineCountThatFits(int rowIndex, int startLine, double available) {
                                int remainingLines = table.RowLineCounts[rowIndex] - startLine;
                                int best = 0;
                                for (int candidate = 1; candidate <= remainingLines; candidate++) {
                                    double candidateHeight = MeasureColumnTableRowSegmentHeight(rowIndex, startLine, candidate, suppressCellObjects: false);
                                    if (candidateHeight > available + 0.001) {
                                        break;
                                    }

                                    best = candidate;
                                }

                                return Math.Max(1, best);
                            }

                            bool CanSplitColumnTableRowIntoRemainingSpace(int rowIndex) =>
                                rowIndex >= table.HeaderRowCount &&
                                GetTableRowAllowBreakAcrossPages(tableStyle, rowIndex) &&
                                table.RowLineCounts[rowIndex] > 1 &&
                                MeasureColumnTableRowSegmentHeight(rowIndex, 0, Math.Min(2, table.RowLineCounts[rowIndex]), suppressCellObjects: false) <= remain + 0.001;

                            bool ShouldBreakBeforeFinalColumnTableBodyRows(int rowIndex) {
                                int minimumBodyRows = Math.Min(tableStyle.MinimumBodyRowsOnLastPage, Math.Max(0, table.FooterStartRowIndex - table.HeaderRowCount));
                                if (minimumBodyRows <= 0 || table.FooterStartRowIndex - rowIndex != minimumBodyRows) {
                                    return false;
                                }

                                double currentRowHeight = table.RowHeights[rowIndex] + GetTableRowGapAfter(rowIndex, tbColumn.Rows.Count, columnTableRowGap);
                                double finalGroupHeight = GetTableRowsHeight(table.RowHeights, rowIndex, table.RowHeights.Length, columnTableRowGap);
                                return ShouldBreakBeforeFinalTableBodyRows(
                                    rowIndex,
                                    table.HeaderRowCount,
                                    table.FooterStartRowIndex,
                                    minimumBodyRows,
                                    currentRowHeight,
                                    finalGroupHeight,
                                    remain,
                                    HasRepeatableHeader() ? repeatHeaderHeight : 0D,
                                    maxContentHeight,
                                    consumed > 0.001);
                            }

                            void DrawColumnTableRowSegment(int rowIndex, bool renderAsHeader, int startLine, int lineCount, bool suppressCellObjects = false) {
                                bool renderAsFooter = rowIndex >= table.FooterStartRowIndex;
                                bool rowUsesBold = table.RowBold[rowIndex];
                                double rowSize = table.RowSizes[rowIndex];
                                double rowLeading = table.RowLeadings[rowIndex];
                                bool wholeRowSegment = startLine == 0 && lineCount == table.RowLineCounts[rowIndex];
                                double rowPadTop = GetTableRowMaxPaddingTop(tbColumn, tableStyle, rowIndex, table.Columns);
                                double rowPadBottom = GetTableRowMaxPaddingBottom(tbColumn, tableStyle, rowIndex, table.Columns);
                                double rowHeight = wholeRowSegment ? table.RowHeights[rowIndex] : MeasureColumnTableRowSegmentHeight(rowIndex, startLine, lineCount, suppressCellObjects);
                                if (rowUsesBold) {
                                    currentPage!.UsedBold = true;
                                    usedBold = true;
                                }

                                var cells = GetTableCellLayouts(tbColumn, rowIndex, table.Columns);
                                double rowBottom = yCol - rowHeight;
                                int bodyRowIndex = rowIndex - table.HeaderRowCount;
                                bool stripeBodyRow = bodyRowIndex >= 0 && bodyRowIndex % 2 == 1;
                                bool[] rowFillSkips = GetRowSpanContinuationSkipColumns(tbColumn, rowIndex, table.Columns);
                                if (tableStyle.HeaderFill is not null && renderAsHeader) { pageDirty = true; DrawTableRowFill(sb, tableStyle.HeaderFill.Value, xTable, table.ColumnWidths, columnGap, rowBottom, rowHeight, rowFillSkips, emitGeneratedStructure); }
                                else if (tableStyle.FooterFill is not null && renderAsFooter) { pageDirty = true; DrawTableRowFill(sb, tableStyle.FooterFill.Value, xTable, table.ColumnWidths, columnGap, rowBottom, rowHeight, rowFillSkips, emitGeneratedStructure); }
                                else if (!renderAsHeader && !renderAsFooter && tableStyle.RowStripeFill is not null && stripeBodyRow) { pageDirty = true; DrawTableRowFill(sb, tableStyle.RowStripeFill.Value, xTable, table.ColumnWidths, columnGap, rowBottom, rowHeight, rowFillSkips, emitGeneratedStructure); }

                                if (!renderAsHeader && !renderAsFooter && tableStyle.BodyColumnFills != null) {
                                    bool[] bodyColumnFillSkips = GetMergedCellContinuationSkipColumns(tbColumn, rowIndex, table.Columns);
                                    double fillX = xTable;
                                    for (int fillColumn = 0; fillColumn < table.Columns; fillColumn++) {
                                        PdfColor? fill = fillColumn < tableStyle.BodyColumnFills.Count ? tableStyle.BodyColumnFills[fillColumn] : null;
                                        if (fill.HasValue && (fillColumn >= bodyColumnFillSkips.Length || !bodyColumnFillSkips[fillColumn])) {
                                            pageDirty = true;
                                            DrawRowFill(sb, fill.Value, fillX, rowBottom, table.ColumnWidths[fillColumn], rowHeight, emitGeneratedStructure);
                                        }
                                        fillX += table.ColumnWidths[fillColumn] + columnGap;
                                    }
                                }

                                if (tableStyle.CellFills != null && tableStyle.CellFills.Count > 0) {
                                    double fillX = xTable;
                                    for (int fillColumn = 0; fillColumn < table.Columns; fillColumn++) {
                                        if (tableStyle.CellFills.TryGetValue((rowIndex, fillColumn), out PdfColor fill) &&
                                            TryGetTableCellLayoutAtColumn(cells, fillColumn, out TableCellLayout fillCell) &&
                                            (fillColumn >= rowFillSkips.Length || !rowFillSkips[fillColumn])) {
                                            int span = wholeRowSegment ? fillCell.ColumnSpan : 1;
                                            double fillHeight = rowHeight;
                                            double fillBottom = rowBottom;
                                            if (wholeRowSegment) {
                                                if (fillCell.RowSpan > 1) {
                                                    fillHeight = GetTableCellHeight(table.RowHeights, rowIndex, fillCell.RowSpan, columnTableRowGap);
                                                    fillBottom = yCol - fillHeight;
                                                }
                                            }

                                            pageDirty = true;
                                            DrawRowFill(sb, fill, fillX, fillBottom, GetTableCellWidth(table.ColumnWidths, fillColumn, span, columnGap), fillHeight, emitGeneratedStructure);
                                        }
                                        fillX += table.ColumnWidths[fillColumn] + columnGap;
                                    }
                                }
                                if (DrawTableCellDataBars(sb, tableStyle, cells, rowIndex, table.Columns, xTable, yCol, rowBottom, rowHeight, table.ColumnWidths, columnGap, table.RowHeights, columnTableRowGap, wholeRowSegment, startLine, rowFillSkips, emitGeneratedStructure)) {
                                    pageDirty = true;
                                }
                                if (DrawTableCellIcons(sb, tableStyle, cells, rowIndex, table.Columns, xTable, yCol, rowBottom, rowHeight, table.ColumnWidths, columnGap, table.RowHeights, columnTableRowGap, wholeRowSegment, startLine, rowFillSkips, emitGeneratedStructure)) {
                                    pageDirty = true;
                                }

                                var textColor = renderAsHeader ? tableStyle.HeaderTextColor : renderAsFooter ? tableStyle.FooterTextColor : tableStyle.TextColor;
                                double xi = xTable;
                                int? rowStructureElementIndex = RegisterStructureContainer("TR", EnsureTableStructureElement());
                                for (int cellIndex = 0; cellIndex < cells.Count; cellIndex++) {
                                    TableCellLayout cell = cells[cellIndex];
                                    int c = cell.Column;
                                    xi = xTable;
                                    for (int xColumn = 0; xColumn < c; xColumn++) {
                                        xi += table.ColumnWidths[xColumn] + columnGap;
                                    }

                                    double cellWidth = GetTableCellWidth(table.ColumnWidths, c, cell.ColumnSpan, columnGap);
                                    double cellPadLeft = GetTableCellPaddingLeft(tableStyle, rowIndex, c);
                                    double cellPadRight = GetTableCellPaddingRight(tableStyle, rowIndex, c);
                                    double cellPadTop = GetTableCellPaddingTop(tableStyle, rowIndex, c);
                                    double cellPadBottom = GetTableCellPaddingBottom(tableStyle, rowIndex, c);
                                    double innerW = cellWidth - cellPadLeft - cellPadRight;
                                    double cellHeight = wholeRowSegment && cell.RowSpan > 1 ? GetTableCellHeight(table.RowHeights, rowIndex, cell.RowSpan, columnTableRowGap) : rowHeight;
                                    double cellBottom = yCol - cellHeight;
                                    PdfColumnAlign align = GetTableCellAlignment(tableStyle, rowIndex, c, cell.Text);
                                    PdfCellVerticalAlign verticalAlign = GetTableCellVerticalAlignment(tableStyle, rowIndex, c);
                                    var cellFont = GetTableRowFont(currentOpts, rowUsesBold);
                                    TableCellTextLayout lines = table.RowLines[rowIndex][c];
                                    int sourceStartLine = wholeRowSegment && cell.RowSpan > 1 ? 0 : startLine;
                                    int requestedLineCount = wholeRowSegment && cell.RowSpan > 1 ? lines.LineCount : lineCount;
                                    int visibleLineCount = Math.Max(0, Math.Min(requestedLineCount, lines.LineCount - sourceStartLine));
                                    double verticalOffset = 0;
                                    double visibleTextHeight = 0D;
                                    if (visibleLineCount > 0) {
                                        double availableTextHeight = Math.Max(0, cellHeight - cellPadTop - cellPadBottom);
                                        visibleTextHeight = MeasureTableCellTextHeight(lines, sourceStartLine, visibleLineCount, rowLeading);
                                        double visibleContentHeight = MeasureTableCellContentHeight(cell, lines, sourceStartLine, visibleLineCount, rowLeading, innerW);
                                        double unusedTextHeight = Math.Max(0, availableTextHeight - visibleContentHeight);
                                        if (verticalAlign == PdfCellVerticalAlign.Middle) verticalOffset = unusedTextHeight / 2;
                                        else if (verticalAlign == PdfCellVerticalAlign.Bottom) verticalOffset = unusedTextHeight;
                                    }

                                    double firstBaseline = yCol - cellPadTop - verticalOffset - GetAscenderForOptions(cellFont, rowSize, currentOpts) + tableStyle.RowBaselineOffset;

                                    pageDirty = true;
                                    if (cell.Runs.Any(run => run.Bold || rowUsesBold)) { currentPage!.UsedBold = true; usedBold = true; }
                                    if (cell.Runs.Any(run => run.Italic)) { currentPage!.UsedItalic = true; usedItalic = true; }
                                    if (cell.Runs.Any(run => (run.Bold || rowUsesBold) && run.Italic)) { currentPage!.UsedBoldItalic = true; usedBoldItalic = true; }
                                    MarkRichFonts(cell.Runs);
                                    string? linkUri = cell.LinkUri;
                                    string? linkDestinationName = cell.LinkDestinationName;
                                    string? linkContents = cell.LinkContents;
                                    if (tbColumn.Links.TryGetValue((rowIndex, c), out var uri)) {
                                        linkUri = uri;
                                        linkDestinationName = null;
                                        linkContents = cell.Text;
                                    }

                                    if (sourceStartLine == 0) {
                                        AddTableCellNamedDestinationName(cell.NamedDestinationName, yCol);
                                    }

                                    int? cellLinkStructElementIndex = null;
                                    if (visibleLineCount > 0) {
                                        var visibleLines = SliceTableCellLines(lines, sourceStartLine, visibleLineCount);
                                        visibleLines = StripRichLineLinksWhenCellLinked(visibleLines, linkUri, linkDestinationName);
                                        var visibleHeights = SliceTableCellLineHeights(lines, sourceStartLine, visibleLineCount, rowLeading);
                                        var visibleAlignments = SliceTableCellLineAlignments(lines, sourceStartLine, visibleLineCount);
                                        var visibleXOffsets = SliceTableCellLineXOffsets(lines, sourceStartLine, visibleLineCount);
                                        var visibleWidths = SliceTableCellLineWidths(lines, sourceStartLine, visibleLineCount, innerW);
                                        var paragraph = new RichParagraphBlock(StripRunLinksWhenCellLinked(cell.Runs, linkUri, linkDestinationName), MapTableCellAlignment(align), textColor);
                                        string structureType = renderAsHeader ? "TH" : "TD";
                                        int tableColumnSpan = cell.ColumnSpan > 1 ? cell.ColumnSpan : 1;
                                        int tableRowSpan = wholeRowSegment && cell.RowSpan > 1 ? cell.RowSpan : 1;
                                        bool cellHasLinkTarget = HasCellLinkTarget(linkUri, linkDestinationName);
                                        int? markedContentId;
                                        string markedStructureType = structureType;
                                        if (cellHasLinkTarget && emitGeneratedStructure && currentPage != null) {
                                            int? cellElementIndex = RegisterStructureContainer(structureType, rowStructureElementIndex, renderAsHeader ? "Column" : string.Empty, tableColumnSpan, tableRowSpan);
                                            markedStructureType = "Link";
                                            markedContentId = RegisterTextStructureElement(markedStructureType, cellElementIndex);
                                            cellLinkStructElementIndex = FindStructElementIndex(currentPage, markedContentId, markedStructureType);
                                        } else {
                                            markedContentId = RegisterTextStructureElement(structureType, rowStructureElementIndex, renderAsHeader ? "Column" : string.Empty, tableColumnSpan, tableRowSpan);
                                        }

                                        WriteClippedRichParagraph(sb, paragraph, visibleLines, visibleHeights, currentOpts, firstBaseline, rowSize, rowLeading, currentPage!.Annotations, xi - TableCellClipBleed, cellBottom - TableCellClipBleed, cellWidth + (TableCellClipBleed * 2D), cellHeight + (TableCellClipBleed * 2D), xi + cellPadLeft, innerW, structureType: markedStructureType, markedContentId: markedContentId, structurePage: currentPage, lineAlignments: visibleAlignments, lineXOffsets: visibleXOffsets, lineWidths: visibleWidths);
                                    }
                                    if (!suppressCellObjects && (cell.Images.Count > 0 || cell.CheckBoxes.Count > 0 || cell.FormFields.Count > 0) && sourceStartLine == 0) {
                                        if (CanRenderTableCellCheckBoxInline(cell, lines, sourceStartLine, visibleLineCount)) {
                                            RenderTableCellInlineCheckBox(currentPage!, cell, align, lines.Lines[sourceStartLine], xi + cellPadLeft, innerW, firstBaseline);
                                        } else {
                                            double formFieldTop = yCol - cellPadTop - verticalOffset - (string.IsNullOrEmpty(cell.Text) ? 0D : visibleTextHeight + TableCellCheckBoxGap);
                                            RenderTableCellObjects(currentPage!, cell, align, xi + cellPadLeft, innerW, formFieldTop);
                                        }
                                    }

                                    if (HasCellLinkTarget(linkUri, linkDestinationName)) {
                                        double linkCellHeight = sourceStartLine == 0 && cell.RowSpan > 1
                                            ? GetTableCellHeight(table.RowHeights, rowIndex, cell.RowSpan, columnTableRowGap)
                                            : cellHeight;
                                        currentPage!.Annotations.Add(new LinkAnnotation { X1 = xi + cellPadLeft - TableCellClipBleed, Y1 = yCol - linkCellHeight - TableCellClipBleed, X2 = xi + cellWidth - cellPadRight + TableCellClipBleed, Y2 = yCol + TableCellClipBleed, Uri = linkUri, DestinationName = linkDestinationName, Contents = linkContents ?? cell.Text, StructElementIndex = cellLinkStructElementIndex });
                                    }
                                }

                                if (tableStyle.BorderColor is not null && tableStyle.BorderWidth > 0) {
                                    pageDirty = true;
                                    bool[] topBorderSkips = GetRowSpanBoundarySkipColumns(tbColumn, rowIndex - 1, table.Columns);
                                    bool[] bottomBorderSkips = GetRowSpanBoundarySkipColumns(tbColumn, rowIndex, table.Columns);
                                    bool segmentBorderRows = HasSkippedColumns(topBorderSkips, table.Columns) || HasSkippedColumns(bottomBorderSkips, table.Columns);
                                    if (segmentBorderRows) {
                                        DrawTableHorizontalLine(sb, tableStyle.BorderColor.Value, tableStyle.BorderWidth, xTable, table.ColumnWidths, columnGap, rowBottom + rowHeight, topBorderSkips, emitGeneratedStructure);
                                        DrawTableHorizontalLine(sb, tableStyle.BorderColor.Value, tableStyle.BorderWidth, xTable, table.ColumnWidths, columnGap, rowBottom, bottomBorderSkips, emitGeneratedStructure);
                                        DrawVLine(sb, tableStyle.BorderColor.Value, tableStyle.BorderWidth, xTable, rowBottom + rowHeight, rowBottom, emitGeneratedStructure);
                                        DrawVLine(sb, tableStyle.BorderColor.Value, tableStyle.BorderWidth, xTable + table.Width, rowBottom + rowHeight, rowBottom, emitGeneratedStructure);
                                    } else {
                                        DrawRowRect(sb, tableStyle.BorderColor.Value, tableStyle.BorderWidth, xTable, rowBottom, table.Width, rowHeight, emitGeneratedStructure);
                                    }

                                    double xi2 = xTable;
                                    for (int c = 0; c < table.Columns - 1; c++) {
                                        xi2 += table.ColumnWidths[c];
                                        if (IsTableBoundaryInsideSpannedCell(tbColumn, rowIndex, c, table.Columns)) {
                                            xi2 += columnGap;
                                            continue;
                                        }

                                        DrawVLine(sb, tableStyle.BorderColor.Value, tableStyle.BorderWidth, xi2, rowBottom + rowHeight, rowBottom, emitGeneratedStructure);
                                        xi2 += columnGap;
                                    }
                                }

                                if (renderAsFooter && rowIndex == table.FooterStartRowIndex) {
                                    PdfColor? footerSeparatorColor = tableStyle.FooterSeparatorColor ?? tableStyle.RowSeparatorColor;
                                    double footerSeparatorWidth = tableStyle.FooterSeparatorWidth > 0 ? tableStyle.FooterSeparatorWidth : tableStyle.RowSeparatorWidth;
                                    if (footerSeparatorColor is not null && footerSeparatorWidth > 0) {
                                        pageDirty = true;
                                        DrawTableHorizontalLine(sb, footerSeparatorColor.Value, footerSeparatorWidth, xTable, table.ColumnWidths, columnGap, yCol, GetRowSpanBoundarySkipColumns(tbColumn, rowIndex - 1, table.Columns), emitGeneratedStructure);
                                    }
                                }

                                PdfColor? separatorColor = renderAsHeader && tableStyle.HeaderSeparatorColor is not null ? tableStyle.HeaderSeparatorColor : tableStyle.RowSeparatorColor;
                                double separatorWidth = renderAsHeader && tableStyle.HeaderSeparatorWidth > 0 ? tableStyle.HeaderSeparatorWidth : tableStyle.RowSeparatorWidth;
                                if (separatorColor is not null && separatorWidth > 0) {
                                    pageDirty = true;
                                    DrawTableHorizontalLine(sb, separatorColor.Value, separatorWidth, xTable, table.ColumnWidths, columnGap, rowBottom, GetRowSpanBoundarySkipColumns(tbColumn, rowIndex, table.Columns), emitGeneratedStructure);
                                }

                                if (tableStyle.CellBorders != null && tableStyle.CellBorders.Count > 0) {
                                    double borderX = xTable;
                                    for (int borderColumn = 0; borderColumn < table.Columns; borderColumn++) {
                                        if (tableStyle.CellBorders.TryGetValue((rowIndex, borderColumn), out PdfCellBorder? cellBorder) &&
                                            TryGetTableCellLayoutAtColumn(cells, borderColumn, out TableCellLayout borderCell) &&
                                            (borderColumn >= rowFillSkips.Length || !rowFillSkips[borderColumn]) &&
                                            HasRenderableCellBorder(cellBorder)) {
                                            int span = wholeRowSegment ? borderCell.ColumnSpan : 1;
                                            double borderHeight = rowHeight;
                                            double borderBottom = rowBottom;
                                            if (wholeRowSegment) {
                                                if (borderCell.RowSpan > 1) {
                                                    borderHeight = GetTableCellHeight(table.RowHeights, rowIndex, borderCell.RowSpan, columnTableRowGap);
                                                    borderBottom = yCol - borderHeight;
                                                }
                                            }

                                            pageDirty = true;
                                            DrawCellBorder(sb, cellBorder, borderX, borderBottom, GetTableCellWidth(table.ColumnWidths, borderColumn, span, columnGap), borderHeight, emitGeneratedStructure);
                                        }
                                        borderX += table.ColumnWidths[borderColumn] + columnGap;
                                    }
                                }

                                double rowAdvance = rowHeight + (wholeRowSegment ? GetTableRowGapAfter(rowIndex, tbColumn.Rows.Count, columnTableRowGap) : 0D);
                                yCol -= rowAdvance;
                                remain -= rowAdvance;
                                consumed += rowAdvance;
                            }

                            void DrawColumnTableRow(int rowIndex, bool renderAsHeader, bool suppressCellObjects = false) =>
                                DrawColumnTableRowSegment(rowIndex, renderAsHeader, 0, table.RowLineCounts[rowIndex], suppressCellObjects);

                            int rowIndex = line;
                            int rowStartLine = subline;
                            while (rowIndex < tbColumn.Rows.Count) {
                                double rowHeight = table.RowHeights[rowIndex];
                                if (rowHeight > maxContentHeight + 0.001) {
                                    if (!GetTableRowAllowBreakAcrossPages(tableStyle, rowIndex)) {
                                        throw new ArgumentException("Table row height exceeds the available page content height and row splitting is disabled.");
                                    }

                                    int totalLines = table.RowLineCounts[rowIndex];
                                    double rowPadTop = GetTableRowMaxPaddingTop(tbColumn, tableStyle, rowIndex, table.Columns);
                                    double rowPadBottom = GetTableRowMaxPaddingBottom(tbColumn, tableStyle, rowIndex, table.Columns);
                                    bool repeatHeaderBeforeSegment = rowIndex >= table.HeaderRowCount &&
                                        HasRepeatableHeader() &&
                                        AtContinuationPageTop() &&
                                        repeatHeaderHeight + table.RowLeadings[rowIndex] + rowPadTop + rowPadBottom <= remain + 0.001;
                                    double neededForFirstSegment = table.RowLeadings[rowIndex] + rowPadTop + rowPadBottom + (repeatHeaderBeforeSegment ? repeatHeaderHeight : 0);
                                    if (neededForFirstSegment > remain && consumed > 0) break;
                                    if (neededForFirstSegment > remain && consumed == 0) { remain = 0; break; }

                                    if (repeatHeaderBeforeSegment) {
                                        for (int headerIndex = 0; headerIndex < table.RepeatHeaderRowCount; headerIndex++) {
                                            DrawColumnTableRow(headerIndex, renderAsHeader: true, suppressCellObjects: true);
                                        }
                                    }

                                    int take = Math.Min(totalLines - rowStartLine, GetColumnTableRowSegmentLineCountThatFits(rowIndex, rowStartLine, remain));
                                    DrawColumnTableRowSegment(rowIndex, renderAsHeader: rowIndex < table.HeaderRowCount && rowStartLine == 0, rowStartLine, take);
                                    rowStartLine += take;

                                    if (rowStartLine < totalLines) {
                                        line = rowIndex;
                                        subline = rowStartLine;
                                        break;
                                    }

                                    double gapAfterSplitRow = GetTableRowGapAfter(rowIndex, tbColumn.Rows.Count, columnTableRowGap);
                                    if (gapAfterSplitRow > 0) {
                                        yCol -= gapAfterSplitRow;
                                        remain -= gapAfterSplitRow;
                                        consumed += gapAfterSplitRow;
                                    }

                                    rowIndex++;
                                    line = rowIndex;
                                    subline = 0;
                                    rowStartLine = 0;
                                    continue;
                                }
                                bool repeatHeaderBeforeRow = rowIndex >= table.HeaderRowCount &&
                                    HasRepeatableHeader() &&
                                    AtContinuationPageTop() &&
                                    repeatHeaderHeight + rowHeight <= remain + 0.001;
                                double neededForNextRow = rowHeight + GetTableRowGapAfter(rowIndex, tbColumn.Rows.Count, columnTableRowGap) + (repeatHeaderBeforeRow ? repeatHeaderHeight : 0);
                                if (rowHeight > remain + 0.001 && consumed > 0 && CanSplitColumnTableRowIntoRemainingSpace(rowIndex)) {
                                    int take = Math.Min(table.RowLineCounts[rowIndex], GetColumnTableRowSegmentLineCountThatFits(rowIndex, 0, remain));
                                    DrawColumnTableRowSegment(rowIndex, renderAsHeader: false, 0, take);
                                    line = rowIndex;
                                    subline = take;
                                    break;
                                }

                                if (ShouldBreakBeforeFinalColumnTableBodyRows(rowIndex)) break;
                                if (neededForNextRow > remain && consumed > 0) break;
                                if (neededForNextRow > remain && consumed == 0) { remain = 0; break; }

                                if (repeatHeaderBeforeRow) {
                                    for (int headerIndex = 0; headerIndex < table.RepeatHeaderRowCount; headerIndex++) {
                                        DrawColumnTableRow(headerIndex, renderAsHeader: true, suppressCellObjects: true);
                                    }
                                }

                                DrawColumnTableRow(rowIndex, renderAsHeader: rowIndex < table.HeaderRowCount);
                                rowIndex++;
                                line = rowIndex;
                                subline = 0;
                                rowStartLine = 0;
                            }

                            if (rowIndex >= tbColumn.Rows.Count) {
                                if (tableStyle.SpacingAfter > 0 && tableStyle.SpacingAfter <= remain) {
                                    yCol -= tableStyle.SpacingAfter;
                                    remain -= tableStyle.SpacingAfter;
                                    consumed += tableStyle.SpacingAfter;
                                }
                                idx++;
                                line = 0;
                                subline = 0;
                            } else {
                                break;
                            }
                        } else if (it is ColRule cr) {
                            PdfHorizontalRuleStyle hr2 = ResolveHorizontalRuleStyle(cr.Block, currentOpts);
                            ValidateHorizontalRule(hr2);
                            double spacingBefore = ResolveColumnSpacingBefore(hr2.SpacingBefore, consumed);
                            double needed = spacingBefore + hr2.Thickness + hr2.SpacingAfter;
                            EnsureFixedFlowBlockFits("Horizontal rule", wCol, needed, wCol);
                            if (line == 0 && hr2.KeepWithNext && idx + 1 < items.Count) {
                                double nextHeight = MeasureColKeepWithNextChainHeight(items, idx + 1);
                                double keepHeight = needed + nextHeight;
                                double availableHeight = currentOpts.PageHeight - currentOpts.MarginTop - currentOpts.MarginBottom;
                                if (nextHeight > 0.001 && keepHeight <= availableHeight + 0.001 && keepHeight > remain + 0.001) {
                                    if (consumed > 0) break;
                                    remain = 0;
                                    break;
                                }
                            }

                            if (needed > remain && consumed > 0) break;
                            if (needed > remain && consumed == 0) { remain = 0; break; }
                            if (spacingBefore > 0) yCol -= spacingBefore;
                            double x1 = xCol, x2 = xCol + wCol, yLine = yCol - hr2.Thickness * 0.5;
                            pageDirty = true;
                            DrawHLine(sb, hr2.Color, hr2.Thickness, x1, x2, yLine, emitGeneratedStructure);
                            yCol -= hr2.Thickness + hr2.SpacingAfter; remain -= needed; consumed += needed; idx++;
                        } else if (it is ColImg ciimg) {
                            var ib2 = ciimg.Block;
                            PdfImageStyle imageStyle = ciimg.Style;
                            PdfDocument.ValidateImageStyleForBox(imageStyle, ib2.Width, ib2.Height, nameof(imageStyle.ClipPath));
                            PdfDocument.ValidateImageFitDimensions(ib2.Info, imageStyle.Fit, nameof(imageStyle.Fit));
                            double spacingBefore = ResolveColumnSpacingBefore(imageStyle.SpacingBefore, consumed);
                            double needed = spacingBefore + ciimg.Height + imageStyle.SpacingAfter;
                            EnsureFixedFlowBlockFits("Image", ciimg.Width, needed, wCol);
                            if (imageStyle.KeepWithNext && idx + 1 < items.Count) {
                                double nextHeight = MeasureColKeepWithNextChainHeight(items, idx + 1);
                                double keepHeight = needed + nextHeight;
                                double availableHeight = currentOpts.PageHeight - currentOpts.MarginTop - currentOpts.MarginBottom;
                                if (nextHeight > 0.001 && keepHeight <= availableHeight + 0.001 && keepHeight > remain + 0.001) {
                                    if (consumed > 0) break;
                                    remain = 0;
                                    break;
                                }
                            }

                            if (needed > remain && consumed > 0) break;
                            if (needed > remain && consumed == 0) { remain = 0; break; }
                            if (spacingBefore > 0) yCol -= spacingBefore;
                            double xImg = xCol;
                            if (imageStyle.Align == PdfAlign.Center) xImg = xCol + Math.Max(0, (wCol - ciimg.Width) / 2);
                            else if (imageStyle.Align == PdfAlign.Right) xImg = xCol + Math.Max(0, wCol - ciimg.Width);
                            PageImage pageImage = CreatePageImage(ib2, imageStyle, xImg, yCol - ciimg.Height, ciimg.Width, ciimg.Height);
                            currentPage!.Images.Add(pageImage);
                            AddImageLinkAnnotation(ib2, imageStyle, pageImage, xImg, yCol - ciimg.Height, ciimg.Width, ciimg.Height);
                            pageDirty = true;
                            yCol -= ciimg.Height + imageStyle.SpacingAfter; remain -= needed; consumed += needed; idx++;
                        } else if (it is ColShape cs) {
                            var shape = cs.Block;
                            PdfDrawingStyle shapeStyle = ResolveDrawingStyle(shape, currentOpts);
                            PdfDocument.ValidateDrawingStyle(shapeStyle, "Shape");
                            double spacingBefore = ResolveColumnSpacingBefore(shapeStyle.SpacingBefore, consumed);
                            double needed = spacingBefore + shape.Shape.Height + shapeStyle.SpacingAfter;
                            EnsureFixedFlowBlockFits("Shape", shape.Shape.Width, needed, wCol);
                            if (shapeStyle.KeepWithNext && idx + 1 < items.Count) {
                                double nextHeight = MeasureColKeepWithNextChainHeight(items, idx + 1);
                                double keepHeight = needed + nextHeight;
                                double availableHeight = currentOpts.PageHeight - currentOpts.MarginTop - currentOpts.MarginBottom;
                                if (nextHeight > 0.001 && keepHeight <= availableHeight + 0.001 && keepHeight > remain + 0.001) {
                                    if (consumed > 0) break;
                                    remain = 0;
                                    break;
                                }
                            }

                            if (needed > remain && consumed > 0) break;
                            if (needed > remain && consumed == 0) { remain = 0; break; }
                            if (spacingBefore > 0) yCol -= spacingBefore;
                            int? structElementIndex = DrawShapeAt(shape, shapeStyle, xCol, wCol, yCol);
                            AddShapeLinkAnnotation(shape, shapeStyle, xCol, wCol, yCol, structElementIndex);
                            yCol -= shape.Shape.Height + shapeStyle.SpacingAfter;
                            remain -= needed;
                            consumed += needed;
                            idx++;
                        } else if (it is ColDrawing cd) {
                            var drawing = cd.Block;
                            PdfDrawingStyle drawingStyle = ResolveDrawingStyle(drawing, currentOpts);
                            PdfDocument.ValidateDrawingStyle(drawingStyle, "Drawing");
                            double spacingBefore = ResolveColumnSpacingBefore(drawingStyle.SpacingBefore, consumed);
                            double needed = spacingBefore + drawing.Drawing.Height + drawingStyle.SpacingAfter;
                            EnsureFixedFlowBlockFits("Drawing", drawing.Drawing.Width, needed, wCol);
                            if (drawingStyle.KeepWithNext && idx + 1 < items.Count) {
                                double nextHeight = MeasureColKeepWithNextChainHeight(items, idx + 1);
                                double keepHeight = needed + nextHeight;
                                double availableHeight = currentOpts.PageHeight - currentOpts.MarginTop - currentOpts.MarginBottom;
                                if (nextHeight > 0.001 && keepHeight <= availableHeight + 0.001 && keepHeight > remain + 0.001) {
                                    if (consumed > 0) break;
                                    remain = 0;
                                    break;
                                }
                            }

                            if (needed > remain && consumed > 0) break;
                            if (needed > remain && consumed == 0) { remain = 0; break; }
                            if (spacingBefore > 0) yCol -= spacingBefore;
                            int? structElementIndex = DrawDrawingAt(drawing, drawingStyle, xCol, wCol, yCol);
                            AddDrawingLinkAnnotation(drawing, drawingStyle, xCol, wCol, yCol, structElementIndex);
                            yCol -= drawing.Drawing.Height + drawingStyle.SpacingAfter;
                            remain -= needed;
                            consumed += needed;
                            idx++;
                        } else if (it is ColForm form) {
                            double spacingBefore = ResolveColumnSpacingBefore(GetFormFieldSpacingBefore(form.Block), consumed);
                            double fieldWidth = GetFormFieldWidth(form.Block);
                            double fieldHeight = GetFormFieldHeight(form.Block);
                            double spacingAfter = GetFormFieldSpacingAfter(form.Block);
                            double needed = spacingBefore + fieldHeight + spacingAfter;
                            EnsureFixedFlowBlockFits(GetFormFieldBlockName(form.Block), fieldWidth, needed, wCol);
                            if (needed > remain && consumed > 0) break;
                            if (needed > remain && consumed == 0) { remain = 0; break; }
                            if (spacingBefore > 0) yCol -= spacingBefore;
                            double xField = GetAlignedObjectX(xCol, wCol, fieldWidth, GetFormFieldAlign(form.Block));
                            AddFormFieldAnnotation(form.Block, xField, yCol);
                            pageDirty = true;
                            yCol -= fieldHeight + spacingAfter;
                            remain -= needed;
                            consumed += needed;
                            idx++;
                        } else if (it is ColBookmark bookmarkItem) {
                            AddNamedDestination(bookmarkItem.Block, yCol);
                            idx++;
                        } else if (it is ColSpacer spacerItem) {
                            double needed = spacerItem.Block.Height;
                            double availableHeight = currentOpts.PageHeight - currentOpts.MarginTop - currentOpts.MarginBottom;
                            if (needed > availableHeight + 0.001) {
                                throw new ArgumentException("Spacer height exceeds the available page content height.");
                            }

                            if (needed > remain && consumed > 0) break;
                            if (needed > remain && consumed == 0) { remain = 0; break; }
                            yCol -= needed;
                            remain -= needed;
                            consumed += needed;
                            idx++;
                        }
                    }
                    colStates[ci] = (idx, line, subline);
                    if (colStates[ci] != startState) {
                        anyColumnAdvanced = true;
                    }

                    if (consumed > maxConsumed) maxConsumed = consumed;
                }

                if (maxConsumed <= 0.01) {
                    if (anyColumnAdvanced && !AnyRemaining()) {
                        break;
                    }

                    if (Math.Abs(y - yStart) <= 0.001) {
                        throw new InvalidOperationException("Row column layout could not make progress on an empty page.");
                    }

                    NewPage();
                    continue;
                }
                DrawRowColumnSeparators(y, y - maxConsumed);
                fragmentDecorator?.Invoke(
                    sb,
                    fragmentInsertionIndex,
                    y,
                    y - maxConsumed,
                    isFirstFragment,
                    !AnyRemaining());
                y -= maxConsumed;
                isFirstFragment = false;
            }

            if (rowSpacingAfter > 0) {
                y -= rowSpacingAfter;
            }
        }

    }
}
