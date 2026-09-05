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
            double[] colWs = ResolveRowColumnWidths(rb, columnAreaWidth);
            double xAcc = currentOpts.MarginLeft;
            for (int i = 0; i < ncols; i++) { colXs[i] = xAcc; xAcc += colWs[i] + rowGap; }

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
            var columnGroups = Enumerable.Range(0, ncols).Select(_ => new List<ColumnGroup>()).ToArray();
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
                double nextHeight = MeasureKeepWithNextChainHeight(blockList, blockIndex + 1, currentOpts.MarginLeft, width, currentOpts.DefaultFontSize, rowHeight);
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
                    List<ColumnGroup> activeGroups = columnGroups[ci];
                    ResumeColumnGroups(activeGroups, colXs[ci], ref yCol, ref remain, ref consumed);
                    while (idx < items.Count && (remain > 0.1 || items[idx] is ColGroupEnd)) {
                        var it = items[idx];
                        xCol = colXs[ci] + it.ColumnXOffset;
                        wCol = it.ColumnWidth;
                        if (it is ColGroupStart groupStart) {
                            if (!TryBeginColumnGroup(groupStart.Group, activeGroups, items, idx, colXs[ci], ref yCol, ref remain, ref consumed)) break;
                            idx++;
                            continue;
                        }
                        if (it is ColGroupEnd groupEnd) {
                            EndColumnGroupFragment(groupEnd.Group, colXs[ci], ref yCol, ref remain, ref consumed);
                            activeGroups.RemoveAt(activeGroups.Count - 1);
                            double after = Math.Min(groupEnd.Group.Style?.SpacingAfter ?? 0D, Math.Max(0D, remain));
                            ConsumeColumnSpace(after, ref yCol, ref remain, ref consumed);
                            idx++;
                            continue;
                        }
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
                                double availableHeight = GetFullPageContentHeight() - activeGroups.Sum(group => (group.Style?.PaddingY ?? 0D) * 2D);
                                if (nextHeight > 0.001 && keepHeight <= availableHeight + 0.001 && keepHeight > remain + 0.001) {
                                    if (consumed > 0) break;
                                    remain = 0;
                                    break;
                                }
                            }

                            if (paragraphStyle?.KeepTogether == true && line == 0) {
                                double paragraphHeight = spacingBefore + heights.Sum() + spacingAfter;
                                double availableHeight = GetFullPageContentHeight() - activeGroups.Sum(group => (group.Style?.PaddingY ?? 0D) * 2D);
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
                                double availableHeight = GetFullPageContentHeight() - activeGroups.Sum(group => (group.Style?.PaddingY ?? 0D) * 2D);
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
                                double availableHeight = GetFullPageContentHeight() - activeGroups.Sum(group => (group.Style?.PaddingY ?? 0D) * 2D);
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
                                    double availableHeight = GetFullPageContentHeight() - activeGroups.Sum(group => (group.Style?.PaddingY ?? 0D) * 2D);
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
                        } else if (it is ColTable table) {
                            var state = new ColumnTableCursor { Index = idx, Line = line, Subline = subline, Y = yCol, Remaining = remain, Consumed = consumed };
                            bool completed = RenderColumnTable(table, items, state, xCol, wCol,
                                GetFullPageContentHeight() - activeGroups.Sum(group => (group.Style?.PaddingY ?? 0D) * 2D),
                                GetCurrentFramePageStartY() - activeGroups.Sum(group => group.Style?.PaddingY ?? 0D));
                            (idx, line, subline) = (state.Index, state.Line, state.Subline);
                            (yCol, remain, consumed) = (state.Y, state.Remaining, state.Consumed);
                            if (!completed) break;
                        } else if (it is ColRule cr) {
                            PdfHorizontalRuleStyle hr2 = ResolveHorizontalRuleStyle(cr.Block, currentOpts);
                            ValidateHorizontalRule(hr2);
                            double spacingBefore = ResolveColumnSpacingBefore(hr2.SpacingBefore, consumed);
                            double needed = spacingBefore + hr2.Thickness + hr2.SpacingAfter;
                            EnsureFixedFlowBlockFits("Horizontal rule", wCol, needed, wCol);
                            if (line == 0 && hr2.KeepWithNext && idx + 1 < items.Count) {
                                double nextHeight = MeasureColKeepWithNextChainHeight(items, idx + 1);
                                double keepHeight = needed + nextHeight;
                                double availableHeight = GetFullPageContentHeight() - activeGroups.Sum(group => (group.Style?.PaddingY ?? 0D) * 2D);
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
                                double availableHeight = GetFullPageContentHeight() - activeGroups.Sum(group => (group.Style?.PaddingY ?? 0D) * 2D);
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
                                double availableHeight = GetFullPageContentHeight() - activeGroups.Sum(group => (group.Style?.PaddingY ?? 0D) * 2D);
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
                                double availableHeight = GetFullPageContentHeight() - activeGroups.Sum(group => (group.Style?.PaddingY ?? 0D) * 2D);
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
                        } else if (it is ColAnnotation annotation) {
                            double spacingBefore = ResolveColumnSpacingBefore(GetAnnotationSpacingBefore(annotation.Block), consumed);
                            double annotationWidth = GetAnnotationWidth(annotation.Block);
                            double annotationHeight = GetAnnotationHeight(annotation.Block);
                            double spacingAfter = GetAnnotationSpacingAfter(annotation.Block);
                            double needed = spacingBefore + annotationHeight + spacingAfter;
                            EnsureFixedFlowBlockFits("Annotation", annotationWidth, needed, wCol);
                            if (needed > remain && consumed > 0) break;
                            if (needed > remain && consumed == 0) { remain = 0; break; }
                            if (spacingBefore > 0) yCol -= spacingBefore;
                            double xAnnotation = GetAlignedObjectX(xCol, wCol, annotationWidth, GetAnnotationAlign(annotation.Block));
                            double bottomY = yCol - annotationHeight;
                            if (annotation.Block is TextAnnotationBlock textAnnotation) {
                                AddTextAnnotation(xAnnotation, bottomY, annotationWidth, annotationHeight, textAnnotation.Contents, textAnnotation.Icon, textAnnotation.Color, textAnnotation.Open);
                            } else if (annotation.Block is FreeTextAnnotationBlock freeTextAnnotation) {
                                AddFreeTextAnnotation(xAnnotation, bottomY, annotationWidth, annotationHeight, freeTextAnnotation.Contents, freeTextAnnotation.FontSize, freeTextAnnotation.TextColor, freeTextAnnotation.BorderColor, freeTextAnnotation.BorderWidth, freeTextAnnotation.FillColor, freeTextAnnotation.TextAlign, freeTextAnnotation.Padding, freeTextAnnotation.LineHeight);
                            } else if (annotation.Block is HighlightAnnotationBlock highlightAnnotation) {
                                AddHighlightAnnotation(xAnnotation, bottomY, annotationWidth, annotationHeight, highlightAnnotation.Contents, highlightAnnotation.Color);
                            }

                            DrawDebugFlowObjectBox(xAnnotation, bottomY, annotationWidth, annotationHeight);
                            pageDirty = true;
                            yCol -= annotationHeight + spacingAfter;
                            remain -= needed;
                            consumed += needed;
                            idx++;
                        } else if (it is ColBookmark bookmarkItem) {
                            AddNamedDestination(bookmarkItem.Block, yCol);
                            idx++;
                        } else if (it is ColSpacer spacerItem) {
                            double needed = spacerItem.Block.Height;
                            double availableHeight = GetFullPageContentHeight() - activeGroups.Sum(group => (group.Style?.PaddingY ?? 0D) * 2D);
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
                    FinishColumnGroupsFragment(activeGroups, colXs[ci], ref yCol, ref remain, ref consumed);
                    colStates[ci] = (idx, line, subline);
                    if (colStates[ci] != startState) {
                        anyColumnAdvanced = true;
                    }

                    if (consumed > maxConsumed) maxConsumed = consumed;
                }

                if (!anyColumnAdvanced || maxConsumed <= 0.01) {
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
                if (AnyRemaining()) NewPage();
            }

            if (rowSpacingAfter > 0) {
                y -= rowSpacingAfter;
            }
        }

    }
}
