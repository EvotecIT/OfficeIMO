using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private void RenderHorizontalRuleFlowBlock(HorizontalRuleBlock hr, IPdfBlock? nextBlock) {
            PdfHorizontalRuleStyle ruleStyle = ResolveHorizontalRuleStyle(hr, currentOpts);
            ValidateHorizontalRule(ruleStyle);
            if (ruleStyle.KeepWithNext && nextBlock != null) {
                double needed = ruleStyle.SpacingBefore + ruleStyle.Thickness + ruleStyle.SpacingAfter;
                double nextHeight = MeasureNextBlockFirstVisualHeight(nextBlock, currentOpts.MarginLeft, width, currentOpts.DefaultFontSize);
                KeepFixedBlockWithNext(needed, nextHeight);
            }

            RenderHorizontalRuleBlock(hr, currentOpts.MarginLeft, width);
        }

        private void RenderShapeFlowBlock(ShapeBlock sbk, IPdfBlock? nextBlock) {
            PdfDrawingStyle shapeStyle = ResolveDrawingStyle(sbk, currentOpts);
            PdfDocument.ValidateDrawingStyle(shapeStyle, "Shape");
            if (shapeStyle.KeepWithNext && nextBlock != null) {
                double needed = shapeStyle.SpacingBefore + sbk.Shape.Height + shapeStyle.SpacingAfter;
                double nextHeight = MeasureNextBlockFirstVisualHeight(nextBlock, currentOpts.MarginLeft, width, currentOpts.DefaultFontSize);
                KeepFixedBlockWithNext(needed, nextHeight);
            }

            RenderShapeBlock(sbk, currentOpts.MarginLeft, width);
        }

        private void RenderDrawingFlowBlock(DrawingBlock dbk, IPdfBlock? nextBlock) {
            PdfDrawingStyle drawingStyle = ResolveDrawingStyle(dbk, currentOpts);
            PdfDocument.ValidateDrawingStyle(drawingStyle, "Drawing");
            if (drawingStyle.KeepWithNext && nextBlock != null) {
                double needed = drawingStyle.SpacingBefore + dbk.Drawing.Height + drawingStyle.SpacingAfter;
                double nextHeight = MeasureNextBlockFirstVisualHeight(nextBlock, currentOpts.MarginLeft, width, currentOpts.DefaultFontSize);
                KeepFixedBlockWithNext(needed, nextHeight);
            }

            RenderDrawingBlock(dbk, currentOpts.MarginLeft, width);
        }

        private void RenderImageFlowBlock(ImageBlock ib, IPdfBlock? nextBlock) {
            double xImg = currentOpts.MarginLeft;
            double contentWidth = currentOpts.PageWidth - currentOpts.MarginLeft - currentOpts.MarginRight;
            PdfImageStyle imageStyle = ResolveImageStyle(ib, currentOpts);
            PdfDocument.ValidateImageStyleForBox(imageStyle, ib.Width, ib.Height, nameof(imageStyle.ClipPath));
            PdfDocument.ValidateImageFitDimensions(ib.Info, imageStyle.Fit, nameof(imageStyle.Fit));
            double imageSpacingBefore = ResolveTopLevelSpacingBefore(imageStyle.SpacingBefore);
            var imageBox = ResolveImageFlowBox(ib, imageStyle, contentWidth, imageSpacingBefore, imageStyle.SpacingAfter);
            double needed = imageSpacingBefore + imageBox.Height + imageStyle.SpacingAfter;
            if (imageStyle.Align == PdfAlign.Center) xImg = currentOpts.MarginLeft + Math.Max(0, (contentWidth - imageBox.Width) / 2);
            else if (imageStyle.Align == PdfAlign.Right) xImg = currentOpts.MarginLeft + Math.Max(0, contentWidth - imageBox.Width);
            EnsureFixedFlowBlockFits("Image", imageBox.Width, needed, contentWidth);
            if (imageStyle.KeepWithNext && nextBlock != null) {
                double nextHeight = MeasureNextBlockFirstVisualHeight(nextBlock, currentOpts.MarginLeft, width, currentOpts.DefaultFontSize);
                double keepHeight = needed + nextHeight;
                double availableHeight = currentOpts.PageHeight - currentOpts.MarginTop - currentOpts.MarginBottom;
                if (nextHeight > 0.001 && keepHeight <= availableHeight + 0.001 && y < yStart - 0.001 && y - keepHeight < currentOpts.MarginBottom) {
                    NewPage();
                    imageSpacingBefore = 0D;
                    imageBox = ResolveImageFlowBox(ib, imageStyle, contentWidth, imageSpacingBefore, imageStyle.SpacingAfter);
                    needed = imageBox.Height + imageStyle.SpacingAfter;
                    if (imageStyle.Align == PdfAlign.Center) xImg = currentOpts.MarginLeft + Math.Max(0, (contentWidth - imageBox.Width) / 2);
                    else if (imageStyle.Align == PdfAlign.Right) xImg = currentOpts.MarginLeft + Math.Max(0, contentWidth - imageBox.Width);
                }
            }

            if (y - needed < currentOpts.MarginBottom) {
                NewPage();
                imageSpacingBefore = 0D;
            }
            if (imageSpacingBefore > 0) y -= imageSpacingBefore;
            EnsurePage();
            PageImage pageImage = CreatePageImage(ib, imageStyle, xImg, y - imageBox.Height, imageBox.Width, imageBox.Height);
            currentPage!.Images.Add(pageImage);
            if (!string.IsNullOrWhiteSpace(pageImage.AlternativeText)) {
                int? markedContentId = RegisterFigureStructureElement(pageImage.AlternativeText!);
                pageImage.MarkedContentId = markedContentId;
                pageImage.StructElementIndex = FindStructElementIndex(currentPage, markedContentId, "Figure");
            }

            AddImageLinkAnnotation(ib, imageStyle, pageImage, xImg, y - imageBox.Height, imageBox.Width, imageBox.Height);
            pageDirty = true;
            y -= imageBox.Height + imageStyle.SpacingAfter;
        }

        private void RenderPanelFlowBlock(PanelParagraphBlock ppb, IPdfBlock? nextBlock) {
            double size = currentOpts.DefaultFontSize;
            double leading = size * 1.4;
            var panelFont = ChooseNormal(currentOpts.DefaultFont);
            double firstBaselineOffset = GetAscenderForOptions(panelFont, size, currentOpts);
            double contentWidth = currentOpts.PageWidth - currentOpts.MarginLeft - currentOpts.MarginRight;
            PanelStyle panelStyle = ResolvePanelStyle(ppb, currentOpts);
            double innerWidth = panelStyle.MaxWidth.HasValue ? Math.Min(contentWidth, panelStyle.MaxWidth.Value) : contentWidth;
            ValidatePanelStyle(panelStyle, innerWidth);
            double textWidthAvail = innerWidth - 2 * panelStyle.PaddingX;
            var (lines, lineHeights) = WrapRichRunsCore(ppb.Runs, textWidthAvail, size, panelFont, leading, null, DefaultParagraphTabStopWidth, currentOpts);
            double panelWidth = innerWidth;
            double xLeft = currentOpts.MarginLeft;
            if (panelStyle.Align == PdfAlign.Center) xLeft = currentOpts.MarginLeft + Math.Max(0, (contentWidth - innerWidth) / 2);
            else if (panelStyle.Align == PdfAlign.Right) xLeft = currentOpts.MarginLeft + Math.Max(0, contentWidth - innerWidth);
            double panelSpacingBefore = ResolveTopLevelSpacingBefore(panelStyle.SpacingBefore);

            if (panelStyle.KeepWithNext && nextBlock != null && lines.Count > 0) {
                double panelHeight = panelSpacingBefore + panelStyle.PaddingY + lineHeights.Sum() + panelStyle.PaddingY + panelStyle.SpacingAfter;
                double nextHeight = MeasureNextBlockFirstVisualHeight(nextBlock, currentOpts.MarginLeft, width, size);
                double keepHeight = panelHeight + nextHeight;
                double availableHeight = currentOpts.PageHeight - currentOpts.MarginTop - currentOpts.MarginBottom;
                if (nextHeight > 0.001 && keepHeight <= availableHeight + 0.001 && y < yStart - 0.001 && y - keepHeight < currentOpts.MarginBottom) {
                    NewPage();
                    panelSpacingBefore = 0D;
                }
            }

            if (panelSpacingBefore > 0) {
                if (y - panelSpacingBefore < currentOpts.MarginBottom) {
                    NewPage();
                    panelSpacingBefore = 0D;
                }

                if (panelSpacingBefore > 0) y -= panelSpacingBefore;
            }

            if (panelStyle.KeepTogether) {
                double textHeight = lineHeights.Sum();
                double panelHeight = panelStyle.PaddingY + textHeight + panelStyle.PaddingY;
                double availableHeight = currentOpts.PageHeight - currentOpts.MarginTop - currentOpts.MarginBottom;
                if (panelHeight > availableHeight + 0.001) {
                    throw new ArgumentException("Panel height exceeds the available page content height.");
                }

                double panelTop = y;
                double panelBottom = y - panelHeight;
                if (panelBottom < currentOpts.MarginBottom) { NewPage(); panelTop = y; panelBottom = y - panelHeight; }
                if (panelStyle.Background.HasValue) { pageDirty = true; DrawRowFill(sb, panelStyle.Background.Value, xLeft, panelBottom, panelWidth, panelTop - panelBottom, emitGeneratedStructure); }
                if (DrawPanelBorder(sb, panelStyle, xLeft, panelBottom, panelWidth, panelTop - panelBottom, emitGeneratedStructure)) { pageDirty = true; }
                pageDirty = true;
                int? panelMarkedContentId = RegisterTextStructureElement("P");
                WriteRichParagraph(sb, new RichParagraphBlock(ppb.Runs, ppb.Align, ppb.DefaultColor), lines, lineHeights, currentOpts, panelTop - panelStyle.PaddingY - firstBaselineOffset, size, leading, currentPage!.Annotations, xLeft + panelStyle.PaddingX, textWidthAvail, structureType: "P", markedContentId: panelMarkedContentId, structurePage: currentPage);
                MarkRichFonts(ppb.Runs);
                y = panelBottom;
                if (panelStyle.SpacingAfter > 0) {
                    if (y < yStart - 0.001 && y - panelStyle.SpacingAfter < currentOpts.MarginBottom) {
                        NewPage();
                    } else {
                        y -= panelStyle.SpacingAfter;
                    }
                }
            } else {
                int li = 0; bool firstSeg = true;
                while (li < lines.Count) {
                    double avail = y - currentOpts.MarginBottom;
                    double topPad = firstSeg ? panelStyle.PaddingY : 0;
                    double minLine = lineHeights[li];
                    if (avail < topPad + minLine) {
                        EnsurePanelSegmentCanFitLine(topPad, minLine);
                        NewPage();
                        continue;
                    }

                    double roomForText = avail - topPad - panelStyle.PaddingY;
                    if (roomForText < minLine) {
                        roomForText = avail - topPad;
                    }

                    int take = 0; double hsum = 0;
                    for (int k = li; k < lines.Count; k++) {
                        double h = lineHeights[k];
                        if (hsum + h > roomForText) break;
                        hsum += h; take++;
                    }

                    if (take == 0) {
                        EnsurePanelSegmentCanFitLine(topPad, minLine);
                        NewPage();
                        continue;
                    }

                    bool lastSeg = (li + take) >= lines.Count;
                    double panelTop = y;
                    double usedBottomPad = lastSeg ? panelStyle.PaddingY : Math.Max(0, avail - (topPad + hsum));
                    double panelBottom = y - (topPad + hsum + usedBottomPad);
                    if (panelStyle.Background.HasValue) { pageDirty = true; DrawRowFill(sb, panelStyle.Background.Value, xLeft, panelBottom, panelWidth, panelTop - panelBottom, emitGeneratedStructure); }
                    if (DrawPanelBorder(sb, panelStyle, xLeft, panelBottom, panelWidth, panelTop - panelBottom, emitGeneratedStructure)) { pageDirty = true; }
                    var sliceLines = new System.Collections.Generic.List<System.Collections.Generic.List<RichSeg>>();
                    var sliceHeights = new System.Collections.Generic.List<double>();
                    for (int k = 0; k < take; k++) { sliceLines.Add(lines[li + k]); sliceHeights.Add(lineHeights[li + k]); }
                    pageDirty = true;
                    int? panelMarkedContentId = RegisterTextStructureElement("P");
                    WriteRichParagraph(sb, new RichParagraphBlock(ppb.Runs, ppb.Align, ppb.DefaultColor), sliceLines, sliceHeights, currentOpts, panelTop - topPad - firstBaselineOffset, size, leading, currentPage!.Annotations, xLeft + panelStyle.PaddingX, textWidthAvail, structureType: "P", markedContentId: panelMarkedContentId, structurePage: currentPage);
                    MarkRichFonts(ppb.Runs);
                    y = panelBottom; li += take; firstSeg = false;
                    if (li < lines.Count) {
                        NewPage();
                    } else if (panelStyle.SpacingAfter > 0) {
                        if (y < yStart - 0.001 && y - panelStyle.SpacingAfter < currentOpts.MarginBottom) {
                            NewPage();
                        } else {
                            y -= panelStyle.SpacingAfter;
                        }
                    }
                }
            }
        }

    }
}
