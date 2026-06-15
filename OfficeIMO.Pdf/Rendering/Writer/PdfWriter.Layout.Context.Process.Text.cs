using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private void RenderHeadingFlowBlock(HeadingBlock hb, IPdfBlock? nextBlock) {
            PdfHeadingStyle? headingStyle = ResolveHeadingStyle(hb, currentOpts);
            double size = GetHeadingFontSize(hb, headingStyle);
            double leading = GetHeadingLeading(headingStyle, size);
            double spacingBefore = (y < yStart - 0.001 || headingStyle?.ApplySpacingBeforeAtTop == true) ? headingStyle?.SpacingBefore ?? 0D : 0D;
            double spacingAfter = GetHeadingSpacingAfter(headingStyle, leading);
            var headingFont = GetHeadingFont(currentOpts, headingStyle);
            PdfColor? headingColor = hb.Color ?? headingStyle?.Color;
            System.Collections.Generic.IReadOnlyList<TextRun> headingRuns = CreateHeadingTextRuns(hb, headingStyle, headingColor);
            var (lines, lineHeights) = WrapRichRunsCore(headingRuns, width, size, ChooseNormal(currentOpts.DefaultFont), leading, null, DefaultParagraphTabStopWidth, currentOpts);
            double textHeight = MeasureRichLinesHeight(lineHeights, lines.Count, leading);
            double needed = spacingBefore + textHeight + spacingAfter;
            bool keepWithNext = headingStyle?.KeepWithNext ?? true;
            if (keepWithNext && nextBlock != null) {
                double keepHeight = needed + MeasureNextBlockFirstVisualHeight(nextBlock, currentOpts.MarginLeft, width, currentOpts.DefaultFontSize);
                double availableHeight = currentOpts.PageHeight - currentOpts.MarginTop - currentOpts.MarginBottom;
                if (keepHeight > needed + 0.001 && keepHeight <= availableHeight + 0.001 && y < yStart - 0.001 && y - keepHeight < currentOpts.MarginBottom) {
                    NewPage();
                    spacingBefore = headingStyle?.ApplySpacingBeforeAtTop == true ? headingStyle.SpacingBefore : 0D;
                    needed = spacingBefore + textHeight + spacingAfter;
                }
            }

            if (y - needed < currentOpts.MarginBottom) {
                NewPage();
                spacingBefore = headingStyle?.ApplySpacingBeforeAtTop == true ? headingStyle.SpacingBefore : 0D;
                needed = spacingBefore + textHeight + spacingAfter;
            }
            if (spacingBefore > 0) {
                y -= spacingBefore;
            }

            EnsurePage();
            pageDirty = true;
            if (currentOpts.CreateOutlineFromHeadings) {
                currentPage!.Bookmarks.Add(new PageBookmark { Level = hb.Level, Title = hb.Text, Y = y });
            }
            double firstBaseline = FirstTextBaselineFromTop(headingFont, size, y);
            string headingFontResource = GetHeadingFontResource(headingStyle);
            string structureType = "H" + hb.Level.ToString(CultureInfo.InvariantCulture);
            bool hasLinkTarget = !string.IsNullOrEmpty(hb.LinkUri) || !string.IsNullOrEmpty(hb.LinkDestinationName);
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

            AddHeadingLinkAnnotations(hb, lines, headingFont, size, leading, currentOpts.MarginLeft, width, firstBaseline, linkStructElementIndex);
            WriteRichParagraph(sb, new RichParagraphBlock(headingRuns, hb.Align, headingColor), lines, lineHeights, currentOpts, firstBaseline, size, leading, currentPage!.Annotations, currentOpts.MarginLeft, width, structureType: markedStructureType, markedContentId: markedContentId, structurePage: currentPage);
            MarkRichFonts(headingRuns);
            if (GetHeadingBold(headingStyle)) {
                currentPage!.UsedBold = true;
                usedBold = true;
            }
            y -= textHeight + spacingAfter;
        }

        private void RenderRichParagraphFlowBlock(RichParagraphBlock rpb, IPdfBlock? nextBlock) {
            double size = currentOpts.DefaultFontSize;
            PdfParagraphStyle? paragraphStyle = EffectiveParagraphStyle(rpb);
            double leading = GetParagraphLeading(paragraphStyle, size);
            double spacingBefore = GetParagraphSpacingBefore(paragraphStyle);
            double spacingAfter = GetParagraphSpacingAfter(paragraphStyle, leading);
            var textFrame = GetParagraphTextFrame(paragraphStyle, currentOpts.MarginLeft, width);
            var (lines, lineHeights) = WrapRichRunsCore(rpb.Runs, textFrame.Width, size, ChooseNormal(currentOpts.DefaultFont), leading, textFrame.FirstLineWidth, GetParagraphTabStopWidth(paragraphStyle), currentOpts, GetParagraphTabStops(paragraphStyle));
            if (paragraphStyle?.KeepWithNext == true && nextBlock != null && lines.Count > 0) {
                double nextHeight = MeasureNextBlockFirstVisualHeight(nextBlock, currentOpts.MarginLeft, width, size);
                double keepHeight = spacingBefore + lineHeights.Sum() + spacingAfter + nextHeight;
                double availableHeight = currentOpts.PageHeight - currentOpts.MarginTop - currentOpts.MarginBottom;
                if (nextHeight > 0.001 && keepHeight <= availableHeight + 0.001 && y < yStart - 0.001 && y - keepHeight < currentOpts.MarginBottom) {
                    NewPage();
                }
            }

            if (paragraphStyle?.KeepTogether == true) {
                double paragraphHeight = spacingBefore + lineHeights.Sum();
                double availableHeight = currentOpts.PageHeight - currentOpts.MarginTop - currentOpts.MarginBottom;
                if (paragraphHeight > availableHeight + 0.001) {
                    throw new ArgumentException("Paragraph height exceeds the available page content height.");
                }

                if (y < yStart - 0.001 && y - paragraphHeight < currentOpts.MarginBottom) {
                    NewPage();
                }
            }

            int lineIndex = 0;
            bool firstSegment = true;
            while (lineIndex < lines.Count) {
                double available = y - currentOpts.MarginBottom;
                if (available <= 0.5) {
                    NewPage();
                    firstSegment = false;
                    continue;
                }

                double segmentSpacingBefore = firstSegment && y < yStart - 0.001 ? spacingBefore : 0;
                double minimumLineHeight = lineHeights[lineIndex];
                if (available < segmentSpacingBefore + minimumLineHeight) {
                    NewPage();
                    available = y - currentOpts.MarginBottom;
                    if (y >= yStart - 0.001) {
                        segmentSpacingBefore = 0;
                    }
                    if (available < segmentSpacingBefore + minimumLineHeight) {
                        segmentSpacingBefore = Math.Max(0, available - minimumLineHeight);
                    }
                }

                double roomForText = Math.Max(0, available - segmentSpacingBefore);
                int take = 0;
                double heightSum = 0;
                for (int k = lineIndex; k < lines.Count; k++) {
                    double lineHeight = lineHeights[k];
                    if (heightSum + lineHeight > roomForText) {
                        break;
                    }

                    heightSum += lineHeight;
                    take++;
                }

                if (TryApplyWidowControl(paragraphStyle, lines.Count, lineIndex, ref take, ref heightSum, lineHeights, y < yStart - 0.001)) {
                    NewPage();
                    firstSegment = false;
                    continue;
                }

                if (take == 0) {
                    NewPage();
                    firstSegment = false;
                    continue;
                }

                if (segmentSpacingBefore > 0) {
                    y -= segmentSpacingBefore;
                }

                var sliceLines = new System.Collections.Generic.List<System.Collections.Generic.List<RichSeg>>();
                var sliceHeights = new System.Collections.Generic.List<double>();
                for (int k = 0; k < take; k++) {
                    sliceLines.Add(lines[lineIndex + k]);
                    sliceHeights.Add(lineHeights[lineIndex + k]);
                }

                bool sliceStartsAtFirstLine = lineIndex == 0;
                pageDirty = true;
                var paragraphFont = ChooseNormal(currentOpts.DefaultFont);
                int? markedContentId = RegisterTextStructureElement("P");
                WriteRichParagraph(sb, rpb, sliceLines, sliceHeights, currentOpts, FirstTextBaselineFromTop(paragraphFont, size, y), size, leading, currentPage!.Annotations, textFrame.X, textFrame.Width, sliceStartsAtFirstLine ? textFrame.FirstLineX : null, sliceStartsAtFirstLine ? textFrame.FirstLineWidth : null, "P", markedContentId, currentPage);
                y -= heightSum;
                lineIndex += take;
                firstSegment = false;
                if (lineIndex < lines.Count) {
                    NewPage();
                } else {
                    y -= spacingAfter;
                }
            }

            MarkRichFonts(rpb.Runs);
        }

    }
}
