using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private void RenderHeadingFlowBlock(HeadingBlock hb, IPdfBlock? nextBlock, System.Collections.Generic.IList<IPdfBlock> blockList, int blockIndex) {
            double frameStart = GetCurrentFramePageStartY();
            PdfHeadingStyle? headingStyle = ResolveHeadingStyle(hb, currentOpts);
            double size = GetHeadingFontSize(hb, headingStyle);
            double leading = GetHeadingLeading(headingStyle, size);
            double spacingBefore = (y < frameStart - 0.001 || headingStyle?.ApplySpacingBeforeAtTop == true) ? headingStyle?.SpacingBefore ?? 0D : 0D;
            double spacingAfter = GetHeadingSpacingAfter(headingStyle, leading);
            var headingFont = GetHeadingFont(currentOpts, headingStyle);
            PdfColor? headingColor = hb.Color ?? headingStyle?.Color;
            System.Collections.Generic.IReadOnlyList<PdfTextRun> headingRuns = CreateHeadingTextRuns(hb, headingStyle, headingColor);
            var (lines, lineHeights) = WrapRichRunsCore(headingRuns, width, size, ChooseNormal(currentOpts.DefaultFont), leading, null, DefaultParagraphTabStopWidth, currentOpts);
            double textHeight = MeasureRichLinesHeight(lineHeights, lines.Count, leading);
            double needed = spacingBefore + textHeight + spacingAfter;
            bool keepWithNext = headingStyle?.KeepWithNext ?? true;
            if (keepWithNext && nextBlock != null) {
                double keepHeight = needed + MeasureKeepWithNextChainHeight(blockList, blockIndex + 1, currentOpts.MarginLeft, width, currentOpts.DefaultFontSize, needed);
                double availableHeight = frameStart - currentOpts.MarginBottom;
                if (keepHeight > needed + 0.001 && keepHeight <= availableHeight + 0.001 && y < frameStart - 0.001 && y - keepHeight < currentOpts.MarginBottom) {
                    NewPage();
                    spacingBefore = headingStyle?.ApplySpacingBeforeAtTop == true ? headingStyle.SpacingBefore : 0D;
                    needed = spacingBefore + textHeight + spacingAfter;
                }
            }

            if (y < frameStart - 0.001 && y - needed < currentOpts.MarginBottom) {
                NewPage();
                spacingBefore = headingStyle?.ApplySpacingBeforeAtTop == true ? headingStyle.SpacingBefore : 0D;
                needed = spacingBefore + textHeight + spacingAfter;
            }
            if (lineHeights.Count > 0 && spacingBefore + lineHeights[0] > frameStart - currentOpts.MarginBottom + 0.001)
                throw new ArgumentException("Heading spacing and first line exceed the available page content height.");
            if (spacingBefore > 0) {
                y -= spacingBefore;
            }

            int lineIndex = 0;
            PageStructElement? logicalHeading = null;
            while (lineIndex < lines.Count) {
                double available = y - currentOpts.MarginBottom;
                int take = 0;
                double heightSum = 0;
                for (int index = lineIndex; index < lines.Count; index++) {
                    double lineHeight = lineHeights[index];
                    if (heightSum + lineHeight > available + 0.001) break;
                    heightSum += lineHeight;
                    take++;
                }
                if (take == 0) {
                    if (y >= frameStart - 0.001 || lineHeights[lineIndex] > frameStart - currentOpts.MarginBottom + 0.001)
                        throw new ArgumentException("Heading line height exceeds the available page content height.");
                    NewPage();
                    continue;
                }
                var sliceLines = lines.GetRange(lineIndex, take);
                var sliceHeights = lineHeights.GetRange(lineIndex, take);
                EnsurePage();
                if (lineIndex == 0 && headingStyle?.AnchoredCanvas is { } headingCanvas) RenderCanvasBlock(headingCanvas);
                pageDirty = true;
                if (lineIndex == 0 && currentOpts.CreateOutlineFromHeadings) {
                    currentPage!.Bookmarks.Add(new PageBookmark { Level = hb.Level, Title = hb.Text, Y = y });
                }
                double firstBaseline = FirstTextBaselineFromTop(headingFont, size, y);
                string headingFontResource = GetHeadingFontResource(headingStyle);
                string structureType = "H" + hb.Level.ToString(CultureInfo.InvariantCulture);
                bool hasLinkTarget = !string.IsNullOrEmpty(hb.LinkUri) || !string.IsNullOrEmpty(hb.LinkDestinationName);
                int? linkStructElementIndex = null;
                string markedStructureType = structureType;
                int? markedContentId;
                if ((hasLinkTarget || take < lines.Count) && emitGeneratedStructure && currentPage != null) {
                    logicalHeading ??= RegisterStructureContainer(structureType, parentElement: null);
                    if (logicalHeading != null && lineIndex > 0) {
                        logicalHeading.SpansPages = true;
                        // The heading keeps its first-page parent; notify active ancestor scopes
                        // that their descendant has continued onto this page.
                        ResolveFlowSemanticParent();
                    }
                    if (hasLinkTarget) linkStructElementIndex = currentPage.StructElements.Count;
                    markedStructureType = hasLinkTarget ? "Link" : "Span";
                    markedContentId = RegisterTextStructureElement(markedStructureType, logicalHeading);
                } else {
                    markedContentId = RegisterTextStructureElement(structureType);
                }
                AddHeadingLinkAnnotations(hb, sliceLines, headingFont, size, leading, currentOpts.MarginLeft, width, firstBaseline, linkStructElementIndex);
                WriteRichParagraph(sb, new RichParagraphBlock(headingRuns, hb.Align, headingColor), sliceLines, sliceHeights, currentOpts, firstBaseline, size, leading, currentPage!.Annotations, currentOpts.MarginLeft, width, structureType: markedStructureType, markedContentId: markedContentId, structurePage: currentPage);
                MarkRichFonts(headingRuns);
                if (GetHeadingBold(headingStyle)) {
                    currentPage!.UsedBold = true;
                    usedBold = true;
                }
                y -= heightSum;
                lineIndex += take;
                if (lineIndex < lines.Count) NewPage();
            }
            y -= spacingAfter;
        }

        private void RenderRichParagraphFlowBlock(RichParagraphBlock rpb, IPdfBlock? nextBlock, System.Collections.Generic.IList<IPdfBlock> blockList, int blockIndex) {
            double frameStart = GetCurrentFramePageStartY();
            double size = currentOpts.DefaultFontSize;
            PdfParagraphStyle? paragraphStyle = EffectiveParagraphStyle(rpb);
            double leading = GetParagraphLeading(paragraphStyle, size);
            double spacingBefore = GetParagraphSpacingBefore(paragraphStyle);
            double spacingAfter = GetParagraphSpacingAfter(paragraphStyle, leading);
            var textFrame = GetParagraphTextFrame(paragraphStyle, currentOpts.MarginLeft, width);
            var (lines, lineHeights) = WrapRichRunsCoreWithFirstLineOrigin(rpb.Runs, textFrame.Width, size, ChooseNormal(currentOpts.DefaultFont), leading, textFrame.FirstLineWidth, textFrame.FirstLineX - textFrame.X, GetParagraphTabStopWidth(paragraphStyle), currentOpts, paragraphStyle?.TabStops.ToArray());
            if (paragraphStyle?.KeepWithNext == true && nextBlock != null && lines.Count > 0) {
                double paragraphHeight = spacingBefore + lineHeights.Sum() + spacingAfter;
                double nextHeight = MeasureKeepWithNextChainHeight(blockList, blockIndex + 1, currentOpts.MarginLeft, width, size, paragraphHeight);
                double keepHeight = paragraphHeight + nextHeight;
                double availableHeight = frameStart - currentOpts.MarginBottom;
                if (nextHeight > 0.001 && keepHeight <= availableHeight + 0.001 && y < frameStart - 0.001 && y - keepHeight < currentOpts.MarginBottom) {
                    NewPage();
                }
            }

            if (paragraphStyle?.KeepTogether == true) {
                double paragraphContentHeight = lineHeights.Sum();
                double availableHeight = frameStart - currentOpts.MarginBottom;
                if (paragraphContentHeight > availableHeight + 0.001) {
                    throw new ArgumentException("Paragraph height exceeds the available page content height.");
                }

                double paragraphHeight =
                    (y < frameStart - 0.001D ? spacingBefore : 0D) +
                    paragraphContentHeight;
                if (y < frameStart - 0.001 && y - paragraphHeight < currentOpts.MarginBottom) {
                    NewPage();
                }
            }

            List<double>? floatingLineOffsets = null;
            List<double>? floatingLineWidths = null;
            List<double>? floatingLineGaps = null;
            HashSet<int>? floatingPageStarts = null;
            if (HasFloatingTables) {
                floatingLineOffsets = new(); floatingLineWidths = new(); floatingLineGaps = new();
                floatingPageStarts = new();
                double wrapLeading = Math.Max(leading, rpb.Runs.Select(run => (run.FontSize ?? size) * leading / size).DefaultIfEmpty(leading).Max());
                double simulatedTop = y - (y < frameStart - 0.001 ? spacingBefore : 0);
                double previousHeight = 0;
                bool onOriginalPage = true;
                var wrapped = WrapRichRunsCoreWithFirstLineOrigin(rpb.Runs, textFrame.Width, size,
                    ChooseNormal(currentOpts.DefaultFont), leading, textFrame.FirstLineWidth,
                    textFrame.FirstLineX - textFrame.X, GetParagraphTabStopWidth(paragraphStyle), currentOpts,
                    paragraphStyle?.TabStops.ToArray(), (index, completedHeight) => {
                        simulatedTop -= completedHeight - previousHeight;
                        previousHeight = completedHeight;
                        if (simulatedTop - wrapLeading < currentOpts.MarginBottom) { simulatedTop = frameStart; onOriginalPage = false; floatingPageStarts.Add(index); }
                        double left = index == 0 ? textFrame.FirstLineX : textFrame.X;
                        double availableWidth = index == 0 ? textFrame.FirstLineWidth : textFrame.Width;
                        var frame = onOriginalPage ? GetFloatingTextFrame(left, availableWidth, simulatedTop, wrapLeading) : (X: left, Width: availableWidth, Gap: 0D);
                        if (simulatedTop - frame.Gap - wrapLeading < currentOpts.MarginBottom) {
                            frame = (left, availableWidth, 0D);
                            simulatedTop = frameStart;
                            onOriginalPage = false;
                            floatingPageStarts.Add(index);
                        }
                        floatingLineOffsets.Add(frame.X - textFrame.X);
                        floatingLineWidths.Add(frame.Width);
                        floatingLineGaps.Add(frame.Gap);
                        return (frame.Width, frame.X - textFrame.X, frame.Gap);
                    }, minimumLineHeight: wrapLeading);
                lines = wrapped.Lines; lineHeights = wrapped.LineHeights;
            }

            int lineIndex = 0;
            bool firstSegment = true;
            while (lineIndex < lines.Count) {
                if (floatingPageStarts?.Remove(lineIndex) == true && y < frameStart - 0.001) NewPage();
                double minimumLineHeight = lineHeights[lineIndex];
                if (minimumLineHeight > frameStart - currentOpts.MarginBottom)
                    throw new ArgumentException("Paragraph line height exceeds the available page content height.");
                double available = y - currentOpts.MarginBottom;
                if (available <= 0) {
                    NewPage();
                    firstSegment = false;
                    continue;
                }

                double segmentSpacingBefore = firstSegment && y < frameStart - 0.001 ? spacingBefore : 0;
                if (available < segmentSpacingBefore + minimumLineHeight) {
                    NewPage();
                    available = y - currentOpts.MarginBottom;
                    if (y >= frameStart - 0.001) {
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
                    if (k > lineIndex && floatingPageStarts?.Contains(k) == true) break;
                    double lineHeight = lineHeights[k];
                    if (heightSum + lineHeight > roomForText) {
                        break;
                    }

                    heightSum += lineHeight;
                    take++;
                }

                if (TryApplyWidowControl(paragraphStyle, lines.Count, lineIndex, ref take, ref heightSum, lineHeights, y < frameStart - 0.001)) {
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
                if (sliceStartsAtFirstLine && paragraphStyle?.AnchoredCanvas is { } paragraphCanvas) RenderCanvasBlock(paragraphCanvas);
                pageDirty = true;
                var paragraphFont = ChooseNormal(currentOpts.DefaultFont);
                int? markedContentId = RegisterTextStructureElement("P");
                MarkRichFonts(rpb.Runs);
                WriteRichParagraph(sb, rpb, sliceLines, sliceHeights, currentOpts, FirstTextBaselineFromTop(paragraphFont, size, y), size, leading, currentPage!.Annotations, textFrame.X, textFrame.Width, sliceStartsAtFirstLine ? textFrame.FirstLineX : null, sliceStartsAtFirstLine ? textFrame.FirstLineWidth : null, "P", markedContentId, currentPage,
                    lineXOffsets: floatingLineOffsets?.GetRange(lineIndex, take), lineWidths: floatingLineWidths?.GetRange(lineIndex, take), lineTopGaps: floatingLineGaps?.GetRange(lineIndex, take));
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
