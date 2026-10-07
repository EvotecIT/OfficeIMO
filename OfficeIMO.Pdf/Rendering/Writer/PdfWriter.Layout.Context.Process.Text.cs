using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private void RenderHeadingFlowBlock(HeadingBlock hb, IPdfBlock? nextBlock, System.Collections.Generic.IList<IPdfBlock> blockList, int blockIndex, Action? captureFirstPlacement = null) {
            double frameStart = GetCurrentFramePageStartY();
            PdfHeadingStyle? headingStyle = ResolveHeadingStyle(hb, currentOpts);
            double size = GetHeadingFontSize(hb, headingStyle);
            double leading = GetHeadingLeading(headingStyle, size);
            double spacingBefore = (y < frameStart - 0.001 || headingStyle?.ApplySpacingBeforeAtTop == true) ? headingStyle?.SpacingBefore ?? 0D : 0D;
            double spacingAfter = GetHeadingSpacingAfter(headingStyle, leading);
            var headingFont = GetHeadingFont(currentOpts, headingStyle);
            PdfColor? headingColor = hb.Color ?? headingStyle?.Color;
            System.Collections.Generic.IReadOnlyList<PdfTextRun> headingRuns = CreateHeadingTextRuns(hb, headingStyle, headingColor);
            var (lines, lineHeights) = WrapRichRunsWithSpacing(headingRuns, width, size, ChooseNormal(currentOpts.DefaultFont), leading, null, DefaultParagraphTabStopWidth, currentOpts, headingStyle?.LineSpacing);
            double wrappedWidth = width;
            int lineIndex = 0;
            bool firstLogicalLine = true;
            double textHeight = MeasureRichLinesHeight(lineHeights, lines.Count, leading);
            void NewHeadingFrame() {
                if (firstLogicalLine) QueueColumnBalanceRemainder(hb, blockList, blockIndex);
                else QueueColumnBalanceRemainder(new List<ColumnBalanceUnit> {
                    new(lineHeights.Skip(lineIndex).Sum() + spacingAfter)
                }, hb, blockList, blockIndex);
                NewPage();
                frameStart = GetCurrentFramePageStartY();
                if (activeColumnFlow != null && Math.Abs(wrappedWidth - width) > 0.001D) {
                    var remainingRuns = firstLogicalLine ? headingRuns : BuildTextRunsFromWrappedLines(lines, lineIndex, lines.Count - lineIndex);
                    (lines, lineHeights) = WrapRichRunsWithSpacing(remainingRuns, width, size, ChooseNormal(currentOpts.DefaultFont), leading, null,
                        DefaultParagraphTabStopWidth, currentOpts, headingStyle?.LineSpacing);
                    lineIndex = 0;
                    wrappedWidth = width;
                    textHeight = MeasureRichLinesHeight(lineHeights, lines.Count, leading);
                }
            }
            double needed = spacingBefore + textHeight + spacingAfter;
            bool keepWithNext = headingStyle?.KeepWithNext ?? true;
            if (keepWithNext && nextBlock != null) {
                double keepHeight = needed + MeasureCurrentFrameKeepNextHeight(blockList, blockIndex + 1, currentOpts.MarginLeft, width, currentOpts.DefaultFontSize, needed);
                double availableHeight = GetMaximumBlockContinuationHeight();
                while (keepHeight > needed + 0.001 && keepHeight <= availableHeight + 0.001 && ShouldAdvanceForBlockHeight(keepHeight)) {
                    NewHeadingFrame();
                    spacingBefore = headingStyle?.ApplySpacingBeforeAtTop == true ? headingStyle.SpacingBefore : 0D;
                    needed = spacingBefore + textHeight + spacingAfter;
                    keepHeight = needed + MeasureCurrentFrameKeepNextHeight(blockList, blockIndex + 1, currentOpts.MarginLeft, width, currentOpts.DefaultFontSize, needed);
                    availableHeight = GetMaximumBlockContinuationHeight();
                }
            }

            if (y < frameStart - 0.001 && y - needed < currentOpts.MarginBottom) {
                NewHeadingFrame();
                spacingBefore = headingStyle?.ApplySpacingBeforeAtTop == true ? headingStyle.SpacingBefore : 0D;
                needed = spacingBefore + textHeight + spacingAfter;
            }
            while (lineHeights.Count > 0 && ShouldAdvanceForBlockHeight(spacingBefore + lineHeights[0])) {
                NewHeadingFrame();
                spacingBefore = headingStyle?.ApplySpacingBeforeAtTop == true ? headingStyle.SpacingBefore : 0D;
            }
            if (lineHeights.Count > 0 && spacingBefore + lineHeights[0] > frameStart - currentOpts.MarginBottom + 0.001)
                throw new ArgumentException("Heading spacing and first line exceed the available page content height.");
            if (spacingBefore > 0) {
                y -= spacingBefore;
            }

            PageStructElement? logicalHeading = null;
            while (lineIndex < lines.Count) {
                double available = y - currentOpts.MarginBottom;
                int take = 0;
                double heightSum = 0;
                for (int index = lineIndex; index < lines.Count; index++) {
                    double lineHeight = lineHeights[index];
                    double closingPadding = index == lines.Count - 1 ? GetClosingTextPadding(spacingAfter) : 0D;
                    if (heightSum + lineHeight + closingPadding > available + 0.001) break;
                    heightSum += lineHeight;
                    take++;
                }
                if (take == 0) {
                    if (!ShouldAdvanceForBlockHeight(lineHeights[lineIndex] + (lineIndex == lines.Count - 1 ? GetClosingTextPadding(spacingAfter) : 0D)))
                        throw new ArgumentException("Heading line height exceeds the available page content height.");
                    NewHeadingFrame();
                    continue;
                }
                var sliceLines = lines.GetRange(lineIndex, take);
                var sliceHeights = lineHeights.GetRange(lineIndex, take);
                EnsurePage();
                RecordFlowPlacement(y);
                if (firstLogicalLine) captureFirstPlacement?.Invoke();
                if (firstLogicalLine && headingStyle?.AnchoredCanvas is { } headingCanvas) RenderParagraphCanvas(headingCanvas, y);
                pageDirty = true;
                if (firstLogicalLine && currentOpts.CreateOutlineFromHeadings) {
                    currentPage!.Bookmarks.Add(new PageBookmark { Level = hb.Level, Title = hb.Text, Y = y });
                }
                double firstBaseline = FirstTextBaselineFromTop(headingFont, size, y);
                string headingFontResource = GetHeadingFontResource(headingStyle);
                string structureType = "H" + hb.Level.ToString(CultureInfo.InvariantCulture);
                bool hasLinkTarget = !string.IsNullOrEmpty(hb.LinkUri) || !string.IsNullOrEmpty(hb.LinkDestinationName);
                int? linkStructElementIndex = null;
                string markedStructureType = structureType;
                int? markedContentId;
                if ((logicalHeading != null || hasLinkTarget || take < lines.Count) && emitGeneratedStructure && currentPage != null) {
                    logicalHeading ??= RegisterStructureContainer(structureType, parentElement: null);
                    if (logicalHeading != null && !firstLogicalLine) {
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
                AddHeadingLinkAnnotations(hb, sliceLines, headingFont, size, leading, currentOpts.MarginLeft, width, firstBaseline, linkStructElementIndex, sliceHeights);
                WriteRichParagraph(sb, new RichParagraphBlock(headingRuns, hb.Align, headingColor), sliceLines, sliceHeights, currentOpts, firstBaseline, size, leading, currentPage!.Annotations, currentOpts.MarginLeft, width, structureType: markedStructureType, markedContentId: markedContentId, structurePage: currentPage, baselineFont: headingFont);
                MarkRichFonts(headingRuns);
                if (GetHeadingBold(headingStyle)) {
                    currentPage!.UsedBold = true;
                    usedBold = true;
                }
                y -= heightSum;
                lineIndex += take;
                firstLogicalLine = false;
                if (lineIndex < lines.Count) NewHeadingFrame();
            }
            y -= spacingAfter;
        }

        private void RenderRichParagraphFlowBlock(RichParagraphBlock rpb, IPdfBlock? nextBlock, System.Collections.Generic.IList<IPdfBlock> blockList, int blockIndex) {
            double frameStart = GetCurrentFramePageStartY();
            PdfParagraphStyle? paragraphStyle = EffectiveParagraphStyle(rpb);
            double size = paragraphStyle?.FontSize ?? currentOpts.DefaultFontSize;
            double leading = GetParagraphLeading(paragraphStyle, size);
            double spacingBefore = GetParagraphSpacingBefore(paragraphStyle);
            double spacingAfter = GetParagraphSpacingAfter(paragraphStyle, leading);
            var textFrame = GetParagraphTextFrame(paragraphStyle, currentOpts.MarginLeft, width);
            double wrappedWidth = textFrame.Width;
            var (lines, lineHeights) = WrapRichRunsCoreWithFirstLineOrigin(rpb.Runs, textFrame.Width, size, ChooseNormal(currentOpts.DefaultFont), leading, textFrame.FirstLineWidth, textFrame.FirstLineX - textFrame.X, GetParagraphTabStopWidth(paragraphStyle), currentOpts, paragraphStyle?.TabStops.ToArray(), lineSpacing: paragraphStyle?.LineSpacing);
            if (activeColumnFlow is { Options.BalanceParagraphLines: false, Options.BalanceLastPage: true } &&
                lineHeights.Sum() <= GetMaximumBlockContinuationHeight() + 0.001D) {
                paragraphStyle = paragraphStyle?.Clone() ?? new PdfParagraphStyle();
                paragraphStyle.KeepTogether = true;
            }
            void NewParagraphStartFrame() {
                NewBlockFrame();
                frameStart = GetCurrentFramePageStartY();
                textFrame = GetParagraphTextFrame(paragraphStyle, currentOpts.MarginLeft, width);
                if (Math.Abs(wrappedWidth - textFrame.Width) > .001D) {
                    (lines, lineHeights) = WrapRichRunsCoreWithFirstLineOrigin(rpb.Runs, textFrame.Width, size, ChooseNormal(currentOpts.DefaultFont), leading,
                        textFrame.FirstLineWidth, textFrame.FirstLineX - textFrame.X, GetParagraphTabStopWidth(paragraphStyle), currentOpts,
                        paragraphStyle?.TabStops.ToArray(), lineSpacing: paragraphStyle?.LineSpacing);
                    wrappedWidth = textFrame.Width;
                }
            }
            bool IsKeptInCurrentFrame() => paragraphStyle?.KeepTogether == true &&
                !(activeColumnFlow is { Options.BalanceKeptParagraphLines: true } && IsBalancedColumnFrame() &&
                  lineHeights.Sum() > GetCurrentFramePageStartY() - currentOpts.MarginBottom + .001D);
            if (paragraphStyle?.KeepWithNext == true && nextBlock != null && lines.Count > 0 &&
                (ReservesWholeKeepGroup() || IsKeptInCurrentFrame())) {
                double paragraphHeight = ResolveTopLevelSpacingBefore(spacingBefore) + lineHeights.Sum() + spacingAfter;
                double nextHeight = MeasureCurrentFrameKeepNextHeight(blockList, blockIndex + 1, currentOpts.MarginLeft, width, size, paragraphHeight);
                double keepHeight = paragraphHeight + nextHeight;
                while ((ReservesWholeKeepGroup() || IsKeptInCurrentFrame()) && nextHeight > 0.001 &&
                    keepHeight <= GetMaximumBlockContinuationHeight() + 0.001 && ShouldAdvanceForBlockHeight(keepHeight)) {
                    NewParagraphStartFrame();
                    paragraphHeight = lineHeights.Sum() + spacingAfter;
                    nextHeight = MeasureCurrentFrameKeepNextHeight(blockList, blockIndex + 1, currentOpts.MarginLeft, width, size, paragraphHeight);
                    keepHeight = paragraphHeight + nextHeight;
                }
            }

            if (IsKeptInCurrentFrame()) {
                while (true) {
                    if (!IsKeptInCurrentFrame()) break;
                    double paragraphContentHeight = lineHeights.Sum();
                    if (paragraphContentHeight > GetMaximumBlockContinuationHeight() + 0.001)
                        throw new ArgumentException("Paragraph height exceeds the available page content height.");
                    double paragraphHeight = ResolveTopLevelSpacingBefore(spacingBefore) + paragraphContentHeight;
                    if (!ShouldAdvanceForBlockHeight(paragraphHeight)) break;
                    NewParagraphStartFrame();
                }
            }

            List<double>? floatingLineOffsets = null;
            textFrame = GetParagraphTextFrame(paragraphStyle, currentOpts.MarginLeft, width);
            if (activeColumnFlow != null && Math.Abs(wrappedWidth - textFrame.Width) > 0.001D) {
                (lines, lineHeights) = WrapRichRunsCoreWithFirstLineOrigin(rpb.Runs, textFrame.Width, size, ChooseNormal(currentOpts.DefaultFont), leading,
                    textFrame.FirstLineWidth, textFrame.FirstLineX - textFrame.X, GetParagraphTabStopWidth(paragraphStyle), currentOpts,
                    paragraphStyle?.TabStops.ToArray(), lineSpacing: paragraphStyle?.LineSpacing);
                wrappedWidth = textFrame.Width;
                frameStart = GetCurrentFramePageStartY();
            }
            List<double>? floatingLineWidths = null;
            List<double>? floatingLineGaps = null;
            HashSet<int>? floatingPageStarts = null;
            var originalLines = lines;
            var originalLineHeights = lineHeights;
            void RestoreUnobstructedWrapping() {
                lines = originalLines; lineHeights = originalLineHeights;
                floatingLineOffsets = null; floatingLineWidths = null; floatingLineGaps = null; floatingPageStarts = null;
            }
            if (HasFloatingTables) {
                floatingLineOffsets = new(); floatingLineWidths = new(); floatingLineGaps = new();
                floatingPageStarts = new();
                double simulatedTop = y - (y < frameStart - 0.001 ? spacingBefore : 0);
                double previousHeight = 0;
                bool onOriginalPage = true;
                int layoutLineIndex = -1;
                double lineBaseTop = simulatedTop;
                bool lineBaseOriginalPage = true;
                var wrapped = WrapRichRunsCoreWithFirstLineOrigin(rpb.Runs, textFrame.Width, size,
                    ChooseNormal(currentOpts.DefaultFont), leading, textFrame.FirstLineWidth,
                    textFrame.FirstLineX - textFrame.X, GetParagraphTabStopWidth(paragraphStyle), currentOpts,
                    paragraphStyle?.TabStops.ToArray(), (index, completedHeight, requiredHeight, minimumWidth) => {
                        if (index != layoutLineIndex) {
                            simulatedTop -= completedHeight - previousHeight;
                            previousHeight = completedHeight;
                            layoutLineIndex = index;
                            lineBaseTop = simulatedTop;
                            lineBaseOriginalPage = onOriginalPage;
                        }
                        simulatedTop = lineBaseTop;
                        onOriginalPage = lineBaseOriginalPage;
                        floatingPageStarts.Remove(index);
                        if (simulatedTop - requiredHeight < currentOpts.MarginBottom) { simulatedTop = frameStart; onOriginalPage = false; floatingPageStarts.Add(index); }
                        double left = index == 0 ? textFrame.FirstLineX : textFrame.X;
                        double availableWidth = index == 0 ? textFrame.FirstLineWidth : textFrame.Width;
                        var frame = onOriginalPage ? GetFloatingTextFrame(left, availableWidth, simulatedTop, requiredHeight, minimumWidth) : (X: left, Width: availableWidth, Gap: 0D);
                        if (simulatedTop - frame.Gap - requiredHeight < currentOpts.MarginBottom) {
                            frame = (left, availableWidth, 0D);
                            simulatedTop = frameStart;
                            onOriginalPage = false;
                            floatingPageStarts.Add(index);
                        }
                        if (floatingLineOffsets.Count <= index) {
                            floatingLineOffsets.Add(0); floatingLineWidths.Add(0); floatingLineGaps.Add(0);
                        }
                        floatingLineOffsets[index] = frame.X - textFrame.X;
                        floatingLineWidths[index] = frame.Width;
                        floatingLineGaps[index] = frame.Gap;
                        return (frame.Width, frame.X - textFrame.X, frame.Gap);
                    }, lineSpacing: paragraphStyle?.LineSpacing);
                lines = wrapped.Lines; lineHeights = wrapped.LineHeights;
                double actualHeight = (y < frameStart - 0.001 ? spacingBefore : 0) + lineHeights.Sum();
                double nextHeight = paragraphStyle?.KeepWithNext == true && nextBlock != null
                    ? MeasureKeepWithNextChainHeight(blockList, blockIndex + 1, currentOpts.MarginLeft, width, size, actualHeight + spacingAfter) : 0;
                bool simulatedPageBreak = floatingPageStarts.Count > 0;
                bool mustMove = paragraphStyle?.KeepTogether == true &&
                    (simulatedPageBreak || actualHeight > y - currentOpts.MarginBottom + 0.001);
                mustMove |= nextHeight > 0 && (simulatedPageBreak || actualHeight + spacingAfter + nextHeight > y - currentOpts.MarginBottom + 0.001) &&
                    originalLineHeights.Sum() + spacingAfter + nextHeight <= frameStart - currentOpts.MarginBottom + 0.001;
                if (mustMove) { NewPage(); RestoreUnobstructedWrapping(); }
            }

            int lineIndex = 0;
            bool firstSegment = true;
            bool firstLogicalLine = true;
            textFrame = GetParagraphTextFrame(paragraphStyle, currentOpts.MarginLeft, width);
            void NewParagraphPage() {
                if (activeColumnFlow?.Options.BalanceLastPage == true) {
                    PdfParagraphStyle continuationStyle = paragraphStyle?.Clone() ?? new PdfParagraphStyle();
                    if (!firstLogicalLine) { continuationStyle.SpacingBefore = 0D; continuationStyle.FirstLineIndent = 0D; }
                    var remainingRuns = firstLogicalLine ? rpb.Runs : BuildTextRunsFromWrappedLines(lines, lineIndex, lines.Count - lineIndex);
                    QueueColumnBalanceRemainder(new RichParagraphBlock(remainingRuns, rpb.Align, rpb.DefaultColor, continuationStyle), blockList, blockIndex);
                }
                NewPage();
                frameStart = GetCurrentFramePageStartY();
                textFrame = GetParagraphTextFrame(paragraphStyle, currentOpts.MarginLeft, width);
                if (activeColumnFlow != null && Math.Abs(wrappedWidth - textFrame.Width) > 0.001D) {
                    IReadOnlyList<PdfTextRun> remainingRuns = firstLogicalLine ? rpb.Runs :
                        BuildTextRunsFromWrappedLines(lines, lineIndex, lines.Count - lineIndex);
                    var wrapped = WrapRichRunsCoreWithFirstLineOrigin(remainingRuns, textFrame.Width, size, ChooseNormal(currentOpts.DefaultFont), leading,
                        firstLogicalLine ? textFrame.FirstLineWidth : textFrame.Width,
                        firstLogicalLine ? textFrame.FirstLineX - textFrame.X : 0D, GetParagraphTabStopWidth(paragraphStyle), currentOpts,
                        paragraphStyle?.TabStops.ToArray(), lineSpacing: paragraphStyle?.LineSpacing);
                    lines = originalLines = wrapped.Lines; lineHeights = originalLineHeights = wrapped.LineHeights;
                    lineIndex = 0; wrappedWidth = textFrame.Width;
                    floatingLineOffsets = null; floatingLineWidths = null; floatingLineGaps = null; floatingPageStarts = null;
                    return;
                }
                if (lineIndex == 0) { RestoreUnobstructedWrapping(); return; }
                // Widow/orphan control can carry a previously simulated float-side line forward.
                // Keep its line break, but remove the old page's exclusion offset and clearance.
                if (floatingLineGaps != null) {
                    for (int index = lineIndex; index < lineHeights.Count; index++) {
                        lineHeights[index] -= floatingLineGaps[index];
                        double lineWidth = 0;
                        for (int segmentIndex = 0; segmentIndex < lines[index].Count; segmentIndex++) {
                            var segment = lines[index][segmentIndex];
                            var adjusted = segment;
                            if (segment.LeadingTabStop != null) {
                                double spaceWidth = MeasureRichText(" ", segment.Font, segment.NamedFont, segment.FontSize, segment.Baseline, currentOpts, segment.FeatureSettings, segment.HorizontalTextScaling, segment.CharacterSpacing);
                                adjusted = segment.WithLeadingAdvance(CalculateTabAdvance(lineWidth, segment.MeasuredWidth, spaceWidth,
                                    segment.LeadingTabStop.Alignment, GetParagraphTabStopWidth(paragraphStyle), segment.Text, segment.Font,
                                    segment.FontSize, segment.Baseline, currentOpts, textFrame.Width, segment.LeadingTabStop,
                                    followingNamedFont: segment.NamedFont, featureSettings: segment.FeatureSettings));
                                lines[index][segmentIndex] = adjusted;
                            }
                            lineWidth += adjusted.MeasuredWidth + adjusted.LeadingAdvance;
                        }
                    }
                }
                floatingLineOffsets = null; floatingLineWidths = null; floatingLineGaps = null; floatingPageStarts = null;
            }
            while (lineIndex < lines.Count) {
                if (floatingPageStarts?.Remove(lineIndex) == true && (HasFloatingTables || y < frameStart - 0.001)) NewParagraphPage();
                double minimumLineHeight = lineHeights[lineIndex];
                if (lineIndex == lines.Count - 1) minimumLineHeight += GetClosingTextPadding(spacingAfter);
                if (minimumLineHeight > GetMaximumBlockContinuationHeight() + .001D)
                    throw new ArgumentException("Paragraph line height exceeds the available page content height.");
                while (ShouldAdvanceForBlockHeight(minimumLineHeight)) {
                    NewParagraphPage();
                    minimumLineHeight = lineHeights[lineIndex] + (lineIndex == lines.Count - 1 ? GetClosingTextPadding(spacingAfter) : 0D);
                    if (minimumLineHeight > GetMaximumBlockContinuationHeight() + .001D)
                        throw new ArgumentException("Paragraph line height exceeds the available page content height.");
                }
                double available = y - currentOpts.MarginBottom;
                if (available <= 0) {
                    NewParagraphPage();
                    firstSegment = false;
                    continue;
                }

                double segmentSpacingBefore = firstSegment && y < frameStart - 0.001 ? spacingBefore : 0;
                if (available < segmentSpacingBefore + minimumLineHeight) {
                    NewParagraphPage();
                    minimumLineHeight = lineHeights[lineIndex] + (lineIndex == lines.Count - 1 ? GetClosingTextPadding(spacingAfter) : 0D);
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
                    double closingPadding = k == lines.Count - 1 ? GetClosingTextPadding(spacingAfter) : 0D;
                    if (heightSum + lineHeight + closingPadding > roomForText) {
                        break;
                    }

                    heightSum += lineHeight;
                    take++;
                }

                ReserveBalancedKeepNextTail(lineHeights, lineIndex, lines.Count, spacingAfter, roomForText, blockList,
                    blockIndex, paragraphStyle?.KeepWithNext == true, paragraphStyle == null ? 1 : ResolveMinimumWidowLines(paragraphStyle),
                    ref take, ref heightSum);
                if (TryApplyWidowControl(paragraphStyle, lines.Count, lineIndex, ref take, ref heightSum, lineHeights, y < frameStart - 0.001)) {
                    NewParagraphPage();
                    firstSegment = false;
                    continue;
                }

                if (take == 0) {
                    if (!ShouldAdvanceForBlockHeight(minimumLineHeight))
                        throw new ArgumentException("Paragraph line height exceeds the available page content height.");
                    NewParagraphPage();
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

                bool sliceStartsAtFirstLine = firstLogicalLine;
                double firstLineTop = y - (floatingLineGaps?[lineIndex] ?? 0D);
                RecordFlowPlacement(firstLineTop);
                if (sliceStartsAtFirstLine && paragraphStyle?.AnchoredCanvas is { } paragraphCanvas) RenderParagraphCanvas(paragraphCanvas, firstLineTop);
                pageDirty = true;
                var paragraphFont = ChooseNormal(currentOpts.DefaultFont);
                int? markedContentId = RegisterTextStructureElement("P");
                MarkRichFonts(rpb.Runs);
                WriteRichParagraph(sb, rpb, sliceLines, sliceHeights, currentOpts, FirstTextBaselineFromTop(paragraphFont, size, y), size, leading, currentPage!.Annotations, textFrame.X, textFrame.Width, sliceStartsAtFirstLine ? textFrame.FirstLineX : null, sliceStartsAtFirstLine ? textFrame.FirstLineWidth : null, "P", markedContentId, currentPage,
                    lineXOffsets: floatingLineOffsets?.GetRange(lineIndex, take), lineWidths: floatingLineWidths?.GetRange(lineIndex, take), lineTopGaps: floatingLineGaps?.GetRange(lineIndex, take));
                y -= heightSum;
                lineIndex += take;
                firstSegment = false;
                firstLogicalLine = false;
                if (lineIndex < lines.Count) {
                    NewParagraphPage();
                } else {
                    y -= spacingAfter;
                }
            }

            MarkRichFonts(rpb.Runs);
        }

    }
}
