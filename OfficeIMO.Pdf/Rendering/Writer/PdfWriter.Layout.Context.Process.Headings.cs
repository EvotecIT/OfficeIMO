using System.Globalization;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private void RenderHeadingFlowBlock(HeadingBlock hb, IPdfBlock? nextBlock, System.Collections.Generic.IList<IPdfBlock> blockList, int blockIndex, Action<double>? captureFirstPlacement = null) {
            double frameStart = GetCurrentFramePageStartY();
            PdfHeadingStyle? headingStyle = ResolveHeadingStyle(hb, currentOpts);
            double size = GetHeadingFontSize(hb, headingStyle);
            double leading = GetHeadingLeading(headingStyle, size);
            double spacingBefore = (y < frameStart - 0.001 || headingStyle?.ApplySpacingBeforeAtTop == true) ? (headingStyle?.SpacingBefore.ToPoints()) ?? 0D : 0D;
            double spacingAfter = GetHeadingSpacingAfter(headingStyle, leading);
            var headingFont = GetHeadingFont(currentOpts, headingStyle);
            global::OfficeIMO.OfficeColor? headingColor = hb.Color ?? headingStyle?.Color;
            System.Collections.Generic.IReadOnlyList<PdfTextRun> headingRuns = CreateHeadingTextRuns(hb, headingStyle, headingColor);
            bool pictureWrapping = HasAnchoredPictures(headingStyle?.AnchoredCanvas) || HasPictureExclusions;
            var (lines, lineHeights) = WrapHeadingForFrame(hb, headingRuns, headingStyle,
                currentOpts.MarginLeft.ToPoints(), width, size, leading, y - spacingBefore);
            var lineGaps = lines.Select(line => line is RichLine row ? row.FragmentTopGap : 0D).ToList();
            double wrappedWidth = width;
            int lineIndex = 0;
            bool firstLogicalLine = true;
            double appliedSpacingBefore = 0D;
            RunningFrameAdmissionAttempt? runningAdmission = null;
            double textHeight = MeasureRichLinesHeight(lineHeights, lines.Count, leading);
            void NewHeadingFrame(bool preserveEmptyPage = false) {
                if (firstLogicalLine) QueueColumnBalanceRemainder(hb, blockList, blockIndex);
                else QueueColumnBalanceRemainder(new List<ColumnBalanceUnit> {
                    new(lineHeights.Skip(lineIndex).Sum() + spacingAfter)
                }, hb, blockList, blockIndex);
                NewPage(preserveEmptyPage: preserveEmptyPage);
                appliedSpacingBefore = 0D;
                frameStart = GetCurrentFramePageStartY();
                pictureWrapping |= HasPictureExclusions;
                if (pictureWrapping || activeColumnFlow != null && Math.Abs(wrappedWidth - width) > 0.001D) {
                    var remainingRuns = firstLogicalLine ? headingRuns :
                        BuildTextRunsFromWrappedLines(lines, lineIndex, lines.Count - lineIndex);
                    PdfHeadingStyle? remainingStyle = headingStyle;
                    HeadingBlock source = hb;
                    if (!firstLogicalLine) {
                        remainingStyle = headingStyle?.Clone();
                        if (remainingStyle != null) remainingStyle.AnchoredCanvas = null;
                        source = new HeadingBlock(hb.Level, hb.Text, hb.Align, hb.Color,
                            hb.LinkUri, remainingStyle, hb.LinkContents, hb.LinkDestinationName, remainingRuns);
                    }
                    double before = firstLogicalLine && headingStyle?.ApplySpacingBeforeAtTop == true
                        ? headingStyle.SpacingBefore.ToPoints() : 0D;
                    (lines, lineHeights) = WrapHeadingForFrame(source, remainingRuns, remainingStyle,
                        currentOpts.MarginLeft.ToPoints(), width, size, leading, y - before);
                    lineGaps = lines.Select(line => line is RichLine row ? row.FragmentTopGap : 0D).ToList();
                    lineIndex = 0;
                    wrappedWidth = width;
                    textHeight = MeasureRichLinesHeight(lineHeights, lines.Count, leading);
                }
            }
            double needed = spacingBefore + textHeight + spacingAfter;
            bool keepWithNext = headingStyle?.KeepWithNext ?? true;
            if (keepWithNext && nextBlock != null) {
                double keepHeight = needed + MeasureCurrentFrameKeepNextHeight(blockList, blockIndex + 1, currentOpts.MarginLeft.ToPoints(), width, currentOpts.DefaultFontSize.ToPoints(), needed);
                double availableHeight = GetMaximumBlockContinuationHeight();
                while (keepHeight > needed + 0.001 && keepHeight <= availableHeight + 0.001 && ShouldAdvanceForBlockHeight(keepHeight)) {
                    NewHeadingFrame();
                    spacingBefore = headingStyle?.ApplySpacingBeforeAtTop == true ? headingStyle.SpacingBefore.ToPoints() : 0D;
                    needed = spacingBefore + textHeight + spacingAfter;
                    keepHeight = needed + MeasureCurrentFrameKeepNextHeight(blockList, blockIndex + 1, currentOpts.MarginLeft.ToPoints(), width, currentOpts.DefaultFontSize.ToPoints(), needed);
                    availableHeight = GetMaximumBlockContinuationHeight();
                }
            }

            double fitNeeded = spacingBefore + MeasureRichLinesFrameFitHeight(lines, lineHeights, leading, lineGaps) + spacingAfter;
            if (y < frameStart - 0.001 && y - fitNeeded < FootnoteBodyBottom) {
                NewHeadingFrame();
                spacingBefore = headingStyle?.ApplySpacingBeforeAtTop == true ? headingStyle.SpacingBefore.ToPoints() : 0D;
                needed = spacingBefore + textHeight + spacingAfter;
            }
            while (lineHeights.Count > 0 && ShouldAdvanceForBlockHeight(spacingBefore + GetRichLineFrameFitHeight(lines[0], lineHeights[0], lineGaps[0]))) {
                NewHeadingFrame();
                spacingBefore = headingStyle?.ApplySpacingBeforeAtTop == true ? headingStyle.SpacingBefore.ToPoints() : 0D;
            }
            while (lineHeights.Count > 0 && spacingBefore + GetRichLineFrameFitHeight(lines[0], lineHeights[0], lineGaps[0]) > frameStart - FootnoteBodyBottom + 0.001) {
                if (!TryAdvanceForRunningContent(ref runningAdmission,
                    spacingBefore + GetRichLineFrameFitHeight(lines[0], lineHeights[0], lineGaps[0]),
                    preserve => NewHeadingFrame(preserve)))
                    throw new ArgumentException("Heading spacing and first line exceed the available page content height.");
                spacingBefore = headingStyle?.ApplySpacingBeforeAtTop == true ? headingStyle.SpacingBefore.ToPoints() : 0D;
            }
            if (spacingBefore > 0) {
                y -= spacingBefore;
                appliedSpacingBefore = spacingBefore;
            }

            PageStructElement? logicalHeading = null;
            while (lineIndex < lines.Count) {
                double available = y - FootnoteBodyBottom;
                int take = 0;
                double heightSum = 0;
                for (int index = lineIndex; index < lines.Count; index++) {
                    double lineHeight = lineHeights[index];
                    double closingPadding = index == lines.Count - 1 ? GetFootnoteClosingTextPadding(lines, lineIndex, spacingAfter) : 0D;
                    if (heightSum + GetRichLineFrameFitHeight(lines[index], lineHeight, lineGaps[index]) + closingPadding > available + 0.001) break;
                    heightSum += lineHeight;
                    take++;
                }
                bool noteSliceFits = ConstrainFootnoteSlice(lines, lineHeights, lineIndex, 1, 0, 0D, GetFootnoteClosingTextPadding(lines, lineIndex, spacingAfter), ref take, ref heightSum);
                if (!noteSliceFits && !HasCurrentPageBodyContent())
                    throw new ArgumentException("The heading reference and first footnote fragment exceed the available page body height.");
                if (!noteSliceFits) { NewHeadingFrame(); continue; }
                if (take == 0) {
                    double minimumHeight = GetRichLineFrameFitHeight(lines[lineIndex], lineHeights[lineIndex], lineGaps[lineIndex]) +
                        (lineIndex == lines.Count - 1 ? GetFootnoteClosingTextPadding(lines, lineIndex, spacingAfter) : 0D);
                    if (!ShouldAdvanceForBlockHeight(minimumHeight)) {
                        if (TryAdvanceForRunningContent(ref runningAdmission, minimumHeight,
                            preserve => NewHeadingFrame(preserve))) continue;
                        throw new ArgumentException("Heading line height exceeds the available page content height.");
                    }
                    NewHeadingFrame();
                    continue;
                }
                var sliceLines = lines.GetRange(lineIndex, take);
                var sliceHeights = lineHeights.GetRange(lineIndex, take);
                bool admittedNewFootnote = HasUnscheduledFootnote(sliceLines.SelectMany(line => line));
                RequireFootnoteSlice(lines, lineIndex, take, y - heightSum - (lineIndex + take == lines.Count ? GetFootnoteClosingTextPadding(lines, lineIndex, spacingAfter) : 0D));
                EnsurePage();
                double firstTextTop = y - lineGaps[lineIndex];
                RecordFlowPlacement(firstTextTop);
                if (firstLogicalLine) captureFirstPlacement?.Invoke(firstTextTop);
                if (firstLogicalLine && headingStyle?.AnchoredCanvas is { } headingCanvas)
                    RenderPictureParagraphCanvas(headingCanvas, y, firstTextTop);
                pageDirty = true;
                if (firstLogicalLine && currentOpts.CreateOutlineFromHeadings) {
                    currentPage!.Bookmarks.Add(new PageBookmark { Level = hb.Level, Title = hb.Text, Y = firstTextTop });
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
                if (firstLogicalLine && currentOpts.CreateOutlineFromHeadings && currentPage!.Bookmarks.Count > 0)
                    currentPage.Bookmarks[currentPage.Bookmarks.Count - 1].StructureElement = logicalHeading ?? ResolveMarkedStructure(markedContentId);
                AddHeadingLinkAnnotations(hb, sliceLines, headingFont, size, leading, currentOpts.MarginLeft.ToPoints(), width, firstBaseline, linkStructElementIndex, sliceHeights);
                WriteRichParagraph(sb, new RichParagraphBlock(headingRuns, hb.Align, headingColor), sliceLines, sliceHeights, currentOpts, firstBaseline, size, leading, currentPage!.Annotations, currentOpts.MarginLeft.ToPoints(), width, structureType: markedStructureType, markedContentId: markedContentId, structurePage: currentPage, baselineFont: headingFont,
                    paragraphSpacingAfter: lineIndex + take == lines.Count ? spacingAfter : 0D,
                    paragraphSpacingBefore: firstLogicalLine ? appliedSpacingBefore : 0D,
                    lineTopGaps: lineGaps.GetRange(lineIndex, take));
                MarkRichFonts(headingRuns);
                if (GetHeadingBold(headingStyle)) {
                    currentPage!.UsedBold = true;
                    usedBold = true;
                }
                y -= heightSum;
                lineIndex += take;
                runningAdmission = null;
                firstLogicalLine = false;
                if (lineIndex < lines.Count && (!admittedNewFootnote || y - GetRichLineFrameFitHeight(lines[lineIndex], lineHeights[lineIndex], lineGaps[lineIndex]) < FootnoteBodyBottom - .001D)) NewHeadingFrame();
            }
            y -= spacingAfter;
        }

    }
}
