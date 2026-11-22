using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private void RenderRichParagraphFlowBlock(RichParagraphBlock rpb, IPdfBlock? nextBlock, System.Collections.Generic.IList<IPdfBlock> blockList, int blockIndex) {
            double frameStart = GetCurrentFramePageStartY();
            PdfParagraphStyle? paragraphStyle = EffectiveParagraphStyle(rpb);
            double size = paragraphStyle?.FontSize?.ToPoints() ?? currentOpts.DefaultFontSize.ToPoints();
            double leading = GetParagraphLeading(paragraphStyle, size);
            double spacingBefore = GetParagraphSpacingBefore(paragraphStyle);
            double spacingAfter = GetParagraphSpacingAfter(paragraphStyle, leading);
            var textFrame = GetParagraphTextFrame(paragraphStyle, currentOpts.MarginLeft.ToPoints(), width);
            double wrappedWidth = textFrame.Width;
            RunningFrameAdmissionAttempt? runningAdmission = null;
            bool pictureWrapping = HasParagraphPictures(paragraphStyle) || HasTextExclusions;
            // Beside a floating table, a line that would pass the page bottom starts the next page.
            HashSet<int>? exclusionPageStarts = null;
            HashSet<int>? TrackExclusionPageStarts() => exclusionPageStarts = HasFloatingTables ? new HashSet<int>() : null;
            (List<List<RichSeg>> Lines, List<double> LineHeights) WrapUnobstructedParagraph() =>
                WrapRichRunsCoreWithFirstLineOrigin(rpb.Runs, textFrame.Width, size, ChooseNormal(currentOpts.DefaultFont), leading,
                    textFrame.FirstLineWidth, textFrame.FirstLineX - textFrame.X, GetParagraphTabStopWidth(paragraphStyle), currentOpts,
                    paragraphStyle?.TabStops.ToArray(), lineSpacing: paragraphStyle?.LineSpacing, textWhitespaceMode: paragraphStyle?.TextWhitespaceMode,
                    wordJustificationMode: paragraphStyle?.UseWordJustificationSpacing == true && rpb.Align == OfficeAlignment.Justify ? paragraphStyle.JustificationMode : null,
                    cancellationToken: cancellationToken);
            // Keep, widow and frame admission must see the same physical rows that will paint.
            var (lines, lineHeights) = pictureWrapping
                ? WrapPictureParagraph(rpb, currentOpts.MarginLeft.ToPoints(), width, size,
                    y - ResolveTopLevelSpacingBefore(spacingBefore), TrackExclusionPageStarts())
                : WrapUnobstructedParagraph();
            bool splittingConstrainedParagraph = ResolveParagraphConstraintOverflow(paragraphStyle, lines, lineHeights, leading, GetClosingContainerSpace()) !=
                PdfKeepTogetherOverflowBehavior.Reject;
            bool SkipColumnsForParagraphOverflow() => splittingConstrainedParagraph &&
                paragraphStyle?.KeepTogetherOverflow == PdfKeepTogetherOverflowBehavior.SplitAcrossPages;
            // Every column shares this partial physical frame. Ordinary used
            // columns still advance in reading order when just their tail fills.
            bool AdvancePastPartialColumn(double height) => (CanOverflowImportedExactLine(paragraphStyle) ||
                lines.Any(HasImportedOverflowPicture)) &&
                activeColumnFlow is { } column && IsPartialPhysicalColumnFrame() &&
                height > column.Top - column.ParentOptions.MarginBottom.ToPoints() + .001D;
            if (activeColumnFlow is { Options.BalanceParagraphLines: false, Options.BalanceLastPage: true } &&
                lineHeights.Sum() <= GetMaximumBlockContinuationHeight() + 0.001D) {
                paragraphStyle = paragraphStyle?.Clone() ?? new PdfParagraphStyle();
                paragraphStyle.KeepTogether = true;
            }
            void NewParagraphStartFrame(bool forcePhysicalPage = false, bool preserveEmptyPage = false) {
                NewBlockFrame(forcePhysicalPage: forcePhysicalPage || SkipColumnsForParagraphOverflow(), preserveEmptyPage: preserveEmptyPage);
                frameStart = GetCurrentFramePageStartY();
                textFrame = GetParagraphTextFrame(paragraphStyle, currentOpts.MarginLeft.ToPoints(), width);
                pictureWrapping |= HasTextExclusions;
                if (pictureWrapping) {
                    // A page move changes paragraph anchors even when the frame width is unchanged.
                    (lines, lineHeights) = WrapPictureParagraph(rpb, currentOpts.MarginLeft.ToPoints(), width, size,
                        y - ResolveTopLevelSpacingBefore(spacingBefore), TrackExclusionPageStarts());
                    wrappedWidth = textFrame.Width;
                } else if (Math.Abs(wrappedWidth - textFrame.Width) > .001D) {
                    (lines, lineHeights) = WrapUnobstructedParagraph();
                    wrappedWidth = textFrame.Width;
                }
            }
            bool IsKeptInCurrentFrame() {
                if (paragraphStyle?.KeepTogether != true) return false;
                if (lines.Count == 1 && (CanOverflowImportedExactLine(paragraphStyle) ||
                    CanOverflowImportedPictureLine(lines[0]))) return false;
                if (paragraphStyle.KeepTogetherOverflow != PdfKeepTogetherOverflowBehavior.Reject &&
                    MeasureTextBlockAdmissionHeight(lines, lineHeights, leading, spacingAfter) > GetMaximumBlockContinuationHeight() + .001D) {
                    splittingConstrainedParagraph = true;
                    return false;
                }
                return !(activeColumnFlow is { Options.BalanceKeptParagraphLines: true } && IsBalancedColumnFrame() &&
                    lineHeights.Sum() > GetCurrentFramePageStartY() - FootnoteBodyBottom + .001D);
            }
            if (paragraphStyle?.KeepWithNext == true && nextBlock != null && lines.Count > 0 &&
                (ReservesWholeKeepGroup() || IsKeptInCurrentFrame())) {
                double paragraphHeight = ResolveTopLevelSpacingBefore(spacingBefore) + lineHeights.Sum() + spacingAfter;
                double nextHeight = MeasureCurrentFrameKeepNextHeight(blockList, blockIndex + 1, currentOpts.MarginLeft.ToPoints(), width, size, paragraphHeight);
                double keepHeight = paragraphHeight + nextHeight;
                while ((ReservesWholeKeepGroup() || IsKeptInCurrentFrame()) && nextHeight > 0.001 &&
                    keepHeight <= GetMaximumBlockContinuationHeight() + 0.001 && ShouldAdvanceForBlockHeight(keepHeight)) {
                    NewParagraphStartFrame();
                    paragraphHeight = lineHeights.Sum() + spacingAfter;
                    nextHeight = MeasureCurrentFrameKeepNextHeight(blockList, blockIndex + 1, currentOpts.MarginLeft.ToPoints(), width, size, paragraphHeight);
                    keepHeight = paragraphHeight + nextHeight;
                }
            }

            if (IsKeptInCurrentFrame()) {
                while (true) {
                    if (!IsKeptInCurrentFrame()) break;
                    List<double>? exclusionGaps = pictureWrapping ? lines.Select(line => ((RichLine)line).FragmentTopGap).ToList() : null;
                    double paragraphContentHeight = MeasureTextBlockAdmissionHeight(lines, lineHeights, leading, spacingAfter, exclusionGaps);
                    double maximumHeight = GetMaximumBlockContinuationHeight();
                    if (paragraphContentHeight - (exclusionGaps?.Sum() ?? 0D) > maximumHeight + 0.001) {
                        // A table on this page can add lines as well as clearance. Only reject a paragraph that
                        // is still too tall at the full width of the continuation frame.
                        if (HasFloatingTables) {
                            var unobstructed = WrapUnobstructedParagraph();
                            if (MeasureTextBlockAdmissionHeight(unobstructed.Lines, unobstructed.LineHeights, leading, spacingAfter) <= maximumHeight + 0.001) {
                                NewParagraphStartFrame(forcePhysicalPage: true);
                                continue;
                            }
                        }
                        if (TryAdvanceForRunningContent(ref runningAdmission, paragraphContentHeight,
                            preserve => NewParagraphStartFrame(preserveEmptyPage: preserve))) continue;
                        throw new ArgumentException("Paragraph height exceeds the available page content height.");
                    }
                    double paragraphHeight = ResolveTopLevelSpacingBefore(spacingBefore) + paragraphContentHeight;
                    if (!ShouldAdvanceForBlockHeight(paragraphHeight)) break;
                    NewParagraphStartFrame();
                }
            }
            if (splittingConstrainedParagraph && (y < frameStart - .001D || IsPartialPhysicalColumnFrame()))
                NewParagraphStartFrame(forcePhysicalPage: IsPartialPhysicalColumnFrame());

            textFrame = GetParagraphTextFrame(paragraphStyle, currentOpts.MarginLeft.ToPoints(), width);
            if (!pictureWrapping && activeColumnFlow != null && Math.Abs(wrappedWidth - textFrame.Width) > 0.001D) {
                (lines, lineHeights) = WrapUnobstructedParagraph();
                wrappedWidth = textFrame.Width;
                frameStart = GetCurrentFramePageStartY();
            }
            List<double>? floatingLineGaps = null;
            HashSet<int>? floatingPageStarts = null;
            if (pictureWrapping) {
                floatingLineGaps = lines.Select(line => ((RichLine)line).FragmentTopGap).ToList();
                if (exclusionPageStarts != null) {
                    floatingPageStarts = new HashSet<int>(exclusionPageStarts);
                    // A kept paragraph, or one that keeps with the next block, moves to the next
                    // page when a floating table pushes part of it there.
                    double spacing = y < frameStart - 0.001 ? spacingBefore : 0D;
                    double actualHeight = spacing + lineHeights.Sum();
                    double nextHeight = paragraphStyle?.KeepWithNext == true && nextBlock != null
                        ? MeasureKeepWithNextChainHeight(blockList, blockIndex + 1, currentOpts.MarginLeft.ToPoints(), width, size, actualHeight + spacingAfter) : 0;
                    bool simulatedPageBreak = floatingPageStarts.Count > 0;
                    double fitHeight = spacing + MeasureTextBlockAdmissionHeight(lines, lineHeights, leading, spacingAfter, floatingLineGaps);
                    double unobstructedHeight = lineHeights.Sum() - floatingLineGaps.Sum();
                    bool mustMove = IsKeptInCurrentFrame() && (simulatedPageBreak || fitHeight > y - FootnoteBodyBottom + 0.001);
                    mustMove |= nextHeight > 0 && (simulatedPageBreak || actualHeight + spacingAfter + nextHeight > y - FootnoteBodyBottom + 0.001) &&
                        unobstructedHeight + spacingAfter + nextHeight <= frameStart - FootnoteBodyBottom + 0.001;
                    if (mustMove) {
                        NewPage();
                        frameStart = GetCurrentFramePageStartY();
                        (lines, lineHeights) = WrapPictureParagraph(rpb, currentOpts.MarginLeft.ToPoints(), width, size, y, TrackExclusionPageStarts());
                        floatingLineGaps = lines.Select(line => ((RichLine)line).FragmentTopGap).ToList();
                        floatingPageStarts = exclusionPageStarts == null ? null : new HashSet<int>(exclusionPageStarts);
                    }
                }
            }

            int lineIndex = 0;
            bool firstSegment = true;
            bool firstLogicalLine = true;
            textFrame = GetParagraphTextFrame(paragraphStyle, currentOpts.MarginLeft.ToPoints(), width);
            void NewParagraphPage(bool preserveEmptyPage = false) {
                double firstBodyHeight = GetTextLineFitHeight(lines[lineIndex], lineHeights[lineIndex], lineIndex == lines.Count - 1, floatingLineGaps?[lineIndex] ?? 0D);
                if (lineIndex == lines.Count - 1) firstBodyHeight += GetClosingTextPadding(spacingAfter);
                firstBodyHeight += GetRunningBodyContinuationReservation();
                bool forcePhysicalPage = SkipColumnsForParagraphOverflow() || AdvancePastPartialColumn(firstBodyHeight);
                if (activeColumnFlow?.Options.BalanceLastPage == true) {
                    PdfParagraphStyle continuationStyle = paragraphStyle?.Clone() ?? new PdfParagraphStyle();
                    if (!firstLogicalLine) { continuationStyle.SpacingBefore = global::OfficeIMO.OfficeLength.Points(0D); continuationStyle.FirstLineIndent = global::OfficeIMO.OfficeLength.Points(0D); }
                    var remainingRuns = firstLogicalLine ? rpb.Runs : BuildTextRunsFromWrappedLines(lines, lineIndex, lines.Count - lineIndex);
                    QueueColumnBalanceRemainder(new RichParagraphBlock(remainingRuns, rpb.Align, rpb.DefaultColor, continuationStyle), blockList, blockIndex,
                        forcePhysicalPage);
                }
                NewPage(preserveEmptyPage: preserveEmptyPage, forcePhysicalPage: forcePhysicalPage, firstBodyHeight: firstBodyHeight,
                    firstBodyAllowsFooterOverlap: paragraphStyle?.AllowFirstLineOverlapWithRunningFooter == true && activeColumnFlow == null,
                    firstBodyAllowsExactOverflow: CanOverflowImportedExactLine(paragraphStyle) ||
                        CanOverflowImportedPictureLine(lines[lineIndex]));
                frameStart = GetCurrentFramePageStartY();
                textFrame = GetParagraphTextFrame(paragraphStyle, currentOpts.MarginLeft.ToPoints(), width);
                pictureWrapping |= HasTextExclusions;
                if (pictureWrapping) {
                    IReadOnlyList<PdfTextRun> remaining = firstLogicalLine ? rpb.Runs :
                        BuildTextRunsFromWrappedLines(lines, lineIndex, lines.Count - lineIndex);
                    PdfParagraphStyle? continuation = paragraphStyle?.Clone();
                    if (!firstLogicalLine) {
                        continuation ??= new PdfParagraphStyle();
                        continuation.SpacingBefore = global::OfficeIMO.OfficeLength.Points(0D);
                        continuation.FirstLineIndent = global::OfficeIMO.OfficeLength.Points(0D);
                        continuation.AnchoredCanvas = null;
                    }
                    var remainder = new RichParagraphBlock(remaining, rpb.Align, rpb.DefaultColor, continuation);
                    var wrapped = WrapPictureParagraph(remainder, currentOpts.MarginLeft.ToPoints(), width, size, y, TrackExclusionPageStarts());
                    lines = wrapped.Lines;
                    lineHeights = wrapped.LineHeights;
                    floatingLineGaps = lines.Select(line => ((RichLine)line).FragmentTopGap).ToList();
                    floatingPageStarts = exclusionPageStarts == null ? null : new HashSet<int>(exclusionPageStarts);
                    lineIndex = 0;
                    wrappedWidth = textFrame.Width;
                    return;
                }
                if (activeColumnFlow != null && Math.Abs(wrappedWidth - textFrame.Width) > 0.001D) {
                    IReadOnlyList<PdfTextRun> remainingRuns = firstLogicalLine ? rpb.Runs :
                        BuildTextRunsFromWrappedLines(lines, lineIndex, lines.Count - lineIndex);
                    var wrapped = WrapRichRunsCoreWithFirstLineOrigin(remainingRuns, textFrame.Width, size, ChooseNormal(currentOpts.DefaultFont), leading,
                        firstLogicalLine ? textFrame.FirstLineWidth : textFrame.Width,
                        firstLogicalLine ? textFrame.FirstLineX - textFrame.X : 0D, GetParagraphTabStopWidth(paragraphStyle), currentOpts,
                        paragraphStyle?.TabStops.ToArray(), lineSpacing: paragraphStyle?.LineSpacing, textWhitespaceMode: paragraphStyle?.TextWhitespaceMode, wordJustificationMode: paragraphStyle?.UseWordJustificationSpacing == true && rpb.Align == OfficeAlignment.Justify ? paragraphStyle.JustificationMode : null, cancellationToken: cancellationToken);
                    lines = wrapped.Lines; lineHeights = wrapped.LineHeights;
                    lineIndex = 0; wrappedWidth = textFrame.Width;
                    return;
                }
            }
            while (lineIndex < lines.Count) {
                if (floatingPageStarts?.Remove(lineIndex) == true && (HasFloatingTables || y < frameStart - 0.001)) NewParagraphPage();
                double minimumLineHeight = GetTextLineFitHeight(lines[lineIndex], lineHeights[lineIndex], lineIndex == lines.Count - 1, floatingLineGaps?[lineIndex] ?? 0D);
                if (lineIndex == lines.Count - 1) minimumLineHeight += GetFootnoteClosingTextPadding(lines, lineIndex, spacingAfter);
                if (minimumLineHeight > GetMaximumParagraphContinuationHeight() + .001D &&
                    !CanOverflowImportedExactLine(paragraphStyle) &&
                    !CanOverflowImportedPictureLine(lines[lineIndex]) &&
                    !CanOverlapRunningFooterWithFirstLine(paragraphStyle, minimumLineHeight)) {
                    if (TryAdvanceForRunningContent(ref runningAdmission, minimumLineHeight,
                        preserve => NewParagraphPage(preserve), GetMaximumParagraphContinuationHeight())) continue;
                    throw new ArgumentException("Paragraph line height exceeds the available page content height.");
                }
                while (ShouldAdvanceForBlockHeight(minimumLineHeight) || AdvancePastPartialColumn(minimumLineHeight)) {
                    NewParagraphPage();
                    minimumLineHeight = GetTextLineFitHeight(lines[lineIndex], lineHeights[lineIndex], lineIndex == lines.Count - 1, floatingLineGaps?[lineIndex] ?? 0D) + (lineIndex == lines.Count - 1 ? GetFootnoteClosingTextPadding(lines, lineIndex, spacingAfter) : 0D);
                    while (minimumLineHeight > GetMaximumParagraphContinuationHeight() + .001D &&
                        !CanOverflowImportedExactLine(paragraphStyle) &&
                        !CanOverflowImportedPictureLine(lines[lineIndex]) &&
                        !CanOverlapRunningFooterWithFirstLine(paragraphStyle, minimumLineHeight)) {
                        if (!TryAdvanceForRunningContent(ref runningAdmission, minimumLineHeight,
                            preserve => NewParagraphPage(preserve), GetMaximumParagraphContinuationHeight()))
                            throw new ArgumentException("Paragraph line height exceeds the available page content height.");
                        minimumLineHeight = GetTextLineFitHeight(lines[lineIndex], lineHeights[lineIndex], lineIndex == lines.Count - 1, floatingLineGaps?[lineIndex] ?? 0D) +
                            (lineIndex == lines.Count - 1 ? GetFootnoteClosingTextPadding(lines, lineIndex, spacingAfter) : 0D);
                    }
                }
                double available = y - FootnoteBodyBottom;
                if (available <= 0) {
                    NewParagraphPage();
                    firstSegment = false;
                    continue;
                }

                double segmentSpacingBefore = firstSegment && y < frameStart - 0.001 ? spacingBefore : 0;
                if (available < segmentSpacingBefore + minimumLineHeight &&
                    !(y >= frameStart - .001D && (CanOverflowImportedExactLine(paragraphStyle) ||
                        CanOverflowImportedPictureLine(lines[lineIndex]) ||
                        CanOverlapRunningFooterWithFirstLine(paragraphStyle, minimumLineHeight)))) {
                    NewParagraphPage();
                    minimumLineHeight = GetTextLineFitHeight(lines[lineIndex], lineHeights[lineIndex], lineIndex == lines.Count - 1, floatingLineGaps?[lineIndex] ?? 0D) + (lineIndex == lines.Count - 1 ? GetFootnoteClosingTextPadding(lines, lineIndex, spacingAfter) : 0D);
                    available = y - FootnoteBodyBottom;
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
                    double closingPadding = k == lines.Count - 1 ? GetFootnoteClosingTextPadding(lines, lineIndex, spacingAfter) : 0D;
                    if (heightSum + GetTextLineFitHeight(lines[k], lineHeight, k == lines.Count - 1, floatingLineGaps?[k] ?? 0D) + closingPadding > roomForText) {
                        break;
                    }

                    heightSum += lineHeight;
                    take++;
                }

                if (take == 0 && y >= frameStart - .001D &&
                    (CanOverflowImportedExactLine(paragraphStyle) || CanOverflowImportedPictureLine(lines[lineIndex]) ||
                        CanOverlapRunningFooterWithFirstLine(paragraphStyle, minimumLineHeight))) {
                    take = 1;
                    heightSum = lineHeights[lineIndex];
                }
                bool noteSliceFits = ConstrainFootnoteSlice(lines, lineHeights, lineIndex,
                    IsKeptInCurrentFrame() ? lines.Count - lineIndex : paragraphStyle == null ? 1 : ResolveMinimumOrphanLines(paragraphStyle),
                    paragraphStyle == null ? 0 : ResolveMinimumWidowLines(paragraphStyle), segmentSpacingBefore,
                    GetFootnoteClosingTextPadding(lines, lineIndex, spacingAfter), ref take, ref heightSum);
                if (!noteSliceFits && !HasCurrentPageBodyContent())
                    throw new ArgumentException("The paragraph reference and first footnote fragment exceed the available page body height.");
                if (!noteSliceFits) { NewParagraphPage(); firstSegment = false; continue; }
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

                if (activeColumnFlow is { } column && (CanOverflowImportedExactLine(paragraphStyle) ||
                    lines.Skip(lineIndex).Take(take).Any(HasImportedOverflowPicture))) {
                    double physicalHeight = column.Top - column.ParentOptions.MarginBottom.ToPoints();
                    for (int index = lineIndex; index < lineIndex + take; index++) {
                        if (lineHeights[index] > physicalHeight + .001D) {
                            column.RequiresPhysicalContinuation = true;
                            break;
                        }
                    }
                }

                if (segmentSpacingBefore > 0) {
                    y -= segmentSpacingBefore;
                }
                bool admittedNewFootnote = HasUnscheduledFootnote(lines.Skip(lineIndex).Take(take).SelectMany(line => line));
                RequireFootnoteSlice(lines, lineIndex, take, y - heightSum - (lineIndex + take == lines.Count ? GetFootnoteClosingTextPadding(lines, lineIndex, spacingAfter) : 0D));

                var sliceLines = new System.Collections.Generic.List<System.Collections.Generic.List<RichSeg>>();
                var sliceHeights = new System.Collections.Generic.List<double>();
                for (int k = 0; k < take; k++) {
                    sliceLines.Add(lines[lineIndex + k]);
                    sliceHeights.Add(lineHeights[lineIndex + k]);
                }

                bool sliceStartsAtFirstLine = firstLogicalLine;
                double firstLineTop = y - (floatingLineGaps?[lineIndex] ?? 0D);
                RecordFlowPlacement(firstLineTop);
                if (sliceStartsAtFirstLine && paragraphStyle?.AnchoredCanvas is { } paragraphCanvas)
                    RenderPictureParagraphCanvas(paragraphCanvas, y, firstLineTop);
                pageDirty = true;
                var paragraphFont = ChooseNormal(currentOpts.DefaultFont);
                int? markedContentId = RegisterTextStructureElement(rpb.StructureType);
                MarkRichFonts(rpb.Runs);
                ReportRichPaintDiagnostics(PaintDiagnosticCodes.Paragraph, "PdfParagraph", rpb, sliceLines, sliceHeights,
                    FirstTextBaselineFromTop(paragraphFont, size, y), size, leading, textFrame.X, textFrame.Width,
                    new TableCellContentFrame(textFrame.X, currentOpts.PageHeight.ToPoints(), textFrame.Width, currentOpts.PageHeight.ToPoints()), clipsToFrame: false,
                    firstLineX: sliceStartsAtFirstLine ? textFrame.FirstLineX : null,
                    firstLineWidth: sliceStartsAtFirstLine ? textFrame.FirstLineWidth : null,
                    topGaps: floatingLineGaps?.GetRange(lineIndex, take),
                    hasFollowingParagraphLines: lineIndex + take < lines.Count);
                WriteRichParagraph(sb, rpb, sliceLines, sliceHeights, currentOpts, FirstTextBaselineFromTop(paragraphFont, size, y), size, leading, currentPage!.Annotations, textFrame.X, textFrame.Width, sliceStartsAtFirstLine ? textFrame.FirstLineX : null, sliceStartsAtFirstLine ? textFrame.FirstLineWidth : null, rpb.StructureType, markedContentId, currentPage,
                    lineTopGaps: floatingLineGaps?.GetRange(lineIndex, take),
                    hasFollowingParagraphLines: lineIndex + take < lines.Count,
                    paragraphSpacingAfter: lineIndex + take == lines.Count ? spacingAfter : 0D,
                    paragraphSpacingBefore: segmentSpacingBefore, suppressInlineAccessibility: _suppressCanvasAccessibilityWrappers);
                y -= heightSum;
                lineIndex += take;
                runningAdmission = null;
                firstSegment = false;
                firstLogicalLine = false;
                if (lineIndex < lines.Count) {
                    if (!admittedNewFootnote || y - GetRichLineFrameFitHeight(lines[lineIndex], lineHeights[lineIndex]) < FootnoteBodyBottom - .001D) NewParagraphPage();
                } else {
                    y -= spacingAfter;
                }
            }

            MarkRichFonts(rpb.Runs);
        }

    }
}
