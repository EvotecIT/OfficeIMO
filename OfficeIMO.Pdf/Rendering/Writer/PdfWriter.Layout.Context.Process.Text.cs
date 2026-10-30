using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private void RenderRichParagraphFlowBlock(RichParagraphBlock rpb, IPdfBlock? nextBlock, System.Collections.Generic.IList<IPdfBlock> blockList, int blockIndex) {
            double frameStart = GetCurrentFramePageStartY();
            PdfParagraphStyle? paragraphStyle = EffectiveParagraphStyle(rpb);
            double size = paragraphStyle?.FontSize ?? currentOpts.DefaultFontSize.ToPoints();
            double leading = GetParagraphLeading(paragraphStyle, size);
            double spacingBefore = GetParagraphSpacingBefore(paragraphStyle);
            double spacingAfter = GetParagraphSpacingAfter(paragraphStyle, leading);
            var textFrame = GetParagraphTextFrame(paragraphStyle, currentOpts.MarginLeft.ToPoints(), width);
            double wrappedWidth = textFrame.Width;
            RunningFrameAdmissionAttempt? runningAdmission = null;
            bool pictureWrapping = HasParagraphPictures(paragraphStyle) || HasPictureExclusions;
            // Keep, widow and frame admission must see the same physical rows that will paint.
            var (lines, lineHeights) = pictureWrapping
                ? WrapPictureParagraph(rpb, currentOpts.MarginLeft.ToPoints(), width, size,
                    y - ResolveTopLevelSpacingBefore(spacingBefore))
                : WrapRichRunsCoreWithFirstLineOrigin(rpb.Runs, textFrame.Width, size, ChooseNormal(currentOpts.DefaultFont), leading, textFrame.FirstLineWidth, textFrame.FirstLineX - textFrame.X, GetParagraphTabStopWidth(paragraphStyle), currentOpts, paragraphStyle?.TabStops.ToArray(), lineSpacing: paragraphStyle?.LineSpacing, textWhitespaceMode: paragraphStyle?.TextWhitespaceMode, wordJustificationMode: paragraphStyle?.UseWordJustificationSpacing == true && rpb.Align == PdfAlign.Justify ? paragraphStyle.JustificationMode : null);
            bool splittingConstrainedParagraph = ResolveParagraphConstraintOverflow(paragraphStyle, lines, lineHeights, leading) !=
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
                pictureWrapping |= HasPictureExclusions;
                if (pictureWrapping) {
                    // A page move changes paragraph anchors even when the frame width is unchanged.
                    (lines, lineHeights) = WrapPictureParagraph(rpb, currentOpts.MarginLeft.ToPoints(), width, size,
                        y - ResolveTopLevelSpacingBefore(spacingBefore));
                    wrappedWidth = textFrame.Width;
                } else if (Math.Abs(wrappedWidth - textFrame.Width) > .001D) {
                    (lines, lineHeights) = WrapRichRunsCoreWithFirstLineOrigin(rpb.Runs, textFrame.Width, size, ChooseNormal(currentOpts.DefaultFont), leading,
                        textFrame.FirstLineWidth, textFrame.FirstLineX - textFrame.X, GetParagraphTabStopWidth(paragraphStyle), currentOpts,
                        paragraphStyle?.TabStops.ToArray(), lineSpacing: paragraphStyle?.LineSpacing, textWhitespaceMode: paragraphStyle?.TextWhitespaceMode, wordJustificationMode: paragraphStyle?.UseWordJustificationSpacing == true && rpb.Align == PdfAlign.Justify ? paragraphStyle.JustificationMode : null);
                    wrappedWidth = textFrame.Width;
                }
            }
            bool IsKeptInCurrentFrame() {
                if (paragraphStyle?.KeepTogether != true) return false;
                if (lines.Count == 1 && (CanOverflowImportedExactLine(paragraphStyle) ||
                    CanOverflowImportedPictureLine(lines[0]))) return false;
                if (paragraphStyle.KeepTogetherOverflow != PdfKeepTogetherOverflowBehavior.Reject &&
                    MeasureRichLinesFrameFitHeight(lines, lineHeights, leading) > GetMaximumBlockContinuationHeight() + .001D) {
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
                    double paragraphContentHeight = MeasureRichLinesFrameFitHeight(lines, lineHeights, leading);
                    if (paragraphContentHeight > GetMaximumBlockContinuationHeight() + 0.001) {
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

            List<double>? floatingLineOffsets = null;
            textFrame = GetParagraphTextFrame(paragraphStyle, currentOpts.MarginLeft.ToPoints(), width);
            if (!pictureWrapping && activeColumnFlow != null && Math.Abs(wrappedWidth - textFrame.Width) > 0.001D) {
                (lines, lineHeights) = WrapRichRunsCoreWithFirstLineOrigin(rpb.Runs, textFrame.Width, size, ChooseNormal(currentOpts.DefaultFont), leading,
                    textFrame.FirstLineWidth, textFrame.FirstLineX - textFrame.X, GetParagraphTabStopWidth(paragraphStyle), currentOpts,
                    paragraphStyle?.TabStops.ToArray(), lineSpacing: paragraphStyle?.LineSpacing, textWhitespaceMode: paragraphStyle?.TextWhitespaceMode, wordJustificationMode: paragraphStyle?.UseWordJustificationSpacing == true && rpb.Align == PdfAlign.Justify ? paragraphStyle.JustificationMode : null);
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
            if (pictureWrapping) {
                floatingLineGaps = lines.Select(line => ((RichLine)line).FragmentTopGap).ToList();
            } else if (HasFloatingTables) {
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
                        if (simulatedTop - requiredHeight < FootnoteBodyBottom) { simulatedTop = frameStart; onOriginalPage = false; floatingPageStarts.Add(index); }
                        double left = index == 0 ? textFrame.FirstLineX : textFrame.X;
                        double availableWidth = index == 0 ? textFrame.FirstLineWidth : textFrame.Width;
                        var frame = onOriginalPage ? GetFloatingTextFrame(left, availableWidth, simulatedTop, requiredHeight, minimumWidth) : (X: left, Width: availableWidth, Gap: 0D);
                        if (simulatedTop - frame.Gap - requiredHeight < FootnoteBodyBottom) {
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
                    }, lineSpacing: paragraphStyle?.LineSpacing, textWhitespaceMode: paragraphStyle?.TextWhitespaceMode, wordJustificationMode: paragraphStyle?.UseWordJustificationSpacing == true && rpb.Align == PdfAlign.Justify ? paragraphStyle.JustificationMode : null);
                lines = wrapped.Lines; lineHeights = wrapped.LineHeights;
                double actualHeight = (y < frameStart - 0.001 ? spacingBefore : 0) + lineHeights.Sum();
                double nextHeight = paragraphStyle?.KeepWithNext == true && nextBlock != null
                    ? MeasureKeepWithNextChainHeight(blockList, blockIndex + 1, currentOpts.MarginLeft.ToPoints(), width, size, actualHeight + spacingAfter) : 0;
                bool simulatedPageBreak = floatingPageStarts.Count > 0;
                double fitHeight = (y < frameStart - 0.001 ? spacingBefore : 0D) +
                    MeasureRichLinesFrameFitHeight(lines, lineHeights, leading, floatingLineGaps);
                bool mustMove = IsKeptInCurrentFrame() &&
                    (simulatedPageBreak || fitHeight > y - FootnoteBodyBottom + 0.001);
                mustMove |= nextHeight > 0 && (simulatedPageBreak || actualHeight + spacingAfter + nextHeight > y - FootnoteBodyBottom + 0.001) &&
                    originalLineHeights.Sum() + spacingAfter + nextHeight <= frameStart - FootnoteBodyBottom + 0.001;
                if (mustMove) { NewPage(); RestoreUnobstructedWrapping(); }
            }

            int lineIndex = 0;
            bool firstSegment = true;
            bool firstLogicalLine = true;
            textFrame = GetParagraphTextFrame(paragraphStyle, currentOpts.MarginLeft.ToPoints(), width);
            void NewParagraphPage(bool preserveEmptyPage = false) {
                double firstBodyHeight = GetRichLineFrameFitHeight(lines[lineIndex], lineHeights[lineIndex], floatingLineGaps?[lineIndex] ?? 0D);
                if (lineIndex == lines.Count - 1) firstBodyHeight += GetClosingTextPadding(spacingAfter);
                firstBodyHeight += GetRunningBodyContinuationReservation();
                bool forcePhysicalPage = SkipColumnsForParagraphOverflow() || AdvancePastPartialColumn(firstBodyHeight);
                if (activeColumnFlow?.Options.BalanceLastPage == true) {
                    PdfParagraphStyle continuationStyle = paragraphStyle?.Clone() ?? new PdfParagraphStyle();
                    if (!firstLogicalLine) { continuationStyle.SpacingBefore = 0D; continuationStyle.FirstLineIndent = 0D; }
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
                pictureWrapping |= HasPictureExclusions;
                if (pictureWrapping) {
                    IReadOnlyList<PdfTextRun> remaining = firstLogicalLine ? rpb.Runs :
                        BuildTextRunsFromWrappedLines(lines, lineIndex, lines.Count - lineIndex);
                    PdfParagraphStyle? continuation = paragraphStyle?.Clone();
                    if (!firstLogicalLine) {
                        continuation ??= new PdfParagraphStyle();
                        continuation.SpacingBefore = 0D;
                        continuation.FirstLineIndent = 0D;
                        continuation.AnchoredCanvas = null;
                    }
                    var remainder = new RichParagraphBlock(remaining, rpb.Align, rpb.DefaultColor, continuation);
                    var wrapped = WrapPictureParagraph(remainder, currentOpts.MarginLeft.ToPoints(), width, size, y);
                    lines = wrapped.Lines;
                    lineHeights = wrapped.LineHeights;
                    floatingLineOffsets = null;
                    floatingLineWidths = null;
                    floatingLineGaps = lines.Select(line => ((RichLine)line).FragmentTopGap).ToList();
                    floatingPageStarts = null;
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
                        paragraphStyle?.TabStops.ToArray(), lineSpacing: paragraphStyle?.LineSpacing, textWhitespaceMode: paragraphStyle?.TextWhitespaceMode, wordJustificationMode: paragraphStyle?.UseWordJustificationSpacing == true && rpb.Align == PdfAlign.Justify ? paragraphStyle.JustificationMode : null);
                    lines = originalLines = wrapped.Lines; lineHeights = originalLineHeights = wrapped.LineHeights;
                    lineIndex = 0; wrappedWidth = textFrame.Width;
                    floatingLineOffsets = null; floatingLineWidths = null; floatingLineGaps = null; floatingPageStarts = null;
                    return;
                }
                if (lineIndex == 0) { RestoreUnobstructedWrapping(); return; }
                // Widow/orphan control can carry a previously simulated float-side line forward.
                // Keep its line break, but remove the old page's exclusion offset and clearance.
                if (floatingLineGaps != null) {
                    PdfTabStop[]? destinationStops = NormalizeExplicitTabStops(paragraphStyle?.TabStops.ToArray(), textFrame.Width);
                    for (int index = lineIndex; index < lineHeights.Count; index++) {
                        lineHeights[index] -= floatingLineGaps[index];
                        if (lines[index] is RichLine barLine) barLine.BarOriginOffset = 0D;
                        bool changedLeaderInk = false;
                        double lineWidth = 0;
                        int nextDestinationStop = 0;
                        for (int segmentIndex = 0; segmentIndex < lines[index].Count; segmentIndex++) {
                            var segment = lines[index][segmentIndex];
                            var adjusted = segment;
                            if (segment.LeadingIsTab || segment.LeadingTabStop != null) {
                                PdfTextRun? sourceTab = segment.SourceTabRun;
                                PdfTabStop? stop = sourceTab?.RelativeTabStop is { } relative
                                    ? new PdfTabStop(Math.Max(.001D, textFrame.Width * relative.Fraction + relative.Offset),
                                        sourceTab.TabAlignment, sourceTab.TabLeader)
                                    : TakeNextParagraphTabStop(destinationStops, lineWidth, ref nextDestinationStop);
                                // Older positioned segments may have no source-run
                                // provenance. Preserve their existing selected stop.
                                if (sourceTab == null) stop ??= segment.LeadingTabStop;
                                PdfTabAlignment alignment = stop?.Alignment ?? sourceTab?.TabAlignment ?? segment.LeadingTabAlignment;
                                PdfTabLeaderStyle leader = stop?.Leader ?? sourceTab?.TabLeader ?? segment.LeadingTabLeader;
                                double spaceWidth = MeasureRichText(" ", segment.Font, segment.NamedFont, segment.FontSize, segment.Baseline,
                                    currentOpts, segment.FeatureSettings, segment.HorizontalTextScaling, segment.CharacterSpacing,
                                    segment.TextDirection, segment.FontMetricScale);
                                RichTabFieldMeasurement? field = AlignsCompleteTabField(alignment)
                                    ? MeasureWrappedTabField(lines[index], segmentIndex) : null;
                                double fieldWidth = field?.Width ?? GetRichSegmentWidth(segment);
                                double advance = CalculateTabAdvance(lineWidth, fieldWidth, spaceWidth,
                                    alignment, GetParagraphTabStopWidth(paragraphStyle), segment.Text, segment.Font,
                                    segment.FontSize, segment.Baseline, currentOpts, textFrame.Width, stop,
                                    followingNamedFont: segment.NamedFont, featureSettings: segment.FeatureSettings,
                                    horizontalTextScaling: segment.HorizontalTextScaling, characterSpacing: segment.CharacterSpacing,
                                    decimalAnchorWidthOverride: ResolveRichDecimalTabAnchor(field, alignment, stop, currentOpts));
                                adjusted = segment.WithTabPlacement(advance, stop, alignment, leader);
                                changedLeaderInk |= HasPaintedRichTabLeader(segment, currentOpts) !=
                                    HasPaintedRichTabLeader(adjusted, currentOpts);
                                lines[index][segmentIndex] = adjusted;
                            }
                            lineWidth += GetRichSegmentWidth(adjusted) + adjusted.LeadingAdvance;
                        }
                        if (paragraphStyle?.UseWordJustificationSpacing == true && rpb.Align == PdfAlign.Justify)
                            BindWordCompressionPlan(lines[index], textFrame.Width);
                        if (changedLeaderInk && lines[index] is RichLine richLine)
                            lineHeights[index] = RefreshRichLineMetrics(richLine, ChooseNormal(currentOpts.DefaultFont),
                                size, leading, currentOpts, paragraphStyle?.LineSpacing);
                    }
                }
                floatingLineOffsets = null; floatingLineWidths = null; floatingLineGaps = null; floatingPageStarts = null;
            }
            while (lineIndex < lines.Count) {
                if (floatingPageStarts?.Remove(lineIndex) == true && (HasFloatingTables || y < frameStart - 0.001)) NewParagraphPage();
                double minimumLineHeight = GetRichLineFrameFitHeight(lines[lineIndex], lineHeights[lineIndex], floatingLineGaps?[lineIndex] ?? 0D);
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
                    minimumLineHeight = GetRichLineFrameFitHeight(lines[lineIndex], lineHeights[lineIndex], floatingLineGaps?[lineIndex] ?? 0D) + (lineIndex == lines.Count - 1 ? GetFootnoteClosingTextPadding(lines, lineIndex, spacingAfter) : 0D);
                    while (minimumLineHeight > GetMaximumParagraphContinuationHeight() + .001D &&
                        !CanOverflowImportedExactLine(paragraphStyle) &&
                        !CanOverflowImportedPictureLine(lines[lineIndex]) &&
                        !CanOverlapRunningFooterWithFirstLine(paragraphStyle, minimumLineHeight)) {
                        if (!TryAdvanceForRunningContent(ref runningAdmission, minimumLineHeight,
                            preserve => NewParagraphPage(preserve), GetMaximumParagraphContinuationHeight()))
                            throw new ArgumentException("Paragraph line height exceeds the available page content height.");
                        minimumLineHeight = GetRichLineFrameFitHeight(lines[lineIndex], lineHeights[lineIndex], floatingLineGaps?[lineIndex] ?? 0D) +
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
                    minimumLineHeight = GetRichLineFrameFitHeight(lines[lineIndex], lineHeights[lineIndex], floatingLineGaps?[lineIndex] ?? 0D) + (lineIndex == lines.Count - 1 ? GetFootnoteClosingTextPadding(lines, lineIndex, spacingAfter) : 0D);
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
                    if (heightSum + GetRichLineFrameFitHeight(lines[k], lineHeight, floatingLineGaps?[k] ?? 0D) + closingPadding > roomForText) {
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
                int? markedContentId = RegisterTextStructureElement("P");
                MarkRichFonts(rpb.Runs);
                ReportRichPaintDiagnostics(PaintDiagnosticCodes.Paragraph, "PdfParagraph", rpb, sliceLines, sliceHeights,
                    FirstTextBaselineFromTop(paragraphFont, size, y), size, leading, textFrame.X, textFrame.Width,
                    new TableCellContentFrame(textFrame.X, currentOpts.PageHeight.ToPoints(), textFrame.Width, currentOpts.PageHeight.ToPoints()), clipsToFrame: false,
                    offsets: floatingLineOffsets?.GetRange(lineIndex, take), widths: floatingLineWidths?.GetRange(lineIndex, take),
                    firstLineX: sliceStartsAtFirstLine ? textFrame.FirstLineX : null,
                    firstLineWidth: sliceStartsAtFirstLine ? textFrame.FirstLineWidth : null,
                    topGaps: floatingLineGaps?.GetRange(lineIndex, take),
                    hasFollowingParagraphLines: lineIndex + take < lines.Count);
                WriteRichParagraph(sb, rpb, sliceLines, sliceHeights, currentOpts, FirstTextBaselineFromTop(paragraphFont, size, y), size, leading, currentPage!.Annotations, textFrame.X, textFrame.Width, sliceStartsAtFirstLine ? textFrame.FirstLineX : null, sliceStartsAtFirstLine ? textFrame.FirstLineWidth : null, "P", markedContentId, currentPage,
                    lineXOffsets: floatingLineOffsets?.GetRange(lineIndex, take), lineWidths: floatingLineWidths?.GetRange(lineIndex, take), lineTopGaps: floatingLineGaps?.GetRange(lineIndex, take),
                    hasFollowingParagraphLines: lineIndex + take < lines.Count,
                    paragraphSpacingAfter: lineIndex + take == lines.Count ? spacingAfter : 0D,
                    paragraphSpacingBefore: segmentSpacingBefore);
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
