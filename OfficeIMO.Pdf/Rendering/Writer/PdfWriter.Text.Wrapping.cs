using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private static (System.Collections.Generic.List<System.Collections.Generic.List<RichSeg>> Lines, System.Collections.Generic.List<double> LineHeights) WrapRichRuns(System.Collections.Generic.IEnumerable<PdfTextRun> runs, double maxWidthPts, double fontSize, PdfStandardFont baseFont, double lineHeight, double? firstLineWidthPts = null, double tabStopWidth = DefaultParagraphTabStopWidth) =>
        WrapRichRunsCore(runs, maxWidthPts, fontSize, baseFont, lineHeight, firstLineWidthPts, tabStopWidth, options: null);

    private static PdfTabStop[]? NormalizeExplicitTabStops(System.Collections.Generic.IReadOnlyList<PdfTabStop>? tabStops, double frameWidth) {
        if (tabStops == null || tabStops.Count == 0) {
            return null;
        }

        return tabStops
            .Where(tabStop => !double.IsNaN(tabStop.Position) && !double.IsInfinity(tabStop.Position))
            .Select(tabStop => tabStop.ResolveFramePosition(frameWidth))
            .OrderBy(tabStop => tabStop.Position)
            .ToArray();
    }

    private static (System.Collections.Generic.List<System.Collections.Generic.List<RichSeg>> Lines, System.Collections.Generic.List<double> LineHeights) WrapRichRunsCore(System.Collections.Generic.IEnumerable<PdfTextRun> runs, double maxWidthPts, double fontSize, PdfStandardFont baseFont, double lineHeight, double? firstLineWidthPts, double tabStopWidth, PdfOptions? options, System.Collections.Generic.IReadOnlyList<PdfTabStop>? tabStops = null) {
        return WrapRichRunsCoreWithFirstLineOrigin(runs, maxWidthPts, fontSize, baseFont, lineHeight, firstLineWidthPts, null, tabStopWidth, options, tabStops);
    }

    private static (System.Collections.Generic.List<System.Collections.Generic.List<RichSeg>> Lines, System.Collections.Generic.List<double> LineHeights) WrapRichRunsWithSpacing(System.Collections.Generic.IEnumerable<PdfTextRun> runs, double maxWidthPts, double fontSize, PdfStandardFont baseFont, double lineHeight, double? firstLineWidthPts, double tabStopWidth, PdfOptions? options, PdfLineSpacing? lineSpacing, PdfJustificationMode? wordJustificationMode = null) =>
        WrapRichRunsCoreWithFirstLineOrigin(runs, maxWidthPts, fontSize, baseFont, lineHeight, firstLineWidthPts, null, tabStopWidth, options, lineSpacing: lineSpacing, wordJustificationMode: wordJustificationMode);

    private static (System.Collections.Generic.List<System.Collections.Generic.List<RichSeg>> Lines, System.Collections.Generic.List<double> LineHeights) WrapRichRunsCoreWithFirstLineOrigin(System.Collections.Generic.IEnumerable<PdfTextRun> runs, double maxWidthPts, double fontSize, PdfStandardFont baseFont, double lineHeight, double? firstLineWidthPts, double? firstLineOriginOffsetPts, double tabStopWidth, PdfOptions? options, System.Collections.Generic.IReadOnlyList<PdfTabStop>? tabStops = null, Func<int, double, double, double, (double Width, double Origin, double Gap)>? lineLayout = null, double? minimumLineHeight = null, PdfLineSpacing? lineSpacing = null, PdfTextWhitespaceMode? textWhitespaceMode = null, RichWrapPurpose purpose = RichWrapPurpose.Layout, double? inlineRuleWidth = null, double? firstLineInlineRuleWidth = null, PdfJustificationMode? wordJustificationMode = null) {
        PdfTextWhitespaceMode whitespaceMode = textWhitespaceMode ?? options?.TextWhitespaceMode ?? PdfTextWhitespaceMode.Collapse;
        bool compressWordSpaces = wordJustificationMode.HasValue && purpose == RichWrapPurpose.Layout;
        var admittedLineWidths = new List<double>();
        bool preserveWhitespace = whitespaceMode != PdfTextWhitespaceMode.Collapse;
        bool preformattedWhitespace = whitespaceMode == PdfTextWhitespaceMode.Preformatted;
        System.Collections.Generic.IReadOnlyList<PdfTextRun> effectiveRuns = NormalizeFallbackRuns(runs, baseFont, options);
        var runKerning = new RichRunKerning(options);
        PdfTextRun? currentTextRun = null;
        var wordContinuations = MeasureRichWordContinuations(effectiveRuns, baseFont, fontSize, options);
        PdfTabStop[]? explicitTabStops = NormalizeExplicitTabStops(tabStops, maxWidthPts);
        // A field selected at a left stop can wrap and select an earlier non-left stop.
        bool mayReselectNonLeftTabStop = explicitTabStops?.Any(stop => AlignsCompleteTabField(stop.Alignment)) == true;
        var lines = new System.Collections.Generic.List<System.Collections.Generic.List<RichSeg>> { new RichLine() };
        var heights = new System.Collections.Generic.List<double>();
        double lineWidth = 0;
        double pendingLeadingAdvance = 0;
        int pendingLeadingSpaceCount = 0;
        var pendingSpaceFragments = new List<PendingSpaceFragment>();
        List<(List<RichSeg> Line, RichSeg Anchor, RichSeg? SpaceStyle)>? invisibleFootnoteAnchors = null;
        RichSeg? currentSpaceStyle = null;
        RichSeg? pendingTabSourceStyle = null;
        int pendingTabSourceLine = -1;
        bool pendingLeadingSeparator = false;
        bool pendingLeadingIsOrdinarySpace = false;
        bool oversizedWordContinuation = false;
        OfficeIMO.Drawing.OfficeTextDecorationStyle pendingLeadingUnderlineStyle = OfficeIMO.Drawing.OfficeTextDecorationStyle.None;
        PdfColor? pendingLeadingDecorationColor = null;
        double pendingLeadingDecorationFontSize = 0;
        double pendingLeadingDecorationTextRise = 0;
        bool pendingLeadingIsExpandable = true;
        bool pendingLeadingIsTab = false;
        PdfTabAlignment pendingLeadingTabAlignment = PdfTabAlignment.Left;
        PdfTabLeaderStyle pendingLeadingTabLeader = PdfTabLeaderStyle.None;
        PdfTabStop? pendingLeadingTabStop = null;
        PdfTextRun? pendingRelativeTabRun = null;
        int nextExplicitTabStopIndex = 0;
        double lineHeightRatio = fontSize > 0 ? lineHeight / fontSize : 1.2D;
        bool currentRunContributesDescent = true;
        double RunLineHeight(double size) => currentRunContributesDescent
            ? lineSpacing?.GetAdvance(size) ?? size * lineHeightRatio : 0D;
        double completedHeight = 0;
        double currentContentHeight = lineHeight;
        double currentMinimumWidth = 0;
        PdfNamedFontFace? currentRunNamedFont = null;
        double currentRunAscent = 0;
        double currentRunDescent = 0;
        double currentLineTextAscent = 0;
        double currentLineBodyTextAscent = 0;
        double currentLineTextDescent = 0;
        double currentLineInlineAscent = 0;
        double currentLineInlineDescent = 0;
        bool currentRunIsInline = false;
        bool currentLineHasInline = false;
        PdfColor? currentRunDecorationColor = null;
        OfficeTextFeatureSettings currentRunFeatureSettings = OfficeTextFeatureSettings.Default;
        double currentRunHorizontalTextScaling = 100D, currentRunCharacterSpacing = 0D;
        OfficeIMO.Drawing.OfficeTextDecorationStyle currentRunUnderlineStyle = OfficeIMO.Drawing.OfficeTextDecorationStyle.None;
        PdfColor? currentRunUnderlineColor = null;
        double currentRunDecorationFontSize = 0;
        double currentRunDecorationTextRise = 0;
        var currentFrame = lineLayout?.Invoke(0, completedHeight, CurrentFrameFitHeight(lineHeight, includeCurrentRun: false), 0);
        double currentLineHeight = Math.Max(lineHeight, minimumLineHeight ?? 0) + (currentFrame?.Gap ?? 0);
        double CurrentMaxWidth() => currentFrame?.Width ?? (lines.Count == 1 ? firstLineWidthPts ?? maxWidthPts : maxWidthPts);
        double CurrentLineOriginOffset() => currentFrame?.Origin ?? (lines.Count == 1 ? firstLineOriginOffsetPts ?? 0D : 0D);
        double MinimumLineWidth(IReadOnlyList<RichSeg> candidate, double naturalWidth) => compressWordSpaces
            ? MeasureWordCompressionMinimumWidth(candidate, naturalWidth) : naturalWidth;
        List<RichSeg> CompressionCandidate(RichSeg probe, IReadOnlyList<RichSeg>? continuation) {
            bool includeLeading = lines[lines.Count - 1].Count > 0 || preserveWhitespace || pendingLeadingIsTab;
            return WordCompressionCandidate(lines[lines.Count - 1],
                includeLeading ? pendingSpaceFragments : Array.Empty<PendingSpaceFragment>(),
                includeLeading ? pendingLeadingAdvance : 0D, includeLeading ? pendingLeadingSpaceCount : 0,
                pendingLeadingIsTab, pendingLeadingIsExpandable, includeLeading && pendingLeadingIsOrdinarySpace,
                pendingLeadingTabAlignment, probe, continuation);
        }
        double MinimumCandidateWidth(RichSeg probe, IReadOnlyList<RichSeg>? continuation, double naturalWidth) =>
            compressWordSpaces ? MeasureWordCompressionMinimumWidth(CompressionCandidate(probe, continuation), naturalWidth) : naturalWidth;

        void RegisterLineHeight(double runFontSize) {
            currentLineHeight = Math.Max(currentLineHeight, RunLineHeight(runFontSize) + (currentFrame?.Gap ?? 0));
            RegisterLineMetrics();
        }

        void RegisterLineMetrics() {
            if (lineSpacing?.IsExact == true) return;
            if (currentRunIsInline) {
                currentLineInlineAscent = Math.Max(currentLineInlineAscent, currentRunAscent);
                currentLineInlineDescent = Math.Max(currentLineInlineDescent, currentRunDescent);
            } else {
                currentLineTextAscent = Math.Max(currentLineTextAscent, currentRunAscent);
                currentLineTextDescent = Math.Max(currentLineTextDescent, currentRunDescent);
                if (currentRunContributesDescent)
                    currentLineBodyTextAscent = Math.Max(currentLineBodyTextAscent, currentRunAscent);
            }
            currentLineHasInline |= currentRunIsInline;
            currentLineHeight = Math.Max(currentLineHeight,
                CombinedContentHeight(0, includeCurrentRun: false) + (currentFrame?.Gap ?? 0));
        }

        void StartNewLine() {
            if (lines.Count >= MaximumTextLayoutLines)
                throw new System.IO.InvalidDataException("PDF paragraph layout exceeds the 100,000-line limit.");
            admittedLineWidths.Add(CurrentMaxWidth());
            heights.Add(currentLineHeight);
            completedHeight += currentLineHeight;
            lines.Add(new RichLine());
            lineWidth = 0;
            currentLineTextAscent = 0;
            currentLineBodyTextAscent = 0;
            currentLineTextDescent = 0;
            currentLineInlineAscent = 0;
            currentLineInlineDescent = 0;
            currentLineHasInline = false;
            currentContentHeight = lineHeight;
            currentFrame = lineLayout?.Invoke(lines.Count - 1, completedHeight, CurrentFrameFitHeight(currentContentHeight, includeCurrentRun: false), 0);
            currentLineHeight = Math.Max(currentContentHeight, minimumLineHeight ?? 0) + (currentFrame?.Gap ?? 0);
            currentMinimumWidth = 0;
            nextExplicitTabStopIndex = 0;
        }

        void PrepareLineFrame(double contentHeight, double minimumWidth = 0) {
            if (lineLayout == null) return;
            if (lineSpacing?.IsExact == true) contentHeight = lineHeight;
            double incomingContentHeight = contentHeight;
            contentHeight = CombinedContentHeight(contentHeight);
            currentContentHeight = Math.Max(lineHeight, contentHeight);
            double oldHeight = currentLineHeight - (currentFrame?.Gap ?? 0);
            var oldFrame = currentFrame;
            double height = lines[lines.Count - 1].Count == 0 ? currentContentHeight : Math.Max(oldHeight, contentHeight);
            double requiredWidth = Math.Max(currentMinimumWidth, minimumWidth);
            var frame = lineLayout(lines.Count - 1, completedHeight, CurrentFrameFitHeight(height), requiredWidth);
            if (lines[lines.Count - 1].Count > 0 &&
                (MinimumLineWidth(lines[lines.Count - 1], lineWidth) > frame.Width + 0.001 || minimumWidth > 0 && frame.Gap > (oldFrame?.Gap ?? 0) + 0.001)) {
                lineLayout(lines.Count - 1, completedHeight, CurrentFrameFitHeight(oldHeight, includeCurrentRun: false), currentMinimumWidth);
                currentFrame = oldFrame;
                StartNewLine();
                height = Math.Max(lineHeight, CombinedContentHeight(incomingContentHeight));
                requiredWidth = minimumWidth;
                frame = lineLayout(lines.Count - 1, completedHeight, CurrentFrameFitHeight(height), requiredWidth);
            }
            currentFrame = frame;
            currentMinimumWidth = requiredWidth;
            currentLineHeight = height + frame.Gap;
        }

        double CombinedContentHeight(double contentHeight, bool includeCurrentRun = true) {
            if (lineSpacing?.IsExact == true) return contentHeight;
            bool hasInline = currentLineHasInline || includeCurrentRun && currentRunIsInline;
            double textAscent = Math.Max(currentLineTextAscent,
                includeCurrentRun && !currentRunIsInline ? currentRunAscent : 0);
            double textDescent = Math.Max(currentLineTextDescent,
                includeCurrentRun && !currentRunIsInline ? currentRunDescent : 0);
            double bodyAscent = Math.Max(currentLineBodyTextAscent,
                includeCurrentRun && !currentRunIsInline && currentRunContributesDescent ? currentRunAscent : 0);
            double inlineAscent = Math.Max(currentLineInlineAscent,
                includeCurrentRun && currentRunIsInline ? currentRunAscent : 0);
            double inlineDescent = Math.Max(currentLineInlineDescent,
                includeCurrentRun && currentRunIsInline ? currentRunDescent : 0);
            double combined = Math.Max(textAscent, inlineAscent) + Math.Max(textDescent, inlineDescent);
            if (lineSpacing?.FontLineBoxMultiplier is not double natural)
                return hasInline ? Math.Max(contentHeight, combined) : contentHeight;

            if (lineSpacing.LineGridPitch.HasValue)
                return System.Math.Max(contentHeight, lineSpacing.GetGridAdvance(combined));

            // Imported spacing scales the visible fonts' natural advance.
            // Horizontal list spacers must not suppress that multiplier.
            double textAdvance = textAscent + textDescent;
            // Multiple spacing scales body metrics. A label's additional ascent
            // still enlarges the first line, without multiplying that excess.
            if (lineSpacing.Rule == PdfLineSpacingRule.Multiple)
                textAdvance = bodyAscent == 0D && textDescent == 0D ? 0D
                    : (bodyAscent + textDescent) * lineSpacing.Value / natural
                        + Math.Max(0D, textAscent - bodyAscent);
            contentHeight = Math.Max(contentHeight, textAdvance);
            // Inline objects retain their unscaled bounds when they extend
            // beyond the text line box; a spacer inside it adds no height.
            return hasInline && (inlineAscent > textAscent || inlineDescent > textDescent)
                ? Math.Max(contentHeight, combined) : contentHeight;
        }

        // Exclusions need the line's occupied box; completedHeight still carries
        // its full advance so subsequent baselines retain the authored spacing.
        double CurrentFrameFitHeight(double advance, bool includeCurrentRun = true) {
            if (lineSpacing?.FontLineBoxMultiplier is not double natural || lineSpacing.Rule != PdfLineSpacingRule.Multiple)
                return advance;
            double textAscent = Math.Max(currentLineTextAscent, includeCurrentRun && !currentRunIsInline ? currentRunAscent : 0D);
            double bodyAscent = Math.Max(currentLineBodyTextAscent,
                includeCurrentRun && !currentRunIsInline && currentRunContributesDescent ? currentRunAscent : 0D);
            double textDescent = Math.Max(currentLineTextDescent, includeCurrentRun && !currentRunIsInline ? currentRunDescent : 0D);
            double inlineAscent = Math.Max(currentLineInlineAscent, includeCurrentRun && currentRunIsInline ? currentRunAscent : 0D);
            double inlineDescent = Math.Max(currentLineInlineDescent, includeCurrentRun && currentRunIsInline ? currentRunDescent : 0D);
            double ascent = Math.Max(textAscent, inlineAscent), descent = Math.Max(textDescent, inlineDescent);
            if (ascent == 0D && descent == 0D) {
                GetRichRunLineMetrics(baseFont, null, fontSize, PdfTextBaseline.Normal, options, lineSpacing, out ascent, out descent);
                bodyAscent = ascent;
            }
            if (lineSpacing.LineGridPitch.HasValue)
                ascent += lineSpacing.GetGridBaselineExpansion(ascent + descent);
            else ascent = Math.Max(inlineAscent, ascent - bodyAscent * (1D - Math.Min(1D, lineSpacing.Value / natural)));
            return Math.Min(advance, Math.Max(0D, ascent + descent));
        }

        PdfTabStop? ResolveNextExplicitTabStop() {
            if (explicitTabStops == null || explicitTabStops.Length == 0) {
                return null;
            }

            while (nextExplicitTabStopIndex < explicitTabStops.Length &&
                   explicitTabStops[nextExplicitTabStopIndex].Position <= CurrentLineOriginOffset() + lineWidth + 0.001D) {
                nextExplicitTabStopIndex++;
            }

            if (nextExplicitTabStopIndex >= explicitTabStops.Length) {
                return null;
            }

            return explicitTabStops[nextExplicitTabStopIndex++];
        }

        PdfTabStop? ResolveRelativeTabStop(PdfTextRun? run) => run?.RelativeTabStop is { } relative
            ? new PdfTabStop(Math.Max(0.001D, maxWidthPts * relative.Fraction + relative.Offset), run.TabAlignment, run.TabLeader)
            : null;

        void ResolvePendingLeadingTabForCurrentLine(double followingTextWidth, double spaceW, string followingText, PdfStandardFont followingFont, double followingFontSize, PdfTextBaseline followingBaseline, RichTabFieldMeasurement? field = null) {
            PdfTabAlignment fallbackAlignment = pendingLeadingTabAlignment;
            PdfTabLeaderStyle fallbackLeader = pendingLeadingTabLeader;
            // Same-line inline measurement retains the selected stop. A new
            // line selects again from its own origin and available stops.
            PdfTabStop? explicitTabStop = pendingTabSourceLine == lines.Count - 1
                ? pendingLeadingTabStop : ResolveRelativeTabStop(pendingRelativeTabRun) ?? ResolveNextExplicitTabStop();
            pendingLeadingTabAlignment = explicitTabStop?.Alignment ?? fallbackAlignment;
            pendingLeadingTabLeader = explicitTabStop?.Leader ?? fallbackLeader;
            pendingLeadingTabStop = explicitTabStop;
            pendingTabSourceLine = lines.Count - 1;
            double selectedWidth = field != null && AlignsCompleteTabField(pendingLeadingTabAlignment)
                ? field.Width : followingTextWidth;
            pendingLeadingAdvance = CalculateTabAdvance(lineWidth, selectedWidth, spaceW, pendingLeadingTabAlignment, tabStopWidth, followingText, followingFont, followingFontSize, followingBaseline, options, CurrentMaxWidth(), pendingLeadingTabStop, CurrentLineOriginOffset(), currentRunNamedFont, currentRunFeatureSettings, currentRunHorizontalTextScaling, currentRunCharacterSpacing, decimalAnchorWidthOverride: ResolveRichDecimalTabAnchor(field, pendingLeadingTabAlignment, pendingLeadingTabStop, options));
        }

        void ResetPendingLeading() {
            pendingLeadingAdvance = 0;
            pendingLeadingSpaceCount = 0;
            pendingSpaceFragments.Clear();
            pendingTabSourceStyle = null;
            pendingTabSourceLine = -1;
            pendingLeadingSeparator = false;
            pendingLeadingIsOrdinarySpace = false;
            pendingLeadingUnderlineStyle = OfficeIMO.Drawing.OfficeTextDecorationStyle.None;
            pendingLeadingDecorationColor = null;
            pendingLeadingDecorationFontSize = 0;
            pendingLeadingDecorationTextRise = 0;
            pendingLeadingIsExpandable = true;
            pendingLeadingIsTab = false;
            pendingLeadingTabAlignment = PdfTabAlignment.Left;
            pendingLeadingTabLeader = PdfTabLeaderStyle.None;
            pendingLeadingTabStop = null;
            pendingRelativeTabRun = null;
        }

        void RevalidateLeadingTabFrame(double contentHeight, double followingWidth, double spaceWidth, string followingText, PdfStandardFont font, double size, PdfTextBaseline baseline, RichTabFieldMeasurement? field = null) {
            if (lineLayout == null || !pendingLeadingIsTab || lines[lines.Count - 1].Count != 0) return;
            double selectedWidth = field != null && AlignsCompleteTabField(pendingLeadingTabAlignment)
                ? field.Width : followingWidth;
            PrepareLineFrame(contentHeight, pendingLeadingAdvance + selectedWidth);
            // Retain the same-line stop, or resolve the new line if the frame refit moved it.
            ResolvePendingLeadingTabForCurrentLine(followingWidth, spaceWidth, followingText, font, size, baseline, field);
            selectedWidth = field != null && AlignsCompleteTabField(pendingLeadingTabAlignment)
                ? field.Width : followingWidth;
            PrepareLineFrame(contentHeight, pendingLeadingAdvance + selectedWidth);
        }

        void SetPendingSeparator(bool hadTab, double spaceW, PdfTabAlignment tabAlignment, PdfTabLeaderStyle tabLeader) {
            oversizedWordContinuation = false;
            pendingLeadingSeparator = true;
            pendingLeadingIsOrdinarySpace = !hadTab;
            if (!hadTab) {
                if (preserveWhitespace && pendingLeadingIsTab && pendingTabSourceStyle != null)
                    pendingSpaceFragments.Add(new PendingSpaceFragment(CreatePreservedTabFragment(
                        pendingTabSourceStyle, pendingLeadingAdvance, pendingLeadingTabStop,
                        pendingLeadingTabAlignment, pendingLeadingTabLeader)));
                pendingLeadingAdvance = preserveWhitespace ? pendingLeadingAdvance + spaceW : spaceW;
                pendingLeadingSpaceCount = preserveWhitespace ? pendingLeadingSpaceCount + 1 : 1;
                if (preserveWhitespace && currentSpaceStyle != null) {
                    if (pendingSpaceFragments.Count > 0 && ReferenceEquals(pendingSpaceFragments[pendingSpaceFragments.Count - 1].Style, currentSpaceStyle))
                        pendingSpaceFragments[pendingSpaceFragments.Count - 1].AddSpace();
                    else pendingSpaceFragments.Add(new PendingSpaceFragment(currentSpaceStyle));
                }
                pendingLeadingUnderlineStyle = currentRunUnderlineStyle;
                pendingLeadingDecorationColor = currentRunUnderlineColor;
                pendingLeadingDecorationFontSize = currentRunDecorationFontSize;
                pendingLeadingDecorationTextRise = currentRunDecorationTextRise;
                pendingLeadingIsExpandable = !preformattedWhitespace;
                pendingLeadingIsTab = false;
                pendingLeadingTabAlignment = PdfTabAlignment.Left;
                pendingLeadingTabLeader = PdfTabLeaderStyle.None;
                pendingLeadingTabStop = null;
                pendingRelativeTabRun = null;
                return;
            }

            // Commit the previous preserved separator before advancing to the next
            // tab stop. Otherwise consecutive tabs all resolve from the same x.
            if (preserveWhitespace && pendingLeadingAdvance > 0D && currentSpaceStyle != null) {
                if (pendingLeadingIsTab) {
                    lines[lines.Count - 1].Add(CreatePreservedTabFragment(currentSpaceStyle,
                        pendingLeadingAdvance, pendingLeadingTabStop, pendingLeadingTabAlignment, pendingLeadingTabLeader));
                } else {
                    lines[lines.Count - 1].AddRange(ConsumePreservedSpaceFragments(pendingSpaceFragments, pendingLeadingAdvance));
                }
                lineWidth += pendingLeadingAdvance;
            }
            pendingRelativeTabRun = currentTextRun?.RelativeTabStop.HasValue == true ? currentTextRun : null;
            PdfTabStop? explicitTabStop = ResolveRelativeTabStop(pendingRelativeTabRun) ?? ResolveNextExplicitTabStop();
            pendingSpaceFragments.Clear();
            pendingLeadingSpaceCount = 0;
            pendingLeadingTabAlignment = explicitTabStop?.Alignment ?? tabAlignment;
            pendingLeadingTabLeader = explicitTabStop?.Leader ?? tabLeader;
            pendingLeadingTabStop = explicitTabStop;
            pendingLeadingAdvance = CalculateTabAdvance(lineWidth, 0D, spaceW, pendingLeadingTabAlignment, tabStopWidth, options: options, maxWidth: CurrentMaxWidth(), explicitTabStop: pendingLeadingTabStop, lineOriginOffset: CurrentLineOriginOffset(), followingNamedFont: currentRunNamedFont, featureSettings: currentRunFeatureSettings, horizontalTextScaling: currentRunHorizontalTextScaling, characterSpacing: currentRunCharacterSpacing);
            pendingLeadingUnderlineStyle = OfficeIMO.Drawing.OfficeTextDecorationStyle.None;
            pendingLeadingIsExpandable = false;
            pendingLeadingIsTab = true;
            pendingTabSourceStyle = currentSpaceStyle;
            pendingTabSourceLine = lines.Count - 1;
        }

        void ResolveBufferedTab(double followingWidth, double spaceWidth, string followingText,
            PdfStandardFont font, double size, PdfTextBaseline baseline, RichTabFieldMeasurement? field = null) {
            if (pendingSpaceFragments.Count == 0 || !pendingSpaceFragments[0].Style.LeadingIsTab) return;
            PendingSpaceFragment fragment = pendingSpaceFragments[0];
            RichSeg tab = fragment.Style;
            double spaces = pendingLeadingAdvance - fragment.Advance;
            double selectedWidth = field != null && AlignsCompleteTabField(tab.LeadingTabStop?.Alignment ?? tab.LeadingTabAlignment)
                ? field.Width : followingWidth;
            double advance = CalculateTabAdvance(lineWidth, spaces + selectedWidth, spaceWidth,
                tab.LeadingTabAlignment, tabStopWidth, new string(' ', pendingLeadingSpaceCount) + followingText,
                font, size, baseline, options, CurrentMaxWidth(), tab.LeadingTabStop, CurrentLineOriginOffset(),
                currentRunNamedFont, currentRunFeatureSettings, currentRunHorizontalTextScaling, currentRunCharacterSpacing,
                decimalAnchorWidthOverride: ResolveRichDecimalTabAnchor(field, tab.LeadingTabAlignment, tab.LeadingTabStop, options, spaces));
            pendingLeadingAdvance += advance - fragment.Advance;
            fragment.Advance = advance;
        }

        void MarkCurrentLineHardBreak(RichSeg breakSegment) =>
            lines[lines.Count - 1].Add(breakSegment);

        for (int runIndex = 0; runIndex < effectiveRuns.Count; runIndex++) {
            PdfTextRun run = effectiveRuns[runIndex];
            currentTextRun = run;
            string text = (run.Text ?? string.Empty).Replace("\r\n", "\n").Replace('\r', '\n');
            bool bold = run.Bold;
            bool underline = run.Underline;
            bool strike = run.Strike;
            var underlineStyle = run.UnderlineStyle;
            var strikeStyle = run.StrikeStyle;
            bool italic = run.Italic;
            var color = run.Color;
            var backgroundColor = run.BackgroundColor;
            currentRunDecorationColor = run.DecorationColor;
            currentRunFeatureSettings = run.FeatureSettings;
            currentRunHorizontalTextScaling = run.HorizontalTextScaling;
            currentRunCharacterSpacing = run.CharacterSpacing;
            currentRunUnderlineStyle = underline && underlineStyle != OfficeIMO.Drawing.OfficeTextDecorationStyle.Words ? (underlineStyle == OfficeIMO.Drawing.OfficeTextDecorationStyle.None ? OfficeIMO.Drawing.OfficeTextDecorationStyle.Single : underlineStyle) : OfficeIMO.Drawing.OfficeTextDecorationStyle.None;
            currentRunUnderlineColor = run.DecorationColor ?? color;
            string? uri = run.LinkUri;
            string? destinationName = run.LinkDestinationName;
            string? contents = run.LinkContents;
            var baseline = run.Baseline;
            currentRunDecorationFontSize = EffectiveRichFontSize(run.FontSize ?? fontSize, baseline);
            currentRunDecorationTextRise = TextRiseForBaseline(run.FontSize ?? fontSize, baseline);
            var tabLeader = run.TabLeader;
            var tabAlignment = run.TabAlignment;
            var runBaseFont = run.Font.HasValue ? ChooseNormal(run.Font.Value) : baseFont;
            var fontForRun = (bold && italic) ? ChooseBoldItalic(runBaseFont) : bold ? ChooseBold(runBaseFont) : italic ? ChooseItalic(runBaseFont) : runBaseFont;
            currentRunNamedFont = options != null &&
                                  options.TryResolveNamedFontFace(run.FontFamily, bold, italic, out PdfNamedFontFace resolvedNamedFont)
                ? resolvedNamedFont
                : null;
            double runFontSize = run.FontSize ?? fontSize;
            bool whitespaceFootnote = run.InlineElement == null && run.FootnoteReference != null && text.All(character => character is ' ' or '\t' or '\n');
            if (whitespaceFootnote) {
                // Defer invisible associations until wrapping is complete so an anchor
                // never changes empty-line/leading-whitespace layout decisions.
                invisibleFootnoteAnchors ??= new List<(List<RichSeg> Line, RichSeg Anchor, RichSeg? SpaceStyle)>();
                invisibleFootnoteAnchors.Add((lines[lines.Count - 1], new RichSeg(string.Empty, bold, italic, underline, strike,
                    color, backgroundColor, uri, destinationName, contents, fontForRun, runFontSize, baseline, 0D,
                    namedFont: currentRunNamedFont, footnoteReference: run.FootnoteReference), null));
            }
            currentRunDecorationTextRise += ResolveFontLineAlignmentRise(fontForRun, currentRunNamedFont,
                currentRunDecorationFontSize, options, lineSpacing);
            currentRunContributesDescent = run.ContributesLineBoxDescent;
            bool usesSourceTextMetrics = UsesInlineRunTextMetrics(run.InlineElement);
            currentRunIsInline = run.InlineElement != null && !usesSourceTextMetrics;
            GetRichRunLineMetrics(fontForRun, currentRunNamedFont, runFontSize, baseline, options, lineSpacing,
                out currentRunAscent, out currentRunDescent);
            if (!currentRunContributesDescent) currentRunDescent = 0D;
            if (run.IsLineBoxStrut) {
                // An empty numbered paragraph still has its body-font paragraph
                // mark. It contributes metrics without adding ink or text advance.
                PrepareLineFrame(RunLineHeight(runFontSize));
                lines[lines.Count - 1].Add(new RichSeg(string.Empty, bold, italic, false, false,
                    null, null, null, null, null, fontForRun, runFontSize, baseline, 0D,
                    namedFont: currentRunNamedFont, contributesLineBoxDescent: currentRunContributesDescent,
                    isLineBoxStrut: true, formField: run.FormField));
                RegisterLineHeight(runFontSize);
                continue;
            }
            double spaceW = text.IndexOfAny(SoftLineSplitChars) >= 0
                ? MeasureRichText(" ", fontForRun, currentRunNamedFont, runFontSize, baseline, options, currentRunFeatureSettings, currentRunHorizontalTextScaling, currentRunCharacterSpacing)
                : 0D;
            currentSpaceStyle = preserveWhitespace ? new RichSeg(string.Empty, bold, italic, underline, strike,
                color, backgroundColor, uri, destinationName, contents, fontForRun, runFontSize, baseline, 0,
                leadingSpace: true, leadingAdvance: spaceW, leadingSpaceIsExpandable: !preformattedWhitespace,
                namedFont: currentRunNamedFont, underlineStyle: underlineStyle, strikeStyle: strikeStyle,
                decorationColor: currentRunDecorationColor, featureSettings: currentRunFeatureSettings,
                textDirection: run.TextDirection, leadingUnderlineStyle: currentRunUnderlineStyle,
                leadingDecorationColor: currentRunUnderlineColor, leadingDecorationFontSize: currentRunDecorationFontSize,
                leadingDecorationTextRise: currentRunDecorationTextRise,
                horizontalTextScaling: currentRunHorizontalTextScaling, characterSpacing: currentRunCharacterSpacing, contributesLineBoxDescent: currentRunContributesDescent, formField: currentTextRun?.FormField,
                footnoteReference: whitespaceFootnote ? run.FootnoteReference : null, leadingIsOrdinarySpace: true) : null;
            if (whitespaceFootnote && currentSpaceStyle != null) {
                int index = invisibleFootnoteAnchors!.Count - 1;
                var anchor = invisibleFootnoteAnchors[index];
                invisibleFootnoteAnchors[index] = (anchor.Line, anchor.Anchor, currentSpaceStyle);
            }
            if (run.InlineElement != null) {
                oversizedWordContinuation = false;
                PdfInlineElement inlineElement = run.InlineElement;
                double inlineHeight;
                if (usesSourceTextMetrics) {
                    if (inlineElement is PdfInlineFormWidget widget)
                        GetInlineFormWidgetLineMetrics(widget, baseFont,
                            fontSize, options, lineSpacing, out currentRunAscent, out currentRunDescent);
                    inlineHeight = RunLineHeight(runFontSize);
                } else {
                    currentRunAscent = Math.Max(0D, inlineElement.BaselineOffset + inlineElement.Height);
                    currentRunDescent = Math.Max(0D, -inlineElement.BaselineOffset);
                    inlineHeight = Math.Max(GetAscenderForOptions(baseFont, fontSize, options), inlineElement.BaselineOffset + inlineElement.Height)
                        + Math.Max(GetDescenderForOptions(baseFont, fontSize, options), -inlineElement.BaselineOffset);
                    if (lineSpacing?.FontLineBoxMultiplier != null)
                        inlineHeight = Math.Max(lineHeight, currentRunAscent + currentRunDescent);
                }
                PrepareLineFrame(inlineHeight, inlineElement.Width);
                double currentMaxWidth = CurrentMaxWidth();
                if (inlineElement.Width > currentMaxWidth + 0.001D &&
                    inlineElement is not PdfInlineImage { AllowFrameOverflow: true }) {
                    throw new ArgumentException("Inline element width exceeds the available paragraph width.");
                }

                RichTabFieldMeasurement? tabField = pendingLeadingIsTab && (AlignsCompleteTabField(pendingLeadingTabAlignment) || mayReselectNonLeftTabStop) ||
                    pendingSpaceFragments.Count > 0 && pendingSpaceFragments[0].Style.LeadingIsTab &&
                    (AlignsCompleteTabField(pendingSpaceFragments[0].Style.LeadingTabAlignment) || mayReselectNonLeftTabStop)
                        ? MeasureRichTabField(effectiveRuns, runIndex, 0, baseFont, fontSize, options, whitespaceMode)
                        : null;
                if (pendingLeadingIsTab) {
                    ResolvePendingLeadingTabForCurrentLine(inlineElement.Width, spaceW, string.Empty, fontForRun, runFontSize, baseline, tabField);
                    RevalidateLeadingTabFrame(inlineHeight, inlineElement.Width, spaceW, string.Empty, fontForRun, runFontSize, baseline, tabField);
                    currentMaxWidth = CurrentMaxWidth();
                }

                ResolveBufferedTab(inlineElement.Width, spaceW, string.Empty, fontForRun, runFontSize, baseline, tabField);

                List<RichSeg> currentLine = lines[lines.Count - 1];
                double leadingAdvance = currentLine.Count > 0 || pendingLeadingIsTab || whitespaceMode == PdfTextWhitespaceMode.Preserve ? pendingLeadingAdvance : 0D;
                if ((currentLine.Count > 0 || whitespaceMode == PdfTextWhitespaceMode.Preserve && leadingAdvance > 0D)
                    && MinimumCandidateWidth(CreateRunKerningProbe(string.Empty, run, baseFont, fontSize, options), null,
                        lineWidth + leadingAdvance + inlineElement.Width) > currentMaxWidth + 0.001D) {
                    if (pendingLeadingAdvance > 0D) {
                        MarkRichLineTextSeparator(currentLine);
                    }

                    StartNewLine();
                    PrepareLineFrame(inlineHeight, inlineElement.Width);
                    currentLine = lines[lines.Count - 1];
                    if (pendingLeadingIsTab) {
                        ResolvePendingLeadingTabForCurrentLine(inlineElement.Width, spaceW, string.Empty, fontForRun, runFontSize, baseline, tabField);
                        RevalidateLeadingTabFrame(inlineHeight, inlineElement.Width, spaceW, string.Empty, fontForRun, runFontSize, baseline, tabField);
                    }

                    leadingAdvance = pendingLeadingIsTab ? pendingLeadingAdvance : 0D;
                }

                if (preserveWhitespace && !pendingLeadingIsTab && pendingSpaceFragments.Count > 0
                    && (leadingAdvance != 0D || pendingLeadingAdvance == 0D)) {
                    currentLine.AddRange(ConsumePreservedSpaceFragments(pendingSpaceFragments, leadingAdvance));
                    lineWidth += leadingAdvance;
                    leadingAdvance = 0D;
                }
                // Flexible ink is resolved only after the actual frame, tab and
                // preserved-space placement are known. The source remains reusable.
                if (inlineElement is PdfInlineRule rule) {
                    double ruleFrameWidth = lines.Count == 1
                        ? firstLineInlineRuleWidth ?? inlineRuleWidth ?? CurrentMaxWidth()
                        : inlineRuleWidth ?? CurrentMaxWidth();
                    inlineElement = ResolveInlineRule(rule, Math.Min(CurrentMaxWidth(), ruleFrameWidth) - lineWidth - leadingAdvance, purpose);
                }
                currentLine.Add(new RichSeg(
                    string.Empty,
                    bold,
                    italic,
                    underline,
                    strike,
                    color,
                    backgroundColor,
                    uri,
                    destinationName,
                    contents,
                    fontForRun,
                    runFontSize,
                    baseline,
                    inlineElement.Width,
                    leadingSpace: leadingAdvance != 0D,
                    leadingAdvance: leadingAdvance,
                    leadingSpaceIsExpandable: pendingLeadingIsExpandable,
                    leadingTabLeader: pendingLeadingTabLeader,
                    inlineElement: inlineElement,
                    namedFont: currentRunNamedFont,
                    underlineStyle: underlineStyle,
                    strikeStyle: strikeStyle,
                    decorationColor: currentRunDecorationColor,
                    featureSettings: currentRunFeatureSettings, horizontalTextScaling: currentRunHorizontalTextScaling, characterSpacing: currentRunCharacterSpacing, contributesLineBoxDescent: currentRunContributesDescent, formField: currentTextRun?.FormField, footnoteReference: currentTextRun?.FootnoteReference,
                    leadingTabStop: pendingLeadingIsTab ? pendingLeadingTabStop : null,
                    leadingUnderlineStyle: leadingAdvance > 0D ? pendingLeadingUnderlineStyle : OfficeIMO.Drawing.OfficeTextDecorationStyle.None,
                    leadingDecorationColor: pendingLeadingDecorationColor,
                    leadingDecorationFontSize: pendingLeadingDecorationFontSize,
                    leadingDecorationTextRise: pendingLeadingDecorationTextRise,
                    leadingIsTab: pendingLeadingIsTab, leadingTabAlignment: pendingLeadingTabAlignment, leadingSpaceCount: pendingLeadingSpaceCount, leadingIsOrdinarySpace: pendingLeadingIsOrdinarySpace));
                lineWidth += leadingAdvance + inlineElement.Width;
                if (usesSourceTextMetrics) RegisterLineHeight(runFontSize);
                else RegisterLineMetrics();
                ResetPendingLeading();
                continue;
            }

            int idx = 0;
            while (idx < text.Length) {
                int tokenStart = idx;
                int nextWs = text.IndexOfAny(TokenSplitChars, idx);
                bool hadNewline = false;
                string token;
                if (nextWs == -1) { token = text.Substring(idx); idx = text.Length; } else {
                    token = text.Substring(idx, nextWs - idx);
                    hadNewline = text[nextWs] == '\n';
                    idx = nextWs + 1;
                }
                double tokenW = MeasureRichText(token, fontForRun, currentRunNamedFont, runFontSize, baseline, options, currentRunFeatureSettings, currentRunHorizontalTextScaling, currentRunCharacterSpacing);
                var continuation = nextWs < 0 && runIndex + 1 < effectiveRuns.Count ? wordContinuations[runIndex + 1] : default;
                RichSeg tokenProbe = CreateRunKerningProbe(token, run, baseFont, fontSize, options);
                double wordWidth = tokenW + continuation.Width;
                if (!oversizedWordContinuation && continuation.HasKerning)
                    wordWidth += runKerning.MeasureWordAdjustment(tokenProbe, continuation.Parts);
                double wordFontSize = Math.Max(runFontSize, continuation.FontSize);
                if (token.Length > 0) PrepareLineFrame(RunLineHeight(wordFontSize),
                    continuation.Width > 0D && wordWidth <= maxWidthPts ? wordWidth : 0D);
                if (whitespaceMode == PdfTextWhitespaceMode.Preserve && hadNewline && token.Length == 0 && !pendingLeadingIsTab)
                    ResetPendingLeading();
                RichTabFieldMeasurement? tabField = token.Length > 0 &&
                    (pendingLeadingIsTab && (AlignsCompleteTabField(pendingLeadingTabAlignment) || mayReselectNonLeftTabStop) ||
                     pendingSpaceFragments.Count > 0 && pendingSpaceFragments[0].Style.LeadingIsTab &&
                     (AlignsCompleteTabField(pendingSpaceFragments[0].Style.LeadingTabAlignment) || mayReselectNonLeftTabStop))
                        ? MeasureRichTabField(effectiveRuns, runIndex, tokenStart, baseFont, fontSize, options, whitespaceMode)
                        : null;
                if (token.Length > 0) ResolveBufferedTab(wordWidth, spaceW, token, fontForRun, runFontSize, baseline, tabField);
                double candidateLeading = lines[lines.Count - 1].Count == 0 && !preserveWhitespace && !pendingLeadingIsTab
                    ? 0D : pendingLeadingAdvance;
                double candidateBoundary = !pendingLeadingSeparator && lines[lines.Count - 1].Count > 0
                    ? runKerning.MeasureBoundary(lines[lines.Count - 1], tokenProbe) : 0D;
                var trialLine = lines[lines.Count - 1];
                var completeWordCandidate = compressWordSpaces && token.Length > 0
                    ? CompressionCandidate(tokenProbe, continuation.Parts) : null;
                double completeWordNaturalWidth = lineWidth + candidateLeading + wordWidth + candidateBoundary;
                bool completeWordFits = completeWordCandidate != null &&
                    MeasureWordCompressionMinimumWidth(completeWordCandidate, completeWordNaturalWidth) <= CurrentMaxWidth() + RichTextWidthTolerance;
                // Stateful frame callbacks cannot be replayed to compare alternatives. In a
                // fixed frame, keep selection separate from feasibility and source consumption.
                bool preferEarlierWordBreak = wordJustificationMode == PdfJustificationMode.Words && lineLayout == null &&
                    completeWordFits && pendingLeadingSeparator && pendingLeadingIsOrdinarySpace && !pendingLeadingIsTab &&
                    wordWidth <= maxWidthPts + RichTextWidthTolerance &&
                    PreferWordExpansionBreak(trialLine, completeWordCandidate!, lineWidth, completeWordNaturalWidth, CurrentMaxWidth());
                if (completeWordFits && !preferEarlierWordBreak && preserveWhitespace && !pendingLeadingIsTab && pendingSpaceFragments.Count > 0) {
                    // Trial the complete word before natural-width clipping can discard
                    // its authored separators. Commit their full natural advances only.
                    lines[lines.Count - 1].AddRange(ConsumePreservedSpaceFragments(pendingSpaceFragments, pendingLeadingAdvance));
                    lineWidth += pendingLeadingAdvance;
                    ResetPendingLeading();
                }
                if (preserveWhitespace && !pendingLeadingIsTab && pendingSpaceFragments.Count > 0 && (token.Length > 0 || hadNewline)) {
                    // Retain the originating runs, including clipped portions of literal
                    // spaces, so a later cell/column continuation can remeasure them.
                    if (pendingLeadingAdvance <= 0.001D) {
                        lines[lines.Count - 1].AddRange(ConsumePreservedSpaceFragments(pendingSpaceFragments, pendingLeadingAdvance));
                        lineWidth += pendingLeadingAdvance;
                    }
                    while (pendingLeadingAdvance > 0.001D) {
                        double available = CurrentMaxWidth() - lineWidth;
                        if (available <= 0.001D) {
                            // Imported indentation can clamp a measurement frame to
                            // effectively zero width. Flow spaces cannot create capacity.
                            if (!preformattedWhitespace && CurrentMaxWidth() <= 0.001D) break;
                            StartNewLine();
                            PrepareLineFrame(RunLineHeight(runFontSize));
                            // Flow separators that overflow a completed line do
                            // not become indentation on its continuation.
                            if (!preformattedWhitespace) break;
                            available = CurrentMaxWidth();
                            if (available <= 0.001D) {
                                if (!preformattedWhitespace) break;
                                throw new InvalidOperationException("No width is available for preserved text spacing.");
                            }
                        }
                        double advance = Math.Min(available, pendingLeadingAdvance);
                        lines[lines.Count - 1].AddRange(ConsumePreservedSpaceFragments(pendingSpaceFragments, advance));
                        RegisterLineHeight(runFontSize);
                        lineWidth += advance;
                        pendingLeadingAdvance -= advance;
                        if (!preformattedWhitespace) break;
                    }
                    ResetPendingLeading();
                }
                var lastLine = lines[lines.Count - 1];
                double needed = lastLine.Count == 0 && !preserveWhitespace ? tokenW : pendingLeadingAdvance + tokenW;
                double currentMaxWidth = CurrentMaxWidth();

                // A list prefix consumes part of the wider hanging first line.
                // Plan its first word against the remaining frame so a later
                // hyphen/CJK chunk can use the narrower continuation frame.
                if (!completeWordFits && lastLine.Count > 0 && pendingLeadingAdvance == 0D && lineWidth + tokenW > currentMaxWidth &&
                    lastLine[lastLine.Count - 1].InlineElement is PdfInlineBox { IsTextSpacer: true } &&
                    TryAppendMultilingualLongToken(token, bold, italic, underline, strike, underlineStyle, strikeStyle,
                        color, backgroundColor, uri, destinationName, contents, fontForRun, runFontSize, baseline,
                        Math.Max(0D, currentMaxWidth - lineWidth))) {
                    if (hadNewline) {
                        oversizedWordContinuation = false;
                        MarkCurrentLineHardBreak(CreateRichLineBreakSegment(run, fontForRun, runFontSize, currentRunNamedFont));
                        StartNewLine();
                        ResetPendingLeading();
                    } else if (nextWs != -1) {
                        SetPendingSeparator(text[nextWs] == '\t', spaceW, tabAlignment, tabLeader);
                    }
                    continue;
                }

                bool continuesOversizedWord = oversizedWordContinuation && !pendingLeadingSeparator && lastLine.Count > 0;
                oversizedWordContinuation |= !completeWordFits && wordWidth > Math.Min(currentMaxWidth, maxWidthPts) + RichTextWidthTolerance;
                double continuationBoundary = continuesOversizedWord ? runKerning.MeasureBoundary(lastLine, tokenProbe) : 0D;
                if (!completeWordFits && (tokenW > currentMaxWidth + RichTextWidthTolerance ||
                    continuesOversizedWord && lineWidth + tokenW + continuationBoundary > currentMaxWidth + RichTextWidthTolerance)) {
                    if (lastLine.Count > 0 && !continuesOversizedWord) { StartNewLine(); lastLine = lines[lines.Count - 1]; }
                    ResetPendingLeading();
                    if (TryAppendSoftLineBreakLongToken(token, bold, italic, underline, strike, underlineStyle, strikeStyle, color, backgroundColor, uri, destinationName, contents, fontForRun, runFontSize, baseline)) {
                        if (hadNewline) {
                        oversizedWordContinuation = false;
                            MarkCurrentLineHardBreak(CreateRichLineBreakSegment(run, fontForRun, runFontSize, currentRunNamedFont));
                            StartNewLine();
                            ResetPendingLeading();
                        } else if (nextWs != -1) {
                            bool hadTab = text[nextWs] == '\t';
                            SetPendingSeparator(hadTab, spaceW, tabAlignment, tabLeader);
                        }
                        continue;
                    }

                    if (TryAppendDelimitedLongToken(token, bold, italic, underline, strike, underlineStyle, strikeStyle, color, backgroundColor, uri, destinationName, contents, fontForRun, runFontSize, baseline)) {
                        if (hadNewline) {
                        oversizedWordContinuation = false;
                            MarkCurrentLineHardBreak(CreateRichLineBreakSegment(run, fontForRun, runFontSize, currentRunNamedFont));
                            StartNewLine();
                            ResetPendingLeading();
                        } else if (nextWs != -1) {
                            bool hadTab = text[nextWs] == '\t';
                            SetPendingSeparator(hadTab, spaceW, tabAlignment, tabLeader);
                        }
                        continue;
                    }

                    if (TryAppendHyphenatedLongToken(token, bold, italic, underline, strike, underlineStyle, strikeStyle, color, backgroundColor, uri, destinationName, contents, fontForRun, runFontSize, baseline)) {
                        if (hadNewline) {
                        oversizedWordContinuation = false;
                            MarkCurrentLineHardBreak(CreateRichLineBreakSegment(run, fontForRun, runFontSize, currentRunNamedFont));
                            StartNewLine();
                            ResetPendingLeading();
                        } else if (nextWs != -1) {
                            bool hadTab = text[nextWs] == '\t';
                            SetPendingSeparator(hadTab, spaceW, tabAlignment, tabLeader);
                        }
                        continue;
                    }

                    if (TryAppendMultilingualLongToken(token, bold, italic, underline, strike, underlineStyle, strikeStyle, color, backgroundColor, uri, destinationName, contents, fontForRun, runFontSize, baseline)) {
                        if (hadNewline) {
                        oversizedWordContinuation = false;
                            MarkCurrentLineHardBreak(CreateRichLineBreakSegment(run, fontForRun, runFontSize, currentRunNamedFont));
                            StartNewLine();
                            ResetPendingLeading();
                        } else if (nextWs != -1) {
                            bool hadTab = text[nextWs] == '\t';
                            SetPendingSeparator(hadTab, spaceW, tabAlignment, tabLeader);
                        }
                        continue;
                    }

                    int pos = 0;
                    bool measureKernedChunks = currentRunFeatureSettings.TryGetValue("kern", out int chunkKerning) && chunkKerning > 0
                        && (options?.TextShapingProviderSnapshot == null || ReferenceEquals(options.TextShapingProviderSnapshot, OfficeManagedTextShapingProvider.Instance))
                        && !OfficeManagedTextShaper.RequiresComplexLayout(token);
                    while (pos < token.Length) {
                        int take = 0;
                        double chunkW = 0;
                        PrepareLineFrame(RunLineHeight(runFontSize));
                        currentMaxWidth = CurrentMaxWidth();
                        double availableChunkWidth = currentMaxWidth - lineWidth;
                        while (pos + take < token.Length) {
                            int scalarLength = GetScalarUtf16Length(token, pos + take);
                            double candidateWidth = measureKernedChunks
                                ? MeasureRichText(token.Substring(pos, take + scalarLength), fontForRun, currentRunNamedFont, runFontSize, baseline, options, currentRunFeatureSettings, currentRunHorizontalTextScaling, currentRunCharacterSpacing)
                                : chunkW + MeasureRichText(token.Substring(pos + take, scalarLength), fontForRun, currentRunNamedFont, runFontSize, baseline, options, currentRunFeatureSettings, currentRunHorizontalTextScaling, currentRunCharacterSpacing);
                            candidateWidth += runKerning.MeasureBoundary(lastLine,
                                CreateRunKerningProbe(token.Substring(pos, take + scalarLength), run, baseFont, fontSize, options));
                            if ((take > 0 || lastLine.Count > 0) && candidateWidth > availableChunkWidth + RichTextWidthTolerance) {
                                break;
                            }

                            chunkW = candidateWidth;
                            take += scalarLength;
                            if (chunkW >= availableChunkWidth) {
                                break;
                            }
                        }

                        if (take == 0) {
                            if (lastLine.Count > 0) {
                                StartNewLine(); lastLine = lines[lines.Count - 1];
                                continue;
                            }
                            take = GetScalarUtf16Length(token, pos);
                            chunkW = MeasureRichText(token.Substring(pos, take), fontForRun, currentRunNamedFont, runFontSize, baseline, options, currentRunFeatureSettings, currentRunHorizontalTextScaling, currentRunCharacterSpacing);
                        }

                        string chunk = token.Substring(pos, take);
                        double measuredChunkWidth = MeasureRichText(chunk, fontForRun, currentRunNamedFont, runFontSize, baseline, options, currentRunFeatureSettings, currentRunHorizontalTextScaling, currentRunCharacterSpacing);
                        lastLine.Add(new RichSeg(chunk, bold, italic, underline, strike, color, backgroundColor, uri, destinationName, contents, fontForRun, runFontSize, baseline, measuredChunkWidth, namedFont: currentRunNamedFont, underlineStyle: underlineStyle, strikeStyle: strikeStyle, decorationColor: currentRunDecorationColor, featureSettings: currentRunFeatureSettings, horizontalTextScaling: currentRunHorizontalTextScaling, characterSpacing: currentRunCharacterSpacing, contributesLineBoxDescent: currentRunContributesDescent, formField: currentTextRun?.FormField, footnoteReference: currentTextRun?.FootnoteReference));
                        RegisterLineHeight(runFontSize);
                        lineWidth += chunkW;
                        pos += take;
                        if (pos < token.Length) { StartNewLine(); lastLine = lines[lines.Count - 1]; }
                    }
                    if (hadNewline) {
                        oversizedWordContinuation = false;
                        MarkCurrentLineHardBreak(CreateRichLineBreakSegment(run, fontForRun, runFontSize, currentRunNamedFont));
                        StartNewLine();
                        ResetPendingLeading();
                    } else if (nextWs != -1) {
                        bool hadTab = text[nextWs] == '\t';
                        SetPendingSeparator(hadTab, spaceW, tabAlignment, tabLeader);
                    }
                    continue;
                }
                if (token.Length > 0 && pendingLeadingIsTab) {
                    ResolvePendingLeadingTabForCurrentLine(wordWidth, spaceW, token, fontForRun, runFontSize, baseline, tabField);
                    RevalidateLeadingTabFrame(RunLineHeight(runFontSize), wordWidth, spaceW, token, fontForRun, runFontSize, baseline, tabField);
                    currentMaxWidth = CurrentMaxWidth();
                }
                // Formatting boundaries do not introduce a word-break opportunity.
                // Reserve the rest of a word spanning runs when it fits a fresh line.
                double wrappingWidth = completeWordFits || !oversizedWordContinuation && wordWidth <= Math.Min(currentMaxWidth, maxWidthPts) + RichTextWidthTolerance ? wordWidth : tokenW;
                double boundaryAdvance = !pendingLeadingSeparator && lastLine.Count > 0
                    ? runKerning.MeasureBoundary(lastLine, tokenProbe) : 0D;
                needed = lastLine.Count == 0
                    ? (pendingLeadingIsTab || preserveWhitespace ? pendingLeadingAdvance + wrappingWidth : wrappingWidth)
                    : pendingLeadingAdvance + wrappingWidth;
                // A word that overflows the line end wraps at its last fitting Unicode line-break
                // opportunity (hyphen, solidus, CJK boundary) or opt-in hyphenation point.
                if (MinimumCandidateWidth(tokenProbe, continuation.Parts, lineWidth + needed + boundaryAdvance) > currentMaxWidth + RichTextWidthTolerance && lastLine.Count > 0
                    && token.Length > 1 && continuation.Width <= 0D && !pendingLeadingIsTab && whitespaceMode != PdfTextWhitespaceMode.Preserve
                    && TrySplitTokenAtLineEnd(token, currentMaxWidth + RichTextWidthTolerance - lineWidth - boundaryAdvance -
                            (pendingLeadingSeparator ? pendingLeadingAdvance : 0D), options,
                        part => MeasureRichText(part, fontForRun, currentRunNamedFont, runFontSize, baseline, options, currentRunFeatureSettings, currentRunHorizontalTextScaling, currentRunCharacterSpacing),
                        out string splitHead, out double splitHeadWidth, out int splitConsumed)) {
                    bool splitLeadingSpace = pendingLeadingSeparator && lineWidth != 0D;
                    double splitLeadingAdvance = splitLeadingSpace ? pendingLeadingAdvance : 0D;
                    lastLine.Add(new RichSeg(splitHead, bold, italic, underline, strike, color, backgroundColor, uri, destinationName, contents, fontForRun, runFontSize, baseline, splitHeadWidth, splitLeadingSpace, splitLeadingAdvance, pendingLeadingIsExpandable, PdfTabLeaderStyle.None, namedFont: currentRunNamedFont, underlineStyle: underlineStyle, strikeStyle: strikeStyle, decorationColor: currentRunDecorationColor, featureSettings: currentRunFeatureSettings, horizontalTextScaling: currentRunHorizontalTextScaling, characterSpacing: currentRunCharacterSpacing, contributesLineBoxDescent: currentRunContributesDescent, formField: currentTextRun?.FormField, footnoteReference: currentTextRun?.FootnoteReference, leadingUnderlineStyle: splitLeadingSpace ? pendingLeadingUnderlineStyle : OfficeIMO.Drawing.OfficeTextDecorationStyle.None, leadingDecorationColor: pendingLeadingDecorationColor, leadingDecorationFontSize: pendingLeadingDecorationFontSize, leadingDecorationTextRise: pendingLeadingDecorationTextRise, leadingSpaceCount: pendingLeadingSpaceCount, leadingIsOrdinarySpace: splitLeadingSpace && pendingLeadingIsOrdinarySpace));
                    RegisterLineHeight(runFontSize);
                    lineWidth += splitHeadWidth + splitLeadingAdvance;
                    ResetPendingLeading();
                    StartNewLine();
                    token = token.Substring(splitConsumed);
                    tokenW = MeasureRichText(token, fontForRun, currentRunNamedFont, runFontSize, baseline, options, currentRunFeatureSettings, currentRunHorizontalTextScaling, currentRunCharacterSpacing);
                    PrepareLineFrame(RunLineHeight(runFontSize), tokenW);
                    lastLine = lines[lines.Count - 1];
                    needed = tokenW;
                    currentMaxWidth = CurrentMaxWidth();
                }
                if ((preferEarlierWordBreak && ReferenceEquals(lastLine, trialLine) ||
                     MinimumCandidateWidth(tokenProbe, continuation.Parts, lineWidth + needed + boundaryAdvance) > currentMaxWidth + RichTextWidthTolerance) && lastLine.Count > 0
                    && (token.Length > 0 || whitespaceMode != PdfTextWhitespaceMode.Preserve)) {
                    if (pendingLeadingAdvance > 0D) {
                        MarkRichLineTextSeparator(lastLine);
                    }

                    if (whitespaceMode == PdfTextWhitespaceMode.Preserve) {
                        // Trailing spaces do not participate in the previous line's
                        // alignment or justification after its following word wraps.
                        while (lastLine.Count > 0 && lastLine[lastLine.Count - 1] is RichSeg trailing
                               && trailing.Text.Length == 0 && trailing.InlineElement == null
                               && trailing.LeadingSpace && !trailing.LeadingIsTab) {
                            lastLine.RemoveAt(lastLine.Count - 1);
                        }
                        MarkRichLineTextSeparator(lastLine);
                    }
                    StartNewLine();
                    // The next line can enter a narrower floating frame. A pending word
                    // that fit the previous frame must obtain enough space again.
                    PrepareLineFrame(RunLineHeight(wordFontSize), wrappingWidth);
                    if (token.Length > 0 && pendingLeadingIsTab) {
                        ResolvePendingLeadingTabForCurrentLine(wordWidth, spaceW, token, fontForRun, runFontSize, baseline, tabField);
                        RevalidateLeadingTabFrame(RunLineHeight(runFontSize), wordWidth, spaceW, token, fontForRun, runFontSize, baseline, tabField);
                    }
                }
                if (token.Length > 0) {
                    bool needsLeadingSpace = pendingLeadingSeparator && (lineWidth != 0D || pendingLeadingIsTab || preserveWhitespace);
                    double leadingAdvance = needsLeadingSpace ? pendingLeadingAdvance : 0;
                    var currentLine = lines[lines.Count - 1];
                    double segmentWidth = tokenW + leadingAdvance + (!pendingLeadingSeparator && currentLine.Count > 0
                        ? runKerning.MeasureBoundary(currentLine, tokenProbe) : 0D);
                    var segmentLeader = needsLeadingSpace ? pendingLeadingTabLeader : PdfTabLeaderStyle.None;
                    lines[lines.Count - 1].Add(new RichSeg(token, bold, italic, underline, strike, color, backgroundColor, uri, destinationName, contents, fontForRun, runFontSize, baseline, tokenW, needsLeadingSpace, leadingAdvance, pendingLeadingIsExpandable, segmentLeader, namedFont: currentRunNamedFont, underlineStyle: underlineStyle, strikeStyle: strikeStyle, decorationColor: currentRunDecorationColor, featureSettings: currentRunFeatureSettings, horizontalTextScaling: currentRunHorizontalTextScaling, characterSpacing: currentRunCharacterSpacing, contributesLineBoxDescent: currentRunContributesDescent, formField: currentTextRun?.FormField, footnoteReference: currentTextRun?.FootnoteReference, leadingTabStop: pendingLeadingIsTab ? pendingLeadingTabStop : null, leadingUnderlineStyle: needsLeadingSpace ? pendingLeadingUnderlineStyle : OfficeIMO.Drawing.OfficeTextDecorationStyle.None, leadingDecorationColor: pendingLeadingDecorationColor, leadingDecorationFontSize: pendingLeadingDecorationFontSize, leadingDecorationTextRise: pendingLeadingDecorationTextRise, leadingIsTab: pendingLeadingIsTab, leadingTabAlignment: pendingLeadingTabAlignment, leadingSpaceCount: pendingLeadingSpaceCount, leadingIsOrdinarySpace: pendingLeadingIsOrdinarySpace));
                    RegisterLineHeight(runFontSize);
                    lineWidth += segmentWidth;
                    ResetPendingLeading();
                }
                if (hadNewline) {
                    oversizedWordContinuation = false;
                    PrepareLineFrame(RunLineHeight(runFontSize));
                    RegisterLineHeight(runFontSize);
                    MarkCurrentLineHardBreak(CreateRichLineBreakSegment(run, fontForRun, runFontSize, currentRunNamedFont));
                    StartNewLine();
                    ResetPendingLeading();
                } else if (nextWs != -1) {
                    bool hadTab = text[nextWs] == '\t';
                    SetPendingSeparator(hadTab, spaceW, tabAlignment, tabLeader);
                }
            }
        }
        if (invisibleFootnoteAnchors != null) {
            // Wrapping can discard a trailing space after it consumed the buffered
            // reference. Only surviving fragments establish actual line ownership.
            var survivingOrigins = new HashSet<RichSeg>(lines.SelectMany(line => line)
                .Where(segment => segment.FootnoteReference != null && segment.FootnoteReferenceOrigin != null)
                .Select(segment => segment.FootnoteReferenceOrigin!));
            foreach (var anchor in invisibleFootnoteAnchors) {
                if (anchor.SpaceStyle != null && survivingOrigins.Contains(anchor.SpaceStyle)) continue;
                int index = anchor.Line.Count;
                if (index > 0 && (anchor.Line[index - 1].EndsWithHardBreak || anchor.Line[index - 1].EndsWithTextSeparator)) index--;
                anchor.Line.Insert(index, anchor.Anchor);
            }
        }
        if (lines.Count > 0 && lines[lines.Count - 1].Count == 0) { lines.RemoveAt(lines.Count - 1); }
        if (heights.Count < lines.Count) heights.Add(currentLineHeight);
        if (admittedLineWidths.Count < lines.Count) admittedLineWidths.Add(CurrentMaxWidth());
        for (int lineIndex = 0; lineIndex < lines.Count; lineIndex++) {
            var line = (RichLine)lines[lineIndex];
            runKerning.Apply(line);
            if (compressWordSpaces) BindWordCompressionPlan(line, admittedLineWidths[lineIndex]);
            foreach (RichSeg segment in line) segment.LineAlignmentRise = ResolveRichRunLineAlignmentRise(segment, options, lineSpacing);
            line.BaselineOffset = ResolveRichLineBaselineOffset(line, baseFont, fontSize, options, lineSpacing);
            if (lineSpacing == null) {
                // A short baseline advance may overlap subsequent lines, but bottom-aligned
                // table content must still reserve the last line's descent before its clip.
                GetRichLineInkMetrics(line, options ?? new PdfOptions(), out _, out double descent);
                line.MinimumContentHeight = line.BaselineOffset + descent;
            }
            if (lineSpacing?.FontLineBoxMultiplier != null && lineSpacing.Rule == PdfLineSpacingRule.Multiple)
                line.FrameFitHeight = ResolveImportedLineFrameFitHeight(line, baseFont, fontSize, options, lineSpacing);
        }
        return (lines, heights);

        double MeasureLongTokenBoundary(string part) => runKerning.MeasureBoundary(lines[lines.Count - 1],
            CreateRunKerningProbe(part, currentTextRun!, baseFont, fontSize, options));

        List<PdfTextTokenChunk>? PlanLongTokenChunks(string token, int[] breakpoints,
            Func<string, double> measure, bool allowFinalOverflow = false, bool appendHyphen = false,
            double? firstChunkWidth = null) {
            double available = firstChunkWidth ?? Math.Max(0D, CurrentMaxWidth() - lineWidth);
            var chunks = TryBuildTokenChunks(token, breakpoints, measure, available, maxWidthPts,
                allowFinalOverflow, part => measure(part) + MeasureLongTokenBoundary(part), appendHyphen, allowSingleChunk: true);
            if (chunks == null && lines[lines.Count - 1].Count > 0) {
                chunks = TryBuildTokenChunks(token, breakpoints, measure, maxWidthPts, maxWidthPts,
                    allowFinalOverflow, appendHyphen: appendHyphen, allowSingleChunk: true);
                if (chunks != null) StartNewLine();
            }
            return chunks;
        }

        bool TryAppendSoftLineBreakLongToken(
            string token,
            bool bold,
            bool italic,
            bool underline,
            bool strike,
            OfficeIMO.Drawing.OfficeTextDecorationStyle underlineStyle,
            OfficeIMO.Drawing.OfficeTextDecorationStyle strikeStyle,
            PdfColor? color,
            PdfColor? backgroundColor,
            string? uri,
            string? destinationName,
            string? contents,
            PdfStandardFont font,
            double runFontSize,
            PdfTextBaseline baseline) {
            var chunks = PlanLongTokenChunks(token, GetValidSoftLineBreakpoints(token, options),
                part => MeasureRichText(part, font, currentRunNamedFont, runFontSize, baseline, options,
                    currentRunFeatureSettings, currentRunHorizontalTextScaling, currentRunCharacterSpacing), allowFinalOverflow: true);

            if (chunks == null) {
                return false;
            }

            for (int chunkIndex = 0; chunkIndex < chunks.Count; chunkIndex++) {
                PdfTextTokenChunk chunk = chunks[chunkIndex];
                PrepareLineFrame(RunLineHeight(runFontSize), chunk.Width);
                double boundaryAdvance = MeasureLongTokenBoundary(chunk.Text);
                lines[lines.Count - 1].Add(new RichSeg(chunk.Text, bold, italic, underline, strike, color, backgroundColor, uri, destinationName, contents, font, runFontSize, baseline, chunk.Width, namedFont: currentRunNamedFont, underlineStyle: underlineStyle, strikeStyle: strikeStyle, decorationColor: currentRunDecorationColor, featureSettings: currentRunFeatureSettings, horizontalTextScaling: currentRunHorizontalTextScaling, characterSpacing: currentRunCharacterSpacing, contributesLineBoxDescent: currentRunContributesDescent, formField: currentTextRun?.FormField, footnoteReference: currentTextRun?.FootnoteReference));
                RegisterLineHeight(runFontSize);
                lineWidth += chunk.Width + boundaryAdvance;
                if (chunkIndex + 1 < chunks.Count) {
                    StartNewLine();
                }
            }

            return true;
        }

        bool TryAppendHyphenatedLongToken(
            string token,
            bool bold,
            bool italic,
            bool underline,
            bool strike,
            OfficeIMO.Drawing.OfficeTextDecorationStyle underlineStyle,
            OfficeIMO.Drawing.OfficeTextDecorationStyle strikeStyle,
            PdfColor? color,
            PdfColor? backgroundColor,
            string? uri,
            string? destinationName,
            string? contents,
            PdfStandardFont font,
            double runFontSize,
            PdfTextBaseline baseline) {
            var plannedChunks = PlanLongTokenChunks(token, GetValidHyphenationBreakpoints(token, options),
                part => MeasureRichText(part, font, currentRunNamedFont, runFontSize, baseline, options,
                    currentRunFeatureSettings, currentRunHorizontalTextScaling, currentRunCharacterSpacing), appendHyphen: true);
            if (plannedChunks == null) return false;

            for (int chunkIndex = 0; chunkIndex < plannedChunks.Count; chunkIndex++) {
                PdfTextTokenChunk planned = plannedChunks[chunkIndex];
                string selectedText = planned.Text;
                double measuredWidth = MeasureRichText(selectedText, font, currentRunNamedFont, runFontSize, baseline, options, currentRunFeatureSettings, currentRunHorizontalTextScaling, currentRunCharacterSpacing);
                PrepareLineFrame(RunLineHeight(runFontSize), measuredWidth);
                double boundaryAdvance = MeasureLongTokenBoundary(selectedText);
                lines[lines.Count - 1].Add(new RichSeg(selectedText, bold, italic, underline, strike, color, backgroundColor, uri, destinationName, contents, font, runFontSize, baseline, measuredWidth, namedFont: currentRunNamedFont, underlineStyle: underlineStyle, strikeStyle: strikeStyle, decorationColor: currentRunDecorationColor, featureSettings: currentRunFeatureSettings, horizontalTextScaling: currentRunHorizontalTextScaling, characterSpacing: currentRunCharacterSpacing, contributesLineBoxDescent: currentRunContributesDescent, formField: currentTextRun?.FormField, footnoteReference: currentTextRun?.FootnoteReference));
                RegisterLineHeight(runFontSize);
                lineWidth += measuredWidth + boundaryAdvance;
                if (chunkIndex < plannedChunks.Count - 1) {
                    StartNewLine();
                }
            }

            return true;
        }

        bool TryAppendDelimitedLongToken(
            string token,
            bool bold,
            bool italic,
            bool underline,
            bool strike,
            OfficeIMO.Drawing.OfficeTextDecorationStyle underlineStyle,
            OfficeIMO.Drawing.OfficeTextDecorationStyle strikeStyle,
            PdfColor? color,
            PdfColor? backgroundColor,
            string? uri,
            string? destinationName,
            string? contents,
            PdfStandardFont font,
            double runFontSize,
            PdfTextBaseline baseline) {
            int[] breakpoints = GetValidLongTokenDelimiterBreakpoints(token);
            if (breakpoints.Length == 0) {
                return false;
            }

            int position = 0;
            bool startOnFreshLine = false;
            double firstAvailable = Math.Max(0D, CurrentMaxWidth() - lineWidth);
            var plannedChunks = new System.Collections.Generic.List<(string Text, double Width)>();
            while (position < token.Length) {
                int selectedBreak = -1;
                string selectedText = string.Empty;
                double selectedWidth = 0D;
                double maxWidthForChunk = plannedChunks.Count == 0 && !startOnFreshLine ? firstAvailable : maxWidthPts;
                if (TryPlanDelimiterBoundedWordGroup(maxWidthForChunk, out selectedText, out selectedWidth)) {
                    selectedBreak = position + selectedText.Length;
                    plannedChunks.Add((selectedText, selectedWidth));
                    position = selectedBreak;
                    continue;
                }

                int[] candidates = breakpoints
                    .Where(point => point > position)
                    .Concat(new[] { token.Length })
                    .Distinct()
                    .OrderBy(point => point)
                    .ToArray();

                foreach (int candidate in candidates) {
                    string chunkText = token.Substring(position, candidate - position);
                    if (chunkText.Length == 0) {
                        continue;
                    }

                    double chunkWidth = MeasurePlannedChunk(chunkText);
                    if (chunkWidth <= maxWidthForChunk || selectedBreak < 0) {
                        if (chunkWidth <= maxWidthForChunk) {
                            selectedBreak = candidate;
                            selectedText = chunkText;
                            selectedWidth = chunkWidth;
                        }
                    }

                    if (chunkWidth > maxWidthForChunk && selectedBreak >= 0) {
                        break;
                    }
                }

                if (selectedBreak > position &&
                    selectedBreak < token.Length &&
                    CanExtendDelimitedIdentifierChunk(token, selectedBreak)) {
                    TryExtendIdentifierChunkToAvailableWidth(maxWidthForChunk, ref selectedText, ref selectedWidth, ref selectedBreak);
                }

                if (selectedBreak <= position || selectedText.Length == 0) {
                    if (!TryPlanCharacterChunkToNextDelimiter(maxWidthForChunk, out selectedText, out selectedWidth)) {
                        if (position == 0 && lineWidth > 0D && !startOnFreshLine) {
                            startOnFreshLine = true;
                            continue;
                        }
                        return false;
                    }

                    selectedBreak = position + selectedText.Length;
                }

                plannedChunks.Add((selectedText, selectedWidth));
                position = selectedBreak;
            }

            if (startOnFreshLine) StartNewLine();
            for (int chunkIndex = 0; chunkIndex < plannedChunks.Count; chunkIndex++) {
                (string selectedText, double selectedWidth) = plannedChunks[chunkIndex];
                double measuredWidth = MeasureRichText(selectedText, font, currentRunNamedFont, runFontSize, baseline, options, currentRunFeatureSettings, currentRunHorizontalTextScaling, currentRunCharacterSpacing);
                PrepareLineFrame(RunLineHeight(runFontSize), measuredWidth);
                double boundaryAdvance = MeasureLongTokenBoundary(selectedText);
                lines[lines.Count - 1].Add(new RichSeg(selectedText, bold, italic, underline, strike, color, backgroundColor, uri, destinationName, contents, font, runFontSize, baseline, measuredWidth, namedFont: currentRunNamedFont, underlineStyle: underlineStyle, strikeStyle: strikeStyle, decorationColor: currentRunDecorationColor, featureSettings: currentRunFeatureSettings, horizontalTextScaling: currentRunHorizontalTextScaling, characterSpacing: currentRunCharacterSpacing, contributesLineBoxDescent: currentRunContributesDescent, formField: currentTextRun?.FormField, footnoteReference: currentTextRun?.FootnoteReference));
                RegisterLineHeight(runFontSize);
                lineWidth += measuredWidth + boundaryAdvance;
                if (chunkIndex < plannedChunks.Count - 1) {
                    StartNewLine();
                }
            }

            return true;

            double MeasurePlannedChunk(string text) =>
                MeasureRichText(text, font, currentRunNamedFont, runFontSize, baseline, options,
                    currentRunFeatureSettings, currentRunHorizontalTextScaling, currentRunCharacterSpacing)
                + (position == 0 && !startOnFreshLine ? MeasureLongTokenBoundary(text) : 0D);

            bool TryPlanDelimiterBoundedWordGroup(double maxWidthForChunk, out string selectedText, out double selectedWidth) {
                selectedText = string.Empty;
                selectedWidth = 0D;
                if (position >= token.Length || !IsLongTokenDelimiterBreakChar(token[position])) {
                    return false;
                }

                int nextDelimiterIndex = -1;
                for (int index = position + 1; index < token.Length; index++) {
                    if (IsLongTokenDelimiterBreakChar(token[index])) {
                        nextDelimiterIndex = index;
                        break;
                    }
                }

                if (nextDelimiterIndex <= position + 1) {
                    return false;
                }

                string candidate = token.Substring(position, nextDelimiterIndex - position + 1);
                if (!IsDelimiterBoundedWordGroup(candidate)) {
                    return false;
                }

                double candidateWidth = MeasurePlannedChunk(candidate);
                if (candidateWidth <= maxWidthForChunk ||
                    CanKeepDelimitedWordSegmentOverSoftLimit(candidate, candidateWidth, maxWidthForChunk, allowBoundaryDelimiters: true)) {
                    selectedText = candidate;
                    selectedWidth = candidateWidth;
                    return true;
                }

                return false;
            }

            bool TryPlanCharacterChunkToNextDelimiter(double maxWidthForChunk, out string selectedText, out double selectedWidth) {
                selectedText = string.Empty;
                selectedWidth = 0D;
                int segmentEnd = breakpoints.FirstOrDefault(point => point > position);
                if (segmentEnd <= position) {
                    segmentEnd = token.Length;
                }

                if (segmentEnd < token.Length && segmentEnd > position + 1 && IsLongTokenDelimiterBreakChar(token[segmentEnd - 1])) {
                    string textWithoutTrailingDelimiter = token.Substring(position, segmentEnd - position - 1);
                    if (textWithoutTrailingDelimiter.Length > 0) {
                        double widthWithoutTrailingDelimiter = MeasurePlannedChunk(textWithoutTrailingDelimiter);
                        if (widthWithoutTrailingDelimiter <= maxWidthForChunk ||
                            CanKeepDelimitedWordSegmentOverSoftLimit(textWithoutTrailingDelimiter, widthWithoutTrailingDelimiter, maxWidthForChunk, allowBoundaryDelimiters: false)) {
                            selectedText = textWithoutTrailingDelimiter;
                            selectedWidth = widthWithoutTrailingDelimiter;
                            return true;
                        }
                    }
                }

                int take = 0;
                double chunkWidth = 0D;
                while (position + take < segmentEnd) {
                    int scalarLength = GetScalarUtf16Length(token, position + take);
                    double candidateWidth = MeasurePlannedChunk(token.Substring(position, take + scalarLength));
                    if ((take > 0 || position == 0 && lineWidth > 0D && !startOnFreshLine) && candidateWidth > maxWidthForChunk) {
                        break;
                    }

                    chunkWidth = candidateWidth;
                    take += scalarLength;
                    if (chunkWidth >= maxWidthForChunk) {
                        break;
                    }
                }

                if (take == 0) {
                    return false;
                }

                selectedText = token.Substring(position, take);
                selectedWidth = chunkWidth;
                return true;
            }

            void TryExtendIdentifierChunkToAvailableWidth(double maxWidthForChunk, ref string selectedText, ref double selectedWidth, ref int selectedBreak) {
                int extendedBreak = selectedBreak;
                string extendedText = selectedText;
                double extendedWidth = selectedWidth;
                while (extendedBreak < token.Length && IsIdentifierContinuationChar(token[extendedBreak])) {
                    int scalarLength = GetScalarUtf16Length(token, extendedBreak);
                    string candidateText = token.Substring(position, extendedBreak + scalarLength - position);
                    double candidateWidth = MeasurePlannedChunk(candidateText);
                    if (candidateWidth > maxWidthForChunk) {
                        break;
                    }

                    extendedBreak += scalarLength;
                    extendedText = candidateText;
                    extendedWidth = candidateWidth;
                }

                if (extendedBreak > selectedBreak) {
                    selectedBreak = extendedBreak;
                    selectedText = extendedText;
                    selectedWidth = extendedWidth;
                }
            }

            bool CanKeepDelimitedWordSegmentOverSoftLimit(string text, double width, double maxWidth, bool allowBoundaryDelimiters) {
                if (maxWidth <= 0D || width <= maxWidth) {
                    return false;
                }

                bool validSegment = allowBoundaryDelimiters
                    ? IsDelimiterBoundedWordGroup(text)
                    : IsDelimitedWordSegment(text);
                if (!validSegment) {
                    return false;
                }

                double widestScalar = 0D;
                for (int offset = 0; offset < text.Length;) {
                    int scalarLength = GetScalarUtf16Length(text, offset);
                    string scalar = text.Substring(offset, scalarLength);
                    widestScalar = Math.Max(widestScalar, MeasureRichText(scalar, font, currentRunNamedFont, runFontSize, baseline, options, currentRunFeatureSettings, currentRunHorizontalTextScaling, currentRunCharacterSpacing));
                    offset += scalarLength;
                }

                double overflow = width - maxWidth;
                return widestScalar > 0D && overflow <= widestScalar + 0.25D;
            }

            static bool IsDelimitedWordSegment(string text) {
                if (text.Length < 5) {
                    return false;
                }

                for (int index = 0; index < text.Length; index++) {
                    if (!char.IsLetter(text[index])) {
                        return false;
                    }
                }

                return true;
            }

            static bool IsDelimiterBoundedWordGroup(string text) {
                if (text.Length < 4 ||
                    !IsLongTokenDelimiterBreakChar(text[0]) ||
                    !IsLongTokenDelimiterBreakChar(text[text.Length - 1])) {
                    return false;
                }

                for (int index = 1; index < text.Length - 1; index++) {
                    if (!char.IsLetter(text[index])) {
                        return false;
                    }
                }

                return true;
            }

            static bool CanExtendDelimitedIdentifierChunk(string token, int breakIndex) {
                char delimiter = token[breakIndex - 1];
                if (breakIndex >= token.Length) {
                    return false;
                }

                return (delimiter == '_' || delimiter == '/' || delimiter == '\\') &&
                    IsIdentifierContinuationChar(token[breakIndex]);
            }

            static bool IsIdentifierContinuationChar(char value) => char.IsLetterOrDigit(value);
        }

        bool TryAppendMultilingualLongToken(
            string token,
            bool bold,
            bool italic,
            bool underline,
            bool strike,
            OfficeIMO.Drawing.OfficeTextDecorationStyle underlineStyle,
            OfficeIMO.Drawing.OfficeTextDecorationStyle strikeStyle,
            PdfColor? color,
            PdfColor? backgroundColor,
            string? uri,
            string? destinationName,
            string? contents,
            PdfStandardFont font,
            double runFontSize,
            PdfTextBaseline baseline,
            double? firstChunkWidth = null) {
            var chunks = PlanLongTokenChunks(token, OfficeTextLineBreaks.GetBreakPositions(token).ToArray(),
                part => MeasureRichText(part, font, currentRunNamedFont, runFontSize, baseline, options,
                    currentRunFeatureSettings, currentRunHorizontalTextScaling, currentRunCharacterSpacing), firstChunkWidth: firstChunkWidth);

            if (chunks == null) {
                return false;
            }

            for (int chunkIndex = 0; chunkIndex < chunks.Count; chunkIndex++) {
                PdfTextTokenChunk chunk = chunks[chunkIndex];
                PrepareLineFrame(RunLineHeight(runFontSize), chunk.Width);
                double boundaryAdvance = MeasureLongTokenBoundary(chunk.Text);
                lines[lines.Count - 1].Add(new RichSeg(chunk.Text, bold, italic, underline, strike, color, backgroundColor, uri, destinationName, contents, font, runFontSize, baseline, chunk.Width, namedFont: currentRunNamedFont, underlineStyle: underlineStyle, strikeStyle: strikeStyle, decorationColor: currentRunDecorationColor, featureSettings: currentRunFeatureSettings, horizontalTextScaling: currentRunHorizontalTextScaling, characterSpacing: currentRunCharacterSpacing, contributesLineBoxDescent: currentRunContributesDescent, formField: currentTextRun?.FormField, footnoteReference: currentTextRun?.FootnoteReference));
                RegisterLineHeight(runFontSize);
                lineWidth += chunk.Width + boundaryAdvance;
                if (chunkIndex + 1 < chunks.Count) {
                    StartNewLine();
                }
            }

            return true;
        }
    }
}
