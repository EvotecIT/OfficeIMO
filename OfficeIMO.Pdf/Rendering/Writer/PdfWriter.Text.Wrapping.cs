using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private static (System.Collections.Generic.List<System.Collections.Generic.List<RichSeg>> Lines, System.Collections.Generic.List<double> LineHeights) WrapRichRuns(System.Collections.Generic.IEnumerable<PdfTextRun> runs, double maxWidthPts, double fontSize, PdfStandardFont baseFont, double lineHeight, double? firstLineWidthPts = null, double tabStopWidth = DefaultParagraphTabStopWidth) =>
        WrapRichRunsCore(runs, maxWidthPts, fontSize, baseFont, lineHeight, firstLineWidthPts, tabStopWidth, options: null);

    private static PdfTabStop[]? NormalizeExplicitTabStops(System.Collections.Generic.IReadOnlyList<PdfTabStop>? tabStops) {
        if (tabStops == null || tabStops.Count == 0) {
            return null;
        }

        return tabStops
            .Where(tabStop => tabStop.Position > 0 && !double.IsNaN(tabStop.Position) && !double.IsInfinity(tabStop.Position))
            .OrderBy(tabStop => tabStop.Position)
            .Select(tabStop => tabStop.Clone())
            .ToArray();
    }

    private static (System.Collections.Generic.List<System.Collections.Generic.List<RichSeg>> Lines, System.Collections.Generic.List<double> LineHeights) WrapRichRunsCore(System.Collections.Generic.IEnumerable<PdfTextRun> runs, double maxWidthPts, double fontSize, PdfStandardFont baseFont, double lineHeight, double? firstLineWidthPts, double tabStopWidth, PdfOptions? options, System.Collections.Generic.IReadOnlyList<PdfTabStop>? tabStops = null) {
        return WrapRichRunsCoreWithFirstLineOrigin(runs, maxWidthPts, fontSize, baseFont, lineHeight, firstLineWidthPts, null, tabStopWidth, options, tabStops);
    }

    private static (System.Collections.Generic.List<System.Collections.Generic.List<RichSeg>> Lines, System.Collections.Generic.List<double> LineHeights) WrapRichRunsWithSpacing(System.Collections.Generic.IEnumerable<PdfTextRun> runs, double maxWidthPts, double fontSize, PdfStandardFont baseFont, double lineHeight, double? firstLineWidthPts, double tabStopWidth, PdfOptions? options, PdfLineSpacing? lineSpacing) =>
        WrapRichRunsCoreWithFirstLineOrigin(runs, maxWidthPts, fontSize, baseFont, lineHeight, firstLineWidthPts, null, tabStopWidth, options, lineSpacing: lineSpacing);

    private static (System.Collections.Generic.List<System.Collections.Generic.List<RichSeg>> Lines, System.Collections.Generic.List<double> LineHeights) WrapRichRunsCoreWithFirstLineOrigin(System.Collections.Generic.IEnumerable<PdfTextRun> runs, double maxWidthPts, double fontSize, PdfStandardFont baseFont, double lineHeight, double? firstLineWidthPts, double? firstLineOriginOffsetPts, double tabStopWidth, PdfOptions? options, System.Collections.Generic.IReadOnlyList<PdfTabStop>? tabStops = null, Func<int, double, double, double, (double Width, double Origin, double Gap)>? lineLayout = null, double? minimumLineHeight = null, PdfLineSpacing? lineSpacing = null) {
        bool preserveWhitespace = options?.PreserveTextWhitespace == true;
        System.Collections.Generic.IReadOnlyList<PdfTextRun> effectiveRuns = NormalizeFallbackRuns(runs, baseFont, options);
        var wordContinuations = MeasureRichWordContinuations(effectiveRuns, baseFont, fontSize, options);
        PdfTabStop[]? explicitTabStops = NormalizeExplicitTabStops(tabStops);
        var lines = new System.Collections.Generic.List<System.Collections.Generic.List<RichSeg>> { new RichLine() };
        var heights = new System.Collections.Generic.List<double>();
        double lineWidth = 0;
        double pendingLeadingAdvance = 0;
        bool pendingLeadingSeparator = false;
        OfficeIMO.Drawing.OfficeTextDecorationStyle pendingLeadingUnderlineStyle = OfficeIMO.Drawing.OfficeTextDecorationStyle.None;
        PdfColor? pendingLeadingDecorationColor = null;
        double pendingLeadingDecorationFontSize = 0;
        double pendingLeadingDecorationTextRise = 0;
        bool pendingLeadingIsExpandable = true;
        bool pendingLeadingIsTab = false;
        PdfTabAlignment pendingLeadingTabAlignment = PdfTabAlignment.Left;
        PdfTabLeaderStyle pendingLeadingTabLeader = PdfTabLeaderStyle.None;
        PdfTabStop? pendingLeadingTabStop = null;
        int nextExplicitTabStopIndex = 0;
        double lineHeightRatio = fontSize > 0 ? lineHeight / fontSize : 1.2D;
        double RunLineHeight(double size) => lineSpacing?.GetAdvance(size) ?? size * lineHeightRatio;
        double completedHeight = 0;
        double currentContentHeight = lineHeight;
        var currentFrame = lineLayout?.Invoke(0, completedHeight, lineHeight, 0);
        double currentLineHeight = Math.Max(lineHeight, minimumLineHeight ?? 0) + (currentFrame?.Gap ?? 0);
        double currentMinimumWidth = 0;
        PdfNamedFontFace? currentRunNamedFont = null;
        double currentRunAscent = 0;
        double currentRunDescent = 0;
        double currentLineTextAscent = 0;
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
        double CurrentMaxWidth() => currentFrame?.Width ?? (lines.Count == 1 ? firstLineWidthPts ?? maxWidthPts : maxWidthPts);
        double CurrentLineOriginOffset() => currentFrame?.Origin ?? (lines.Count == 1 ? firstLineOriginOffsetPts ?? 0D : 0D);
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
            }
            currentLineHasInline |= currentRunIsInline;
            currentLineHeight = Math.Max(currentLineHeight,
                CombinedContentHeight(0, includeCurrentRun: false) + (currentFrame?.Gap ?? 0));
        }

        void StartNewLine() {
            if (lines.Count >= MaximumTextLayoutLines)
                throw new System.IO.InvalidDataException("PDF paragraph layout exceeds the 100,000-line limit.");
            heights.Add(currentLineHeight);
            completedHeight += currentLineHeight;
            lines.Add(new RichLine());
            lineWidth = 0;
            currentLineTextAscent = 0;
            currentLineTextDescent = 0;
            currentLineInlineAscent = 0;
            currentLineInlineDescent = 0;
            currentLineHasInline = false;
            currentContentHeight = lineHeight;
            currentFrame = lineLayout?.Invoke(lines.Count - 1, completedHeight, currentContentHeight, 0);
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
            var frame = lineLayout(lines.Count - 1, completedHeight, height, requiredWidth);
            if (lines[lines.Count - 1].Count > 0 &&
                (lineWidth > frame.Width + 0.001 || minimumWidth > 0 && frame.Gap > (oldFrame?.Gap ?? 0) + 0.001)) {
                lineLayout(lines.Count - 1, completedHeight, oldHeight, currentMinimumWidth);
                currentFrame = oldFrame;
                StartNewLine();
                height = Math.Max(lineHeight, CombinedContentHeight(incomingContentHeight));
                requiredWidth = minimumWidth;
                frame = lineLayout(lines.Count - 1, completedHeight, height, requiredWidth);
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
            double inlineAscent = Math.Max(currentLineInlineAscent,
                includeCurrentRun && currentRunIsInline ? currentRunAscent : 0);
            double inlineDescent = Math.Max(currentLineInlineDescent,
                includeCurrentRun && currentRunIsInline ? currentRunDescent : 0);
            double combined = Math.Max(textAscent, inlineAscent) + Math.Max(textDescent, inlineDescent);
            if (lineSpacing?.FontLineBoxMultiplier is not double natural)
                return hasInline ? Math.Max(contentHeight, combined) : contentHeight;

            // Imported spacing scales the visible fonts' natural advance.
            // Horizontal list spacers must not suppress that multiplier.
            double textAdvance = textAscent + textDescent;
            if (lineSpacing.Rule == PdfLineSpacingRule.Multiple)
                textAdvance *= lineSpacing.Value / natural;
            contentHeight = Math.Max(contentHeight, textAdvance);
            // Inline objects retain their unscaled bounds when they extend
            // beyond the text line box; a spacer inside it adds no height.
            return hasInline && (inlineAscent > textAscent || inlineDescent > textDescent)
                ? Math.Max(contentHeight, combined) : contentHeight;
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

        void ResolvePendingLeadingTabForCurrentLine(double followingTextWidth, double spaceW, string followingText, PdfStandardFont followingFont, double followingFontSize, PdfTextBaseline followingBaseline) {
            PdfTabAlignment fallbackAlignment = pendingLeadingTabAlignment;
            PdfTabLeaderStyle fallbackLeader = pendingLeadingTabLeader;
            PdfTabStop? explicitTabStop = ResolveNextExplicitTabStop();
            pendingLeadingTabAlignment = explicitTabStop?.Alignment ?? fallbackAlignment;
            pendingLeadingTabLeader = explicitTabStop?.Leader ?? fallbackLeader;
            pendingLeadingTabStop = explicitTabStop;
            pendingLeadingAdvance = CalculateTabAdvance(lineWidth, followingTextWidth, spaceW, pendingLeadingTabAlignment, tabStopWidth, followingText, followingFont, followingFontSize, followingBaseline, options, CurrentMaxWidth(), pendingLeadingTabStop, CurrentLineOriginOffset(), currentRunNamedFont, currentRunFeatureSettings, currentRunHorizontalTextScaling, currentRunCharacterSpacing);
        }

        void ResetPendingLeading() {
            pendingLeadingAdvance = 0;
            pendingLeadingSeparator = false;
            pendingLeadingUnderlineStyle = OfficeIMO.Drawing.OfficeTextDecorationStyle.None;
            pendingLeadingDecorationColor = null;
            pendingLeadingDecorationFontSize = 0;
            pendingLeadingDecorationTextRise = 0;
            pendingLeadingIsExpandable = true;
            pendingLeadingIsTab = false;
            pendingLeadingTabAlignment = PdfTabAlignment.Left;
            pendingLeadingTabLeader = PdfTabLeaderStyle.None;
            pendingLeadingTabStop = null;
        }

        void RevalidateLeadingTabFrame(double contentHeight, double followingWidth, double spaceWidth, string followingText, PdfStandardFont font, double size, PdfTextBaseline baseline) {
            if (lineLayout == null || !pendingLeadingIsTab || lines[lines.Count - 1].Count != 0) return;
            PrepareLineFrame(contentHeight, pendingLeadingAdvance + followingWidth);
            // The selected interval may change the tab's origin. Retain its resolved stop,
            // then validate the complete advance again rather than consuming another stop.
            pendingLeadingAdvance = CalculateTabAdvance(lineWidth, followingWidth, spaceWidth, pendingLeadingTabAlignment, tabStopWidth, followingText, font, size, baseline, options, CurrentMaxWidth(), pendingLeadingTabStop, CurrentLineOriginOffset(), currentRunNamedFont, currentRunFeatureSettings, currentRunHorizontalTextScaling, currentRunCharacterSpacing);
            PrepareLineFrame(contentHeight, pendingLeadingAdvance + followingWidth);
        }

        void SetPendingSeparator(bool hadTab, double spaceW, PdfTabAlignment tabAlignment, PdfTabLeaderStyle tabLeader) {
            pendingLeadingSeparator = true;
            if (!hadTab) {
                pendingLeadingAdvance = preserveWhitespace ? pendingLeadingAdvance + spaceW : spaceW;
                pendingLeadingUnderlineStyle = currentRunUnderlineStyle;
                pendingLeadingDecorationColor = currentRunUnderlineColor;
                pendingLeadingDecorationFontSize = currentRunDecorationFontSize;
                pendingLeadingDecorationTextRise = currentRunDecorationTextRise;
                pendingLeadingIsExpandable = !preserveWhitespace;
                pendingLeadingIsTab = false;
                pendingLeadingTabAlignment = PdfTabAlignment.Left;
                pendingLeadingTabLeader = PdfTabLeaderStyle.None;
                pendingLeadingTabStop = null;
                return;
            }

            PdfTabStop? explicitTabStop = ResolveNextExplicitTabStop();
            pendingLeadingTabAlignment = explicitTabStop?.Alignment ?? tabAlignment;
            pendingLeadingTabLeader = explicitTabStop?.Leader ?? tabLeader;
            pendingLeadingTabStop = explicitTabStop;
            pendingLeadingAdvance = CalculateTabAdvance(lineWidth, 0D, spaceW, pendingLeadingTabAlignment, tabStopWidth, options: options, maxWidth: CurrentMaxWidth(), explicitTabStop: pendingLeadingTabStop, lineOriginOffset: CurrentLineOriginOffset(), followingNamedFont: currentRunNamedFont, featureSettings: currentRunFeatureSettings, horizontalTextScaling: currentRunHorizontalTextScaling, characterSpacing: currentRunCharacterSpacing);
            pendingLeadingUnderlineStyle = OfficeIMO.Drawing.OfficeTextDecorationStyle.None;
            pendingLeadingIsExpandable = false;
            pendingLeadingIsTab = true;
        }

        void MarkCurrentLineHardBreak(RichSeg breakSegment) =>
            lines[lines.Count - 1].Add(breakSegment);

        for (int runIndex = 0; runIndex < effectiveRuns.Count; runIndex++) {
            PdfTextRun run = effectiveRuns[runIndex];
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
            currentRunIsInline = run.InlineElement != null;
            GetRichRunLineMetrics(fontForRun, currentRunNamedFont, runFontSize, baseline, options, lineSpacing,
                out currentRunAscent, out currentRunDescent);
            double spaceW = text.IndexOfAny(SoftLineSplitChars) >= 0
                ? MeasureRichText(" ", fontForRun, currentRunNamedFont, runFontSize, baseline, options, currentRunFeatureSettings, currentRunHorizontalTextScaling, currentRunCharacterSpacing)
                : 0D;
            if (run.InlineElement != null) {
                PdfInlineElement inlineElement = run.InlineElement;
                currentRunAscent = Math.Max(0D, inlineElement.BaselineOffset + inlineElement.Height);
                currentRunDescent = Math.Max(0D, -inlineElement.BaselineOffset);
                double inlineHeight = Math.Max(GetAscenderForOptions(baseFont, fontSize, options), inlineElement.BaselineOffset + inlineElement.Height)
                    + Math.Max(GetDescenderForOptions(baseFont, fontSize, options), -inlineElement.BaselineOffset);
                if (lineSpacing?.FontLineBoxMultiplier != null)
                    inlineHeight = Math.Max(lineHeight, currentRunAscent + currentRunDescent);
                PrepareLineFrame(inlineHeight, inlineElement.Width);
                double currentMaxWidth = CurrentMaxWidth();
                if (inlineElement.Width > currentMaxWidth + 0.001D) {
                    throw new ArgumentException("Inline element width exceeds the available paragraph width.");
                }

                if (pendingLeadingIsTab) {
                    ResolvePendingLeadingTabForCurrentLine(inlineElement.Width, spaceW, string.Empty, fontForRun, runFontSize, baseline);
                    RevalidateLeadingTabFrame(inlineHeight, inlineElement.Width, spaceW, string.Empty, fontForRun, runFontSize, baseline);
                    currentMaxWidth = CurrentMaxWidth();
                }

                List<RichSeg> currentLine = lines[lines.Count - 1];
                double leadingAdvance = currentLine.Count > 0 || pendingLeadingIsTab ? pendingLeadingAdvance : 0D;
                if (currentLine.Count > 0 && lineWidth + leadingAdvance + inlineElement.Width > currentMaxWidth + 0.001D) {
                    if (pendingLeadingAdvance > 0D) {
                        MarkRichLineTextSeparator(currentLine);
                    }

                    StartNewLine();
                    PrepareLineFrame(inlineHeight, inlineElement.Width);
                    currentLine = lines[lines.Count - 1];
                    if (pendingLeadingIsTab) {
                        ResolvePendingLeadingTabForCurrentLine(inlineElement.Width, spaceW, string.Empty, fontForRun, runFontSize, baseline);
                        RevalidateLeadingTabFrame(inlineHeight, inlineElement.Width, spaceW, string.Empty, fontForRun, runFontSize, baseline);
                    }

                    leadingAdvance = pendingLeadingIsTab ? pendingLeadingAdvance : 0D;
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
                    featureSettings: currentRunFeatureSettings, horizontalTextScaling: currentRunHorizontalTextScaling, characterSpacing: currentRunCharacterSpacing,
                    leadingTabStop: pendingLeadingIsTab ? pendingLeadingTabStop : null,
                    leadingUnderlineStyle: leadingAdvance > 0D ? pendingLeadingUnderlineStyle : OfficeIMO.Drawing.OfficeTextDecorationStyle.None,
                    leadingDecorationColor: pendingLeadingDecorationColor,
                    leadingDecorationFontSize: pendingLeadingDecorationFontSize,
                    leadingDecorationTextRise: pendingLeadingDecorationTextRise,
                    leadingIsTab: pendingLeadingIsTab, leadingTabAlignment: pendingLeadingTabAlignment));
                lineWidth += leadingAdvance + inlineElement.Width;
                RegisterLineMetrics();
                ResetPendingLeading();
                continue;
            }

            int idx = 0;
            while (idx < text.Length) {
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
                double wordWidth = tokenW + continuation.Width;
                double wordFontSize = Math.Max(runFontSize, continuation.FontSize);
                if (token.Length > 0) PrepareLineFrame(RunLineHeight(wordFontSize),
                    continuation.Width > 0D && wordWidth <= maxWidthPts ? wordWidth : 0D);
                if (preserveWhitespace && !pendingLeadingIsTab && pendingLeadingAdvance > 0 && (token.Length > 0 || hadNewline)) {
                    // Literal spacing consumes line capacity just like visible text. Keep
                    // its advance on the line it occupies, including completely blank lines.
                    while (pendingLeadingAdvance > 0.001D) {
                        double available = CurrentMaxWidth() - lineWidth;
                        if (available <= 0.001D) {
                            StartNewLine();
                            PrepareLineFrame(RunLineHeight(runFontSize));
                            available = CurrentMaxWidth();
                            if (available <= 0.001D) throw new InvalidOperationException("No width is available for preserved text spacing.");
                        }
                        double advance = Math.Min(available, pendingLeadingAdvance);
                        lines[lines.Count - 1].Add(new RichSeg(string.Empty, bold, italic, underline, strike, color, backgroundColor,
                            uri, destinationName, contents, fontForRun, runFontSize, baseline, 0,
                            leadingSpace: true, leadingAdvance: advance, leadingSpaceIsExpandable: false,
                            namedFont: currentRunNamedFont, featureSettings: currentRunFeatureSettings, horizontalTextScaling: currentRunHorizontalTextScaling, characterSpacing: currentRunCharacterSpacing));
                        RegisterLineHeight(runFontSize);
                        lineWidth += advance;
                        pendingLeadingAdvance -= advance;
                    }
                    ResetPendingLeading();
                }
                var lastLine = lines[lines.Count - 1];
                double needed = lastLine.Count == 0 && !preserveWhitespace ? tokenW : pendingLeadingAdvance + tokenW;
                double currentMaxWidth = CurrentMaxWidth();

                if (tokenW > currentMaxWidth) {
                    if (lastLine.Count > 0) { StartNewLine(); lastLine = lines[lines.Count - 1]; }
                    ResetPendingLeading();
                    if (TryAppendSoftLineBreakLongToken(token, bold, italic, underline, strike, underlineStyle, strikeStyle, color, backgroundColor, uri, destinationName, contents, fontForRun, runFontSize, baseline)) {
                        if (hadNewline) {
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
                    while (pos < token.Length) {
                        int take = 0;
                        double chunkW = 0;
                        PrepareLineFrame(RunLineHeight(runFontSize));
                        currentMaxWidth = CurrentMaxWidth();
                        while (pos + take < token.Length) {
                            int scalarLength = GetScalarUtf16Length(token, pos + take);
                            string scalar = token.Substring(pos + take, scalarLength);
                            double charW = MeasureRichText(scalar, fontForRun, currentRunNamedFont, runFontSize, baseline, options, currentRunFeatureSettings, currentRunHorizontalTextScaling, currentRunCharacterSpacing);
                            if (take > 0 && chunkW + charW > currentMaxWidth) {
                                break;
                            }

                            chunkW += charW;
                            take += scalarLength;
                            if (chunkW >= currentMaxWidth) {
                                break;
                            }
                        }

                        if (take == 0) {
                            take = GetScalarUtf16Length(token, pos);
                            chunkW = MeasureRichText(token.Substring(pos, take), fontForRun, currentRunNamedFont, runFontSize, baseline, options, currentRunFeatureSettings, currentRunHorizontalTextScaling, currentRunCharacterSpacing);
                        }

                        string chunk = token.Substring(pos, take);
                        double measuredChunkWidth = MeasureRichText(chunk, fontForRun, currentRunNamedFont, runFontSize, baseline, options, currentRunFeatureSettings, currentRunHorizontalTextScaling, currentRunCharacterSpacing);
                        lastLine.Add(new RichSeg(chunk, bold, italic, underline, strike, color, backgroundColor, uri, destinationName, contents, fontForRun, runFontSize, baseline, measuredChunkWidth, namedFont: currentRunNamedFont, underlineStyle: underlineStyle, strikeStyle: strikeStyle, decorationColor: currentRunDecorationColor, featureSettings: currentRunFeatureSettings, horizontalTextScaling: currentRunHorizontalTextScaling, characterSpacing: currentRunCharacterSpacing));
                        RegisterLineHeight(runFontSize);
                        lineWidth += chunkW;
                        pos += take;
                        if (pos < token.Length) { StartNewLine(); lastLine = lines[lines.Count - 1]; }
                    }
                    if (hadNewline) {
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
                    pendingLeadingAdvance = CalculateTabAdvance(lineWidth, tokenW, spaceW, pendingLeadingTabAlignment, tabStopWidth, token, fontForRun, runFontSize, baseline, options, CurrentMaxWidth(), pendingLeadingTabStop, CurrentLineOriginOffset(), currentRunNamedFont, currentRunFeatureSettings, currentRunHorizontalTextScaling, currentRunCharacterSpacing);
                    RevalidateLeadingTabFrame(RunLineHeight(runFontSize), tokenW, spaceW, token, fontForRun, runFontSize, baseline);
                    currentMaxWidth = CurrentMaxWidth();
                }
                // Formatting boundaries do not introduce a word-break opportunity.
                // Reserve the rest of a word spanning runs when it fits a fresh line.
                double wrappingWidth = wordWidth <= Math.Min(currentMaxWidth, maxWidthPts) ? wordWidth : tokenW;
                needed = lastLine.Count == 0
                    ? (pendingLeadingIsTab || preserveWhitespace ? pendingLeadingAdvance + wrappingWidth : wrappingWidth)
                    : pendingLeadingAdvance + wrappingWidth;
                if (lineWidth + needed > currentMaxWidth && lastLine.Count > 0) {
                    if (pendingLeadingAdvance > 0D) {
                        MarkRichLineTextSeparator(lastLine);
                    }

                    StartNewLine();
                    // The next line can enter a narrower floating frame. A pending word
                    // that fit the previous frame must obtain enough space again.
                    PrepareLineFrame(RunLineHeight(wordFontSize), wrappingWidth);
                    if (token.Length > 0 && pendingLeadingIsTab) {
                        ResolvePendingLeadingTabForCurrentLine(tokenW, spaceW, token, fontForRun, runFontSize, baseline);
                        RevalidateLeadingTabFrame(RunLineHeight(runFontSize), tokenW, spaceW, token, fontForRun, runFontSize, baseline);
                    }
                }
                if (token.Length > 0) {
                    bool needsLeadingSpace = pendingLeadingSeparator && (lastLine.Count > 0 || pendingLeadingIsTab || preserveWhitespace);
                    double leadingAdvance = needsLeadingSpace ? pendingLeadingAdvance : 0;
                    double segmentWidth = tokenW + leadingAdvance;
                    var segmentLeader = needsLeadingSpace ? pendingLeadingTabLeader : PdfTabLeaderStyle.None;
                    lines[lines.Count - 1].Add(new RichSeg(token, bold, italic, underline, strike, color, backgroundColor, uri, destinationName, contents, fontForRun, runFontSize, baseline, tokenW, needsLeadingSpace, leadingAdvance, pendingLeadingIsExpandable, segmentLeader, namedFont: currentRunNamedFont, underlineStyle: underlineStyle, strikeStyle: strikeStyle, decorationColor: currentRunDecorationColor, featureSettings: currentRunFeatureSettings, horizontalTextScaling: currentRunHorizontalTextScaling, characterSpacing: currentRunCharacterSpacing, leadingTabStop: pendingLeadingIsTab ? pendingLeadingTabStop : null, leadingUnderlineStyle: needsLeadingSpace ? pendingLeadingUnderlineStyle : OfficeIMO.Drawing.OfficeTextDecorationStyle.None, leadingDecorationColor: pendingLeadingDecorationColor, leadingDecorationFontSize: pendingLeadingDecorationFontSize, leadingDecorationTextRise: pendingLeadingDecorationTextRise, leadingIsTab: pendingLeadingIsTab, leadingTabAlignment: pendingLeadingTabAlignment));
                    RegisterLineHeight(runFontSize);
                    lineWidth += segmentWidth;
                    ResetPendingLeading();
                }
                if (hadNewline) {
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
        if (lines.Count > 0 && lines[lines.Count - 1].Count == 0) { lines.RemoveAt(lines.Count - 1); }
        if (heights.Count < lines.Count) heights.Add(currentLineHeight);
        foreach (RichLine line in lines)
            line.BaselineOffset = ResolveRichLineBaselineOffset(line, baseFont, fontSize, options, lineSpacing);
        return (lines, heights);

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
            var chunks = TryBuildSoftLineBreakTokenChunks(
                token,
                options,
                part => MeasureRichText(part, font, currentRunNamedFont, runFontSize, baseline, options, currentRunFeatureSettings, currentRunHorizontalTextScaling, currentRunCharacterSpacing),
                CurrentMaxWidth(),
                maxWidthPts);

            if (chunks == null) {
                return false;
            }

            for (int chunkIndex = 0; chunkIndex < chunks.Count; chunkIndex++) {
                PdfTextTokenChunk chunk = chunks[chunkIndex];
                PrepareLineFrame(RunLineHeight(runFontSize), chunk.Width);
                lines[lines.Count - 1].Add(new RichSeg(chunk.Text, bold, italic, underline, strike, color, backgroundColor, uri, destinationName, contents, font, runFontSize, baseline, chunk.Width, namedFont: currentRunNamedFont, underlineStyle: underlineStyle, strikeStyle: strikeStyle, decorationColor: currentRunDecorationColor, featureSettings: currentRunFeatureSettings, horizontalTextScaling: currentRunHorizontalTextScaling, characterSpacing: currentRunCharacterSpacing));
                RegisterLineHeight(runFontSize);
                lineWidth += chunk.Width;
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
            int[] breakpoints = GetValidHyphenationBreakpoints(token, options);
            if (breakpoints.Length == 0) {
                return false;
            }

            int position = 0;
            var plannedChunks = new System.Collections.Generic.List<(string Text, double Width)>();
            while (position < token.Length) {
                int selectedBreak = -1;
                string selectedText = string.Empty;
                double selectedWidth = 0D;
                double maxWidthForChunk = plannedChunks.Count == 0 ? CurrentMaxWidth() : maxWidthPts;
                int[] candidates = breakpoints
                    .Where(point => point > position)
                    .Concat(new[] { token.Length })
                    .Distinct()
                    .OrderBy(point => point)
                    .ToArray();

                foreach (int candidate in candidates) {
                    bool finalChunk = candidate >= token.Length;
                    string chunkText = token.Substring(position, candidate - position);
                    if (!finalChunk) {
                        chunkText += "-";
                    }

                    if (chunkText.Length == 0) {
                        continue;
                    }

                    double chunkWidth = MeasureRichText(chunkText, font, currentRunNamedFont, runFontSize, baseline, options, currentRunFeatureSettings, currentRunHorizontalTextScaling, currentRunCharacterSpacing);
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

                if (selectedBreak <= position || selectedText.Length == 0) {
                    return false;
                }

                plannedChunks.Add((selectedText, selectedWidth));
                position = selectedBreak;
            }

            for (int chunkIndex = 0; chunkIndex < plannedChunks.Count; chunkIndex++) {
                (string selectedText, double selectedWidth) = plannedChunks[chunkIndex];
                double measuredWidth = MeasureRichText(selectedText, font, currentRunNamedFont, runFontSize, baseline, options, currentRunFeatureSettings, currentRunHorizontalTextScaling, currentRunCharacterSpacing);
                PrepareLineFrame(RunLineHeight(runFontSize), measuredWidth);
                lines[lines.Count - 1].Add(new RichSeg(selectedText, bold, italic, underline, strike, color, backgroundColor, uri, destinationName, contents, font, runFontSize, baseline, measuredWidth, namedFont: currentRunNamedFont, underlineStyle: underlineStyle, strikeStyle: strikeStyle, decorationColor: currentRunDecorationColor, featureSettings: currentRunFeatureSettings, horizontalTextScaling: currentRunHorizontalTextScaling, characterSpacing: currentRunCharacterSpacing));
                RegisterLineHeight(runFontSize);
                lineWidth += selectedWidth;
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
            var plannedChunks = new System.Collections.Generic.List<(string Text, double Width)>();
            while (position < token.Length) {
                int selectedBreak = -1;
                string selectedText = string.Empty;
                double selectedWidth = 0D;
                double maxWidthForChunk = plannedChunks.Count == 0 ? CurrentMaxWidth() : maxWidthPts;
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

                    double chunkWidth = MeasureRichText(chunkText, font, currentRunNamedFont, runFontSize, baseline, options, currentRunFeatureSettings, currentRunHorizontalTextScaling, currentRunCharacterSpacing);
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
                        return false;
                    }

                    selectedBreak = position + selectedText.Length;
                }

                plannedChunks.Add((selectedText, selectedWidth));
                position = selectedBreak;
            }

            for (int chunkIndex = 0; chunkIndex < plannedChunks.Count; chunkIndex++) {
                (string selectedText, double selectedWidth) = plannedChunks[chunkIndex];
                double measuredWidth = MeasureRichText(selectedText, font, currentRunNamedFont, runFontSize, baseline, options, currentRunFeatureSettings, currentRunHorizontalTextScaling, currentRunCharacterSpacing);
                PrepareLineFrame(RunLineHeight(runFontSize), measuredWidth);
                lines[lines.Count - 1].Add(new RichSeg(selectedText, bold, italic, underline, strike, color, backgroundColor, uri, destinationName, contents, font, runFontSize, baseline, measuredWidth, namedFont: currentRunNamedFont, underlineStyle: underlineStyle, strikeStyle: strikeStyle, decorationColor: currentRunDecorationColor, featureSettings: currentRunFeatureSettings, horizontalTextScaling: currentRunHorizontalTextScaling, characterSpacing: currentRunCharacterSpacing));
                RegisterLineHeight(runFontSize);
                lineWidth += selectedWidth;
                if (chunkIndex < plannedChunks.Count - 1) {
                    StartNewLine();
                }
            }

            return true;

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

                double candidateWidth = MeasureRichText(candidate, font, currentRunNamedFont, runFontSize, baseline, options, currentRunFeatureSettings, currentRunHorizontalTextScaling, currentRunCharacterSpacing);
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
                        double widthWithoutTrailingDelimiter = MeasureRichText(textWithoutTrailingDelimiter, font, currentRunNamedFont, runFontSize, baseline, options, currentRunFeatureSettings, currentRunHorizontalTextScaling, currentRunCharacterSpacing);
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
                    string scalar = token.Substring(position + take, scalarLength);
                    double scalarWidth = MeasureRichText(scalar, font, currentRunNamedFont, runFontSize, baseline, options, currentRunFeatureSettings, currentRunHorizontalTextScaling, currentRunCharacterSpacing);
                    if (take > 0 && chunkWidth + scalarWidth > maxWidthForChunk) {
                        break;
                    }

                    chunkWidth += scalarWidth;
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
                    double candidateWidth = MeasureRichText(candidateText, font, currentRunNamedFont, runFontSize, baseline, options, currentRunFeatureSettings, currentRunHorizontalTextScaling, currentRunCharacterSpacing);
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
            PdfTextBaseline baseline) {
            var chunks = TryBuildMultilingualTokenChunks(
                token,
                part => MeasureRichText(part, font, currentRunNamedFont, runFontSize, baseline, options, currentRunFeatureSettings, currentRunHorizontalTextScaling, currentRunCharacterSpacing),
                CurrentMaxWidth(),
                maxWidthPts);

            if (chunks == null) {
                return false;
            }

            for (int chunkIndex = 0; chunkIndex < chunks.Count; chunkIndex++) {
                PdfTextTokenChunk chunk = chunks[chunkIndex];
                PrepareLineFrame(RunLineHeight(runFontSize), chunk.Width);
                lines[lines.Count - 1].Add(new RichSeg(chunk.Text, bold, italic, underline, strike, color, backgroundColor, uri, destinationName, contents, font, runFontSize, baseline, chunk.Width, namedFont: currentRunNamedFont, underlineStyle: underlineStyle, strikeStyle: strikeStyle, decorationColor: currentRunDecorationColor, featureSettings: currentRunFeatureSettings, horizontalTextScaling: currentRunHorizontalTextScaling, characterSpacing: currentRunCharacterSpacing));
                RegisterLineHeight(runFontSize);
                lineWidth += chunk.Width;
                if (chunkIndex + 1 < chunks.Count) {
                    StartNewLine();
                }
            }

            return true;
        }
    }
}
