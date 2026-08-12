using System.Globalization;
using System.Text;
using AngleSharp.Dom;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private HtmlInlineLayout LayoutInlineNodes(
        IEnumerable<INode> nodes,
        double width,
        HtmlRenderBoxStyle parentStyle,
        int depth,
        string? prefix,
        IElement? generatedContentOwner,
        int skipLogicalCharacters = 0) {
        var runs = new List<HtmlInlineRun>();
        IElement? formattingContainer = generatedContentOwner ?? nodes.FirstOrDefault()?.ParentElement;
        if (!string.IsNullOrEmpty(prefix)) {
            runs.Add(new HtmlInlineRun(prefix!, parentStyle, null, "list-marker"));
        }

        double? containingHeight = ResolveContainingBlockHeight(parentStyle);
        if (generatedContentOwner != null) {
            AddGeneratedInlineRun(generatedContentOwner, HtmlPseudoElementKind.Before, width, containingHeight, parentStyle, null, 0D, 0D, runs);
        }

        foreach (INode node in nodes) {
            CollectInlineRuns(node, width, containingHeight, parentStyle, null, depth, 0D, 0D, runs);
        }

        if (generatedContentOwner != null) {
            AddGeneratedInlineRun(generatedContentOwner, HtmlPseudoElementKind.After, width, containingHeight, parentStyle, null, 0D, 0D, runs);
        }

        runs = ApplyScopedFontFallbacks(runs);

        if (formattingContainer != null && HtmlRenderHeading.TryGetLevel(parentStyle.SemanticRole, out _)) {
            int semanticNodeId = GetSemanticNodeId(formattingContainer);
            foreach (HtmlInlineRun run in runs) run.AssignSemanticNode(parentStyle.SemanticRole, semanticNodeId);
        }

        return LayoutInlineRuns(runs, width, parentStyle, formattingContainer, skipLogicalCharacters);
    }

    private List<HtmlInlineRun> ApplyScopedFontFallbacks(IEnumerable<HtmlInlineRun> sourceRuns) {
        var resolvedRuns = new List<HtmlInlineRun>();
        foreach (HtmlInlineRun run in sourceRuns) {
            if (run.Text.Length == 0 || run.AtomicBlock != null || run.FloatingBlock != null || run.PositionedMarkerElement != null) {
                resolvedRuns.Add(run);
                continue;
            }

            IReadOnlyList<OfficeFontFallbackRun> fallbacks = _fonts.PlanFallbackRuns(run.Text, run.Style.Font.FamilyName, run.Style.Font.Style);
            string shapedText = OfficeArabicTextShaper.Shape(fallbacks.Count == 1 ? fallbacks[0].Text : run.Text);
            if (fallbacks.Count == 1
                && string.Equals(fallbacks[0].Text, run.Text, StringComparison.Ordinal)
                && string.Equals(fallbacks[0].FamilyName, run.Style.Font.FamilyName, StringComparison.Ordinal)
                && string.Equals(shapedText, run.Text, StringComparison.Ordinal)) {
                resolvedRuns.Add(run);
                continue;
            }

            foreach (OfficeFontFallbackRun fallback in fallbacks) {
                HtmlRenderBoxStyle style = run.Style.Clone();
                style.Font = style.Font.WithFamilyName(fallback.FamilyName);
                resolvedRuns.Add(new HtmlInlineRun(
                    OfficeArabicTextShaper.Shape(fallback.Text),
                    style,
                    run.LinkUri,
                    run.Source,
                    run.PaintOffsetX,
                    run.PaintOffsetY,
                    run.OwnerElement,
                    run.PositionedMarkerElement,
                    fallback.Text));
            }
        }

        return resolvedRuns;
    }

    private void CollectInlineRuns(
        INode node,
        double width,
        double? containingHeight,
        HtmlRenderBoxStyle inheritedStyle,
        string? inheritedLink,
        int depth,
        double inheritedPaintOffsetX,
        double inheritedPaintOffsetY,
        ICollection<HtmlInlineRun> runs) {
        if (depth > _options.MaxLayoutDepth) {
            if (node is IElement limitedElement) EnsureDepth(depth, limitedElement);
            throw new InvalidOperationException("HTML inline layout exceeded the configured maximum depth.");
        }

        if (node is IText textNode) {
            if (textNode.Data.Length > 0) {
                ReportUnsupportedBidi(textNode, inheritedStyle);
                runs.Add(new HtmlInlineRun(ApplyTextTransform(textNode.Data, inheritedStyle.TextTransform), inheritedStyle, inheritedLink, inheritedStyle.SemanticRole, inheritedPaintOffsetX, inheritedPaintOffsetY, textNode.ParentElement));
            }

            return;
        }

        if (!(node is IElement element) || ShouldSkipElement(element)) return;
        string tag = element.TagName.ToLowerInvariant();
        if (tag == "br") {
            runs.Add(new HtmlInlineRun("\u2028", inheritedStyle, inheritedLink, HtmlRenderStyleResolver.DescribeSource(element), inheritedPaintOffsetX, inheritedPaintOffsetY, element));
            return;
        }

        HtmlRenderBoxStyle style = _styleResolver.Resolve(element, width, inheritedStyle);
        _layoutStyles[element] = style.Clone();
        if (style.Display == "none") return;
        ReportUnsupportedFloatValues(element, style);
        ReportUnsupportedOverflowValues(element, style);
        ReportUnsupportedMultiColumnValues(element, style);
        string? link = inheritedLink;
        if (tag == "a") {
            link = ResolveSafeLink(element.GetAttribute("href"), element);
        }
        if ((style.Position == "relative" || style.Position == "sticky") && style.ZIndex != "auto") {
            _inlineStackingElements.Add(element);
        }
        if (style.Position == "absolute" || style.Position == "fixed") {
            RegisterOutOfFlowElement(element.ParentElement ?? element, element, style, inheritedStyle, depth);
            runs.Add(new HtmlInlineRun(
                string.Empty,
                style,
                null,
                HtmlRenderStyleResolver.DescribeSource(element),
                inheritedPaintOffsetX,
                inheritedPaintOffsetY,
                element.ParentElement,
                element));
            return;
        }
        if (style.FloatSide != "none") {
            AddFloatingRun(element, width, inheritedStyle, depth, style, link, runs);
            return;
        }

        if (IsFormControlElement(tag)) {
            if (!string.IsNullOrWhiteSpace(style.StringSet)) {
                runs.Add(new HtmlInlineRun(
                    element,
                    style,
                    HtmlRenderStyleResolver.DescribeSource(element)));
            }
            ResolvePositionPaintOffset(style, width, containingHeight, HtmlRenderStyleResolver.DescribeSource(element), out double controlOffsetX, out double controlOffsetY);
            double outerWidth = ResolveFormControlOuterWidth(element, style, width);
            HtmlRenderFlowBlock control = LayoutFormControl(element, outerWidth, style);
            control = ApplyElementPaintEffects(control, style, outerWidth, element, out _);
            runs.Add(new HtmlInlineRun(
                control,
                style,
                link,
                HtmlRenderStyleResolver.DescribeSource(element),
                inheritedPaintOffsetX + controlOffsetX,
                inheritedPaintOffsetY + controlOffsetY,
                element,
                isReplacedImage: true));
            return;
        }

        if (tag != "img" && tag != "math" && style.Display == "inline-block") {
            AddInlineBlockRun(element, width, inheritedStyle, depth, style, link, inheritedPaintOffsetX, inheritedPaintOffsetY, runs);
            return;
        }
        if (tag != "img" && tag != "math" && style.Display == "inline-flex") {
            AddInlineFlexRun(element, width, inheritedStyle, depth, style, link, inheritedPaintOffsetX, inheritedPaintOffsetY, runs);
            return;
        }
        if (tag != "img" && tag != "math" && style.Display == "inline-grid") {
            AddInlineGridRun(element, width, inheritedStyle, depth, style, link, inheritedPaintOffsetX, inheritedPaintOffsetY, runs);
            return;
        }

        if (!string.IsNullOrWhiteSpace(style.StringSet)) {
            runs.Add(new HtmlInlineRun(
                element,
                style,
                HtmlRenderStyleResolver.DescribeSource(element)));
        }

        ReportUnsupportedInlinePaintEffects(element, style);

        ResolvePositionPaintOffset(style, width, containingHeight, HtmlRenderStyleResolver.DescribeSource(element), out double elementPaintOffsetX, out double elementPaintOffsetY);
        double paintOffsetX = inheritedPaintOffsetX + elementPaintOffsetX;
        double paintOffsetY = inheritedPaintOffsetY + elementPaintOffsetY;

        List<HtmlInlineRun>? semanticRuns = HtmlRenderHeading.TryGetLevel(style.SemanticRole, out _)
            ? new List<HtmlInlineRun>()
            : null;
        ICollection<HtmlInlineRun> targetRuns = semanticRuns ?? runs;
        AddGeneratedInlineRun(element, HtmlPseudoElementKind.Before, width, containingHeight, style, link, paintOffsetX, paintOffsetY, targetRuns);

        if (tag == "img") {
            AddInlineImageRun(element, style, link, paintOffsetX, paintOffsetY, targetRuns);
            AppendSemanticInlineRuns(element, style, semanticRuns, runs);
            return;
        }
        if (tag == "math" && TryAddInlineMathRun(element, width, style, link, paintOffsetX, paintOffsetY, targetRuns)) {
            AppendSemanticInlineRuns(element, style, semanticRuns, runs);
            return;
        }

        foreach (INode child in element.ChildNodes) {
            CollectInlineRuns(child, width, containingHeight, style, link, depth + 1, paintOffsetX, paintOffsetY, targetRuns);
        }

        AddGeneratedInlineRun(element, HtmlPseudoElementKind.After, width, containingHeight, style, link, paintOffsetX, paintOffsetY, targetRuns);
        AppendSemanticInlineRuns(element, style, semanticRuns, runs);
    }

    private void AppendSemanticInlineRuns(
        IElement element,
        HtmlRenderBoxStyle style,
        IReadOnlyList<HtmlInlineRun>? semanticRuns,
        ICollection<HtmlInlineRun> destination) {
        if (semanticRuns == null) return;
        int nodeId = GetSemanticNodeId(element);
        foreach (HtmlInlineRun run in semanticRuns) {
            run.AssignSemanticNode(style.SemanticRole, nodeId);
            destination.Add(run);
        }
    }

    private void ReportUnsupportedBidi(IText textNode, HtmlRenderBoxStyle style) {
        IElement? element = textNode.ParentElement;
        if (element == null || string.IsNullOrWhiteSpace(textNode.Data) || _reportedBidiElements.Contains(element)) return;
        bool joiningScript = OfficeTextElements.ContainsJoiningScript(textNode.Data)
            && !OfficeArabicTextShaper.CanShapeAllJoiningCharacters(textNode.Data);
        if (!joiningScript) return;
        _reportedBidiElements.Add(element);
        _diagnostics.Add(
            ComponentName,
            HtmlRenderDiagnosticCodes.ComplexTextShapingUnsupported,
            "A joining script outside the bounded core-Arabic shaper used scalar glyphs.",
            HtmlDiagnosticSeverity.Warning,
            HtmlRenderStyleResolver.DescribeSource(element),
            "joining-script");
    }

    private HtmlInlineLayout LayoutInlineRuns(
        IReadOnlyList<HtmlInlineRun> runs,
        double width,
        HtmlRenderBoxStyle paragraphStyle,
        IElement? formattingContainer = null,
        int skipLogicalCharacters = 0) {
        if (runs.Count == 0 || width <= 0D) return new HtmlInlineLayout(Array.Empty<HtmlRenderVisual>(), 0D);
        if (runs.Any(run => run.FloatingBlock != null)) {
            return LayoutInlineRunsWithFloats(runs, width, paragraphStyle, formattingContainer);
        }
        bool supportsContinuationReflow = runs.All(run =>
            run.AtomicBlock == null
            && run.PositionedMarkerElement == null
            && run.RunningStringElement == null
            && run.Text.IndexOf('\u2028') < 0
            && run.Text.IndexOf('\n') < 0
            && run.Text.IndexOf('\r') < 0);
        int canonicalProgress = 0;
        bool canonicalHasContent = false;
        bool canonicalPreviousWasCollapsibleSpace = false;
        var lines = new List<InlineLine>();
        var line = new InlineLine();
        bool previousWasCollapsibleSpace = false;
        foreach (HtmlInlineRun run in runs) {
            if (run.RunningStringElement != null) {
                line.Add(new InlineSegment(string.Empty, 0D, run));
                continue;
            }
            if (run.PositionedMarkerElement != null) {
                line.Add(new InlineSegment(string.Empty, 0D, run));
                previousWasCollapsibleSpace = false;
                continue;
            }
            if (run.AtomicBlock != null) {
                previousWasCollapsibleSpace = false;
                double atomicWidth = run.AtomicBlock.Width;
                if (line.HasFlowContent && line.Width + atomicWidth > width) {
                    TrimTrailingWhitespace(line);
                    lines.Add(line);
                    line = new InlineLine();
                }

                line.Add(new InlineSegment(string.Empty, atomicWidth, run));
                continue;
            }

            int logicalOffset = 0;
            bool preserveWhitespace = run.Style.PreserveWhitespace;
            foreach (string token in Tokenize(run.Text, preserveWhitespace, run.Style.BreakSpaces)) {
                string logicalToken = SliceLogicalToken(run, token, ref logicalOffset);
                if (token == "\u2028" || preserveWhitespace && (token == "\n" || token == "\r\n")) {
                    lines.Add(line);
                    line = new InlineLine();
                    previousWasCollapsibleSpace = false;
                    continue;
                }

                bool whitespace = IsWhitespaceToken(token);
                string normalizedToken = !preserveWhitespace && whitespace ? " " : token;
                string normalizedLogicalToken = !preserveWhitespace && whitespace ? " " : logicalToken;
                bool contributesCanonicalProgress = preserveWhitespace
                    || !whitespace
                    || canonicalHasContent && !canonicalPreviousWasCollapsibleSpace;
                int tokenStart = canonicalProgress;
                if (contributesCanonicalProgress) canonicalProgress += normalizedLogicalToken.Length;
                int tokenEnd = canonicalProgress;
                if (!preserveWhitespace) {
                    if (whitespace) {
                        canonicalPreviousWasCollapsibleSpace = true;
                    } else {
                        canonicalHasContent = true;
                        canonicalPreviousWasCollapsibleSpace = false;
                    }
                }

                if (!preserveWhitespace && whitespace) {
                    if (!line.HasFlowContent || previousWasCollapsibleSpace) continue;
                    previousWasCollapsibleSpace = true;
                } else {
                    previousWasCollapsibleSpace = false;
                }

                int visibleTokenStart = tokenStart;
                if (skipLogicalCharacters > tokenStart) {
                    int skipWithinToken = skipLogicalCharacters - tokenStart;
                    if (skipWithinToken >= normalizedLogicalToken.Length) {
                        continue;
                    }
                    normalizedToken = normalizedToken.Substring(skipWithinToken);
                    normalizedLogicalToken = normalizedLogicalToken.Substring(skipWithinToken);
                    visibleTokenStart += skipWithinToken;
                    whitespace = IsWhitespaceToken(normalizedToken);
                }

                string paintToken = preserveWhitespace && normalizedToken.IndexOf('\t') >= 0
                    ? ExpandTabs(normalizedToken, run.Style, line.Width)
                    : normalizedToken;
                HyphenationToken hyphenation = PrepareHyphenationToken(paintToken, normalizedLogicalToken, run.Style);
                paintToken = hyphenation.PaintText;
                string logicalPaintToken = hyphenation.LogicalText;
                double measured = MeasureInlineText(paintToken, run.Style);
                if (!paragraphStyle.PreventTextWrapping
                    && !whitespace
                    && measured > Math.Max(0D, width - line.Width)
                    && TryAddHyphenatedToken(
                        lines,
                        ref line,
                        run,
                        hyphenation,
                        width,
                        visibleTokenStart,
                        tokenEnd)) {
                    continue;
                }
                if (!paragraphStyle.PreventTextWrapping
                    && !whitespace
                    && measured > Math.Max(0D, width - line.Width)
                    && TryAddPreferredBreakToken(
                        lines,
                        ref line,
                        run,
                        paintToken,
                        logicalPaintToken,
                        width,
                        visibleTokenStart)) {
                    continue;
                }
                bool breakAllIntoRemainingSpace = run.Style.WordBreak == "break-all"
                    && line.HasFlowContent
                    && measured > Math.Max(0D, width - line.Width);
                if (!paragraphStyle.PreventTextWrapping
                    && !whitespace
                    && AllowsEmergencyTokenBreak(run.Style)
                    && (measured > width || breakAllIntoRemainingSpace)) {
                    AddBrokenToken(lines, ref line, run, paintToken, logicalPaintToken, width, visibleTokenStart);
                    continue;
                }

                if (!paragraphStyle.PreventTextWrapping && line.HasFlowContent && line.Width + measured > width) {
                    TrimTrailingWhitespace(line);
                    lines.Add(line);
                    line = new InlineLine();
                    if (whitespace && !preserveWhitespace) continue;
                }

                line.Add(new InlineSegment(paintToken, measured, run, logicalPaintToken, logicalEndProgress: tokenEnd));
            }
        }

        TrimTrailingWhitespace(line);
        if (line.Segments.Count > 0 || lines.Count == 0) lines.Add(line);
        int completeLogicalProgress = lines
            .SelectMany(candidate => candidate.Segments)
            .Select(segment => segment.LogicalEndProgress)
            .DefaultIfEmpty(canonicalProgress)
            .Max();
        if (paragraphStyle.LineClamp.HasValue && lines.Count > paragraphStyle.LineClamp.Value) {
            lines.RemoveRange(paragraphStyle.LineClamp.Value, lines.Count - paragraphStyle.LineClamp.Value);
            ApplyEndEllipsis(lines[lines.Count - 1], width, completeLogicalProgress);
        } else if (paragraphStyle.TextOverflow == "ellipsis"
            && paragraphStyle.PreventTextWrapping
            && paragraphStyle.OverflowX != "visible"
            && lines.Count > 0
            && lines[0].Width > width + 0.0001D) {
            ApplyEndEllipsis(lines[0], width, completeLogicalProgress);
        }
        return RenderInlineLines(lines, width, paragraphStyle, formattingContainer, supportsContinuationReflow: supportsContinuationReflow);
    }

    private HyphenationToken PrepareHyphenationToken(string paintToken, string logicalToken, HtmlRenderBoxStyle style) {
        if (paintToken.IndexOf('\u00AD') < 0
            && (style.Hyphens != "auto" || _options.TextHyphenationCallback == null)) {
            return new HyphenationToken(
                paintToken,
                logicalToken,
                Array.Empty<int>(),
                Array.Empty<int>(),
                Array.Empty<int>());
        }
        var paint = new StringBuilder(paintToken.Length);
        var logical = new StringBuilder(logicalToken.Length);
        var manualBreaks = new List<int>();
        var sourceBoundaries = new List<int> { 0 };
        for (int sourceIndex = 0; sourceIndex < logicalToken.Length; sourceIndex++) {
            if (logicalToken[sourceIndex] == '\u00AD') {
                if (logical.Length > 0) manualBreaks.Add(logical.Length);
                sourceBoundaries[logical.Length] = sourceIndex + 1;
                continue;
            }
            logical.Append(logicalToken[sourceIndex]);
            sourceBoundaries.Add(sourceIndex + 1);
        }
        foreach (char character in paintToken) {
            if (character != '\u00AD') paint.Append(character);
        }

        var automaticBreaks = new SortedSet<int>();
        if (style.Hyphens == "auto" && _options.TextHyphenationCallback != null) {
            IReadOnlyList<int>? automatic = _options.TextHyphenationCallback(logical.ToString());
            if (automatic != null) {
                foreach (int point in automatic) automaticBreaks.Add(point);
            }
        }

        string logicalText = logical.ToString();
        int minimumWordLength = Math.Max(1, style.HyphenateMinimumWordLength);
        int minimumPrefix = Math.Max(1, style.HyphenateMinimumPrefixLength);
        int minimumSuffix = Math.Max(1, style.HyphenateMinimumSuffixLength);
        int[] FilterBreaks(IEnumerable<int> candidates) => CountCssHyphenationCharacters(logicalText, 0, logicalText.Length) < minimumWordLength
            ? Array.Empty<int>()
            : candidates
                .Where(point => OfficeTextLineBreaks.IsValidBreakPosition(logicalText, point))
                .Where(point => CountCssHyphenationCharacters(logicalText, 0, point) >= minimumPrefix
                    && CountCssHyphenationCharacters(logicalText, point, logicalText.Length - point) >= minimumSuffix)
                .Distinct()
                .OrderBy(point => point)
                .ToArray();
        int[] primaryBreaks = style.Hyphens == "none"
            ? Array.Empty<int>()
            : FilterBreaks(manualBreaks.Count > 0 ? manualBreaks : automaticBreaks);
        int[] secondaryBreaks = style.Hyphens == "auto" && manualBreaks.Count > 0
            ? FilterBreaks(automaticBreaks.Where(point => !manualBreaks.Contains(point)))
            : Array.Empty<int>();
        return new HyphenationToken(paint.ToString(), logicalText, primaryBreaks, secondaryBreaks, sourceBoundaries.ToArray());
    }

    private static int CountCssHyphenationCharacters(string value, int start, int length) {
        int count = 0;
        foreach (string element in OfficeTextElements.Enumerate(value.Substring(start, length))) {
            UnicodeCategory category = CharUnicodeInfo.GetUnicodeCategory(element, 0);
            if (category == UnicodeCategory.NonSpacingMark || IsPunctuationCategory(category)) continue;
            count++;
        }
        return count;
    }

    private static bool IsPunctuationCategory(UnicodeCategory category) => category == UnicodeCategory.ConnectorPunctuation
        || category == UnicodeCategory.DashPunctuation
        || category == UnicodeCategory.OpenPunctuation
        || category == UnicodeCategory.ClosePunctuation
        || category == UnicodeCategory.InitialQuotePunctuation
        || category == UnicodeCategory.FinalQuotePunctuation
        || category == UnicodeCategory.OtherPunctuation;

    private bool TryAddHyphenatedToken(
        ICollection<InlineLine> lines,
        ref InlineLine line,
        HtmlInlineRun run,
        HyphenationToken token,
        double width,
        int logicalStartProgress,
        int logicalEndProgress) {
        if (!token.HasBreaks || token.PaintText.Length != token.LogicalText.Length) return false;
        if (run.Style.HyphenateLimitLast == "always"
            && line.HasFlowContent
            && MeasureInlineText(token.PaintText, run.Style) <= width + 0.0001D) {
            TrimTrailingWhitespace(line);
            lines.Add(line);
            line = new InlineLine();
            line.Add(new InlineSegment(
                token.PaintText,
                MeasureInlineText(token.PaintText, run.Style),
                run,
                token.LogicalText,
                logicalEndProgress: logicalEndProgress));
            return true;
        }
        if (line.HasFlowContent && run.Style.HyphenateLimitZone > 0D
            && width - line.Width <= run.Style.HyphenateLimitZone + 0.0001D) {
            TrimTrailingWhitespace(line);
            lines.Add(line);
            line = new InlineLine();
        }

        int start = 0;
        while (start < token.PaintText.Length) {
            double available = Math.Max(0D, width - line.Width);
            bool hyphenationAllowed = !run.Style.HyphenateLimitLines.HasValue
                || CountConsecutiveHyphenatedLines(lines) < run.Style.HyphenateLimitLines.Value;
            int selectedEnd = -1;
            bool selectedIsBreak = false;
            if (MeasureInlineText(token.PaintText.Substring(start), run.Style) <= available + 0.0001D) {
                selectedEnd = token.PaintText.Length;
            } else if (hyphenationAllowed) {
                selectedEnd = SelectHyphenationBreak(token.PrimaryBreaks, token.PaintText, start, available, run.Style);
                if (selectedEnd < 0) {
                    selectedEnd = SelectHyphenationBreak(token.SecondaryBreaks, token.PaintText, start, available, run.Style);
                }
                selectedIsBreak = selectedEnd >= 0;
            }

            if (selectedEnd < 0) {
                if (!line.HasFlowContent) {
                    if (AllowsEmergencyTokenBreak(run.Style)) return false;
                    string paintRemainder = token.PaintText.Substring(start);
                    string logicalRemainder = token.LogicalText.Substring(start);
                    line.Add(new InlineSegment(
                        paintRemainder,
                        MeasureInlineText(paintRemainder, run.Style),
                        run,
                        logicalRemainder,
                        logicalEndProgress: logicalEndProgress));
                    return true;
                }
                TrimTrailingWhitespace(line);
                lines.Add(line);
                line = new InlineLine();
                continue;
            }

            string paintChunk = token.PaintText.Substring(start, selectedEnd - start)
                + (selectedIsBreak ? run.Style.HyphenateCharacter : string.Empty);
            string logicalChunk = token.LogicalText.Substring(start, selectedEnd - start);
            int sourceBoundary = selectedEnd < token.SourceBoundaries.Count
                ? token.SourceBoundaries[selectedEnd]
                : logicalEndProgress - logicalStartProgress;
            line.Add(new InlineSegment(
                paintChunk,
                MeasureInlineText(paintChunk, run.Style),
                run,
                logicalChunk,
                logicalEndProgress: selectedEnd == token.PaintText.Length
                    ? logicalEndProgress
                    : logicalStartProgress + sourceBoundary));
            start = selectedEnd;
            if (selectedIsBreak) {
                line.EndsWithHyphenation = true;
                lines.Add(line);
                line = new InlineLine();
            }
        }
        return true;
    }

    private int SelectHyphenationBreak(
        IReadOnlyList<int> candidates,
        string paintText,
        int start,
        double available,
        HtmlRenderBoxStyle style) {
        int selected = -1;
        foreach (int point in candidates) {
            if (point <= start) continue;
            string candidate = paintText.Substring(start, point - start) + style.HyphenateCharacter;
            if (MeasureInlineText(candidate, style) <= available + 0.0001D) selected = point;
        }
        return selected;
    }

    private static int CountConsecutiveHyphenatedLines(ICollection<InlineLine> lines) {
        int count = 0;
        foreach (InlineLine candidate in lines.Reverse()) {
            if (!candidate.EndsWithHyphenation) break;
            count++;
        }
        return count;
    }

    private bool TryAddPreferredBreakToken(
        ICollection<InlineLine> lines,
        ref InlineLine line,
        HtmlInlineRun run,
        string paintToken,
        string logicalToken,
        double width,
        int logicalStartProgress) {
        if (paintToken.Length != logicalToken.Length) return false;
        IReadOnlyList<int> breaks = OfficeTextLineBreaks.GetBreakPositions(
            paintToken,
            allowCjkBreaks: run.Style.WordBreak != "keep-all");
        if (breaks.Count == 0) return false;
        int start = 0;
        foreach (int end in breaks.Concat(new[] { paintToken.Length })) {
            if (end <= start || end > paintToken.Length) continue;
            string paintChunk = paintToken.Substring(start, end - start);
            string logicalChunk = logicalToken.Substring(start, end - start);
            double chunkWidth = MeasureInlineText(paintChunk, run.Style);
            if (chunkWidth > width && AllowsEmergencyTokenBreak(run.Style)) {
                AddBrokenToken(lines, ref line, run, paintChunk, logicalChunk, width, logicalStartProgress + start);
                start = end;
                continue;
            }
            if (line.HasFlowContent && line.Width + chunkWidth > width) {
                TrimTrailingWhitespace(line);
                lines.Add(line);
                line = new InlineLine();
            }
            line.Add(new InlineSegment(
                paintChunk,
                chunkWidth,
                run,
                logicalChunk,
                logicalEndProgress: logicalStartProgress + end));
            start = end;
        }
        return start == paintToken.Length;
    }

    private string ExpandTabs(string value, HtmlRenderBoxStyle style, double currentWidth) {
        if (value.IndexOf('\t') < 0) return value;
        double spaceWidth = Math.Max(0.01D, MeasureInlineText(" ", style));
        double stopWidth = Math.Max(spaceWidth, style.TabSize * spaceWidth);
        double cursor = Math.Max(0D, currentWidth);
        var expanded = new StringBuilder();
        foreach (char character in value) {
            if (character != '\t') {
                expanded.Append(character);
                cursor += MeasureInlineText(character.ToString(), style);
                continue;
            }
            double nextStop = (Math.Floor(cursor / stopWidth) + 1D) * stopWidth;
            int spaces = Math.Max(1, (int)Math.Round((nextStop - cursor) / spaceWidth));
            expanded.Append(' ', spaces);
            cursor += spaces * spaceWidth;
        }
        return expanded.ToString();
    }

    private void ApplyEndEllipsis(InlineLine line, double width, int completeLogicalProgress) {
        double availableWidth = line.HasExplicitPlacement ? line.AvailableWidth : width;
        TrimTrailingWhitespace(line);
        HtmlInlineRun? ellipsisRun = line.Segments
            .LastOrDefault(segment => segment.Run.AtomicBlock == null && segment.Text.Length > 0)?.Run;
        if (ellipsisRun == null && line.Segments.Count > 0) ellipsisRun = line.Segments[line.Segments.Count - 1].Run;
        while (line.Segments.Count > 0) {
            InlineSegment segment = line.Segments[line.Segments.Count - 1];
            if (segment.Run.AtomicBlock != null || segment.Text.Length == 0) {
                line.RemoveAt(line.Segments.Count - 1);
                continue;
            }

            ellipsisRun = segment.Run;
            line.RemoveAt(line.Segments.Count - 1);
            double remainingWidth = Math.Max(0D, availableWidth - line.Width);
            double ellipsisWidth = MeasureInlineText("\u2026", segment.Run.Style);
            if (ellipsisWidth > remainingWidth + 0.0001D) continue;

            var paint = new StringBuilder();
            var logical = new StringBuilder();
            IReadOnlyList<string> paintElements = OfficeTextElements.Split(segment.Text);
            IReadOnlyList<string> logicalElements = OfficeTextElements.Split(segment.LogicalText);
            for (int index = 0; index < paintElements.Count; index++) {
                string candidate = paint.ToString() + paintElements[index];
                if (MeasureInlineText(candidate, segment.Run.Style) + ellipsisWidth > remainingWidth + 0.0001D) break;
                paint.Append(paintElements[index]);
                if (index < logicalElements.Count) logical.Append(logicalElements[index]);
            }

            string text = paint.ToString() + "\u2026";
            string logicalText = logical.ToString() + "\u2026";
            line.Add(new InlineSegment(
                text,
                MeasureInlineText(text, segment.Run.Style),
                segment.Run,
                logicalText,
                logicalEndProgress: completeLogicalProgress));
            return;
        }

        if (ellipsisRun != null) {
            ellipsisRun = CreateEllipsisRun(ellipsisRun);
            double ellipsisWidth = MeasureInlineText("\u2026", ellipsisRun.Style);
            if (ellipsisWidth <= availableWidth + 0.0001D) {
                line.Add(new InlineSegment("\u2026", ellipsisWidth, ellipsisRun, "\u2026", logicalEndProgress: completeLogicalProgress));
            }
        }
    }

    private static HtmlInlineRun CreateEllipsisRun(HtmlInlineRun source) {
        if (source.AtomicBlock == null) return source;
        var run = new HtmlInlineRun(
            "\u2026",
            source.Style,
            source.LinkUri,
            source.Source,
            source.PaintOffsetX,
            source.PaintOffsetY,
            source.OwnerElement,
            logicalText: "\u2026");
        if (source.SemanticNodeId.HasValue) run.AssignSemanticNode(source.SemanticRole, source.SemanticNodeId.Value);
        return run;
    }

    private static bool AllowsEmergencyTokenBreak(HtmlRenderBoxStyle style) =>
        style.OverflowWrap == "anywhere"
        || style.OverflowWrap == "break-word"
        || style.WordBreak == "break-all"
        || style.WordBreak == "break-word";

    private static IReadOnlyList<InlineSegment> MergeAdjacentInlineSegments(IReadOnlyList<InlineSegment> segments) {
        var merged = new List<InlineSegment>(segments.Count);
        foreach (InlineSegment segment in segments) {
            if (segment.Run.AtomicBlock == null && merged.Count > 0 && ReferenceEquals(merged[merged.Count - 1].Run, segment.Run)) {
                InlineSegment previous = merged[merged.Count - 1];
                merged[merged.Count - 1] = new InlineSegment(
                    previous.Text + segment.Text,
                    previous.Width + segment.Width,
                    previous.Run,
                    previous.LogicalText + segment.LogicalText,
                    logicalEndProgress: segment.LogicalEndProgress);
            } else {
                merged.Add(segment);
            }
        }

        return merged;
    }

    private void AddBrokenToken(
        ICollection<InlineLine> lines,
        ref InlineLine line,
        HtmlInlineRun run,
        string token,
        string logicalToken,
        double width,
        int logicalStartProgress) {
        var part = new StringBuilder();
        var logicalPart = new StringBuilder();
        double partWidth = 0D;
        int partLogicalLength = 0;
        IReadOnlyList<string> paintElements = OfficeTextElements.Split(token);
        IReadOnlyList<string> logicalElements = OfficeTextElements.Split(logicalToken);
        double partLimit = line.HasFlowContent ? Math.Max(0D, width - line.Width) : width;
        for (int index = 0; index < paintElements.Count; index++) {
            string value = paintElements[index];
            string logicalValue = index < logicalElements.Count ? logicalElements[index] : OfficeArabicTextShaper.ToLogicalText(value);
            double charWidth = MeasureInlineText(value, run.Style);
            if (part.Length > 0 && partWidth + charWidth > partLimit) {
                line.Add(new InlineSegment(
                    part.ToString(),
                    partWidth,
                    run,
                    logicalPart.ToString(),
                    logicalEndProgress: logicalStartProgress + partLogicalLength));
                lines.Add(line);
                line = new InlineLine();
                part.Clear();
                logicalPart.Clear();
                partWidth = 0D;
                partLimit = width;
            } else if (part.Length == 0 && line.HasFlowContent && charWidth > partLimit) {
                TrimTrailingWhitespace(line);
                lines.Add(line);
                line = new InlineLine();
                partLimit = width;
            }

            part.Append(value);
            logicalPart.Append(logicalValue);
            partLogicalLength += logicalValue.Length;
            partWidth += charWidth;
        }

        if (part.Length > 0) {
            if (line.HasFlowContent && line.Width + partWidth > width) {
                TrimTrailingWhitespace(line);
                lines.Add(line);
                line = new InlineLine();
            }

            line.Add(new InlineSegment(
                part.ToString(),
                partWidth,
                run,
                logicalPart.ToString(),
                logicalEndProgress: logicalStartProgress + partLogicalLength));
        }
    }

    private static string SliceLogicalToken(HtmlInlineRun run, string token, ref int offset) {
        if (offset >= 0 && token.Length <= run.LogicalText.Length - offset) {
            string value = run.LogicalText.Substring(offset, token.Length);
            offset += token.Length;
            return value;
        }

        offset += token.Length;
        return OfficeArabicTextShaper.ToLogicalText(token);
    }

    private double MeasureText(string value, OfficeFontInfo font) {
        if (_fonts.TryMeasureText(value, font.Size, font.FamilyName, font.Style, out double scopedWidth)) {
            return scopedWidth;
        }

        OfficeTextMeasurer measurer = OfficeTextMeasurer.Create(font);
        OfficeTextMeasurementStyle style = measurer.CreateStyle(font, 72D);
        return measurer.MeasureWidth(value, style);
    }

    private double MeasureInlineText(string value, HtmlRenderBoxStyle style) {
        double measured = MeasureText(value, style.Font);
        if (Math.Abs(style.LetterSpacing) <= 0.000001D && Math.Abs(style.WordSpacing) <= 0.000001D) {
            return Math.Max(0.01D, measured);
        }
        IReadOnlyList<string> elements = OfficeTextElements.Split(value);
        if (elements.Count == 0) return measured;
        measured += style.LetterSpacing * elements.Count;
        measured += style.WordSpacing * elements.Count(IsWhitespaceToken);
        return Math.Max(0.01D, measured);
    }

    private string? ResolveSafeLink(string? rawHref, IElement element) {
        if (string.IsNullOrWhiteSpace(rawHref)) return null;
        string resolved = HtmlUrlPolicyEvaluator.ResolveUrl(rawHref, _baseUri, _options.UrlPolicy);
        if (resolved.Length > 0) return resolved;
        _diagnostics.Add(ComponentName, "HyperlinkRejectedByPolicy", "A hyperlink target was rejected before entering the rendered document.", HtmlDiagnosticSeverity.Warning, HtmlRenderStyleResolver.DescribeSource(element), rawHref);
        return null;
    }

    private static IEnumerable<string> Tokenize(string text, bool preserveWhitespace, bool breakSpaces) {
        if (text.Length == 0) yield break;
        var token = new StringBuilder();
        bool? whitespace = null;
        for (int i = 0; i < text.Length; i++) {
            char current = text[i];
            if (current == '\u2028') {
                if (token.Length > 0) {
                    yield return token.ToString();
                    token.Clear();
                }

                whitespace = null;
                yield return "\u2028";
                continue;
            }

            if (preserveWhitespace && (current == '\r' || current == '\n')) {
                if (token.Length > 0) {
                    yield return token.ToString();
                    token.Clear();
                }

                if (current == '\r' && i + 1 < text.Length && text[i + 1] == '\n') i++;
                whitespace = null;
                yield return "\n";
                continue;
            }

            bool currentWhitespace = char.IsWhiteSpace(current);
            if (breakSpaces && currentWhitespace) {
                if (token.Length > 0) {
                    yield return token.ToString();
                    token.Clear();
                }
                whitespace = null;
                yield return current.ToString();
                continue;
            }
            if (whitespace.HasValue && whitespace.Value != currentWhitespace) {
                yield return token.ToString();
                token.Clear();
            }

            whitespace = currentWhitespace;
            token.Append(current);
        }

        if (token.Length > 0) yield return token.ToString();
    }

    private static string ApplyTextTransform(string text, string transform) {
        if (transform == "uppercase") return text.ToUpperInvariant();
        if (transform == "lowercase") return text.ToLowerInvariant();
        if (transform == "capitalize") {
            var builder = new StringBuilder(text.Length);
            bool capitalize = true;
            foreach (char character in text) {
                builder.Append(capitalize ? char.ToUpperInvariant(character) : character);
                capitalize = char.IsWhiteSpace(character);
            }

            return builder.ToString();
        }

        return text;
    }

    private static bool IsWhitespaceToken(string token) => token.Length > 0 && token.All(char.IsWhiteSpace);

    private static void TrimTrailingWhitespace(InlineLine line) {
        for (int index = line.Segments.Count - 1; index >= 0; index--) {
            InlineSegment segment = line.Segments[index];
            if (segment.Run.RunningStringElement != null) continue;
            if (!IsWhitespaceToken(segment.Text)) break;
            if (segment.Run.Style.BreakSpaces) break;
            line.RemoveAt(index);
        }
    }

    private static double ResolveLineOffset(OfficeTextAlignment alignment, double width, double lineWidth) {
        if (alignment == OfficeTextAlignment.Center) return Math.Max(0D, (width - lineWidth) / 2D);
        if (alignment == OfficeTextAlignment.Right) return Math.Max(0D, width - lineWidth);
        return 0D;
    }

    private sealed class InlineLine {
        private int _flowContentCount;

        internal List<InlineSegment> Segments { get; } = new List<InlineSegment>();
        internal double Width { get; private set; }
        internal bool HasFlowContent => _flowContentCount > 0;
        internal bool HasExplicitPlacement { get; private set; }
        internal double X { get; private set; }
        internal double Y { get; private set; }
        internal double AvailableWidth { get; private set; }
        internal bool EndsWithHyphenation { get; set; }

        internal void Place(double x, double y, double availableWidth) {
            HasExplicitPlacement = true;
            X = Math.Max(0D, x);
            Y = Math.Max(0D, y);
            AvailableWidth = Math.Max(0.01D, availableWidth);
        }

        internal void Add(InlineSegment segment) {
            Segments.Add(segment);
            Width += segment.Width;
            if (segment.Run.RunningStringElement == null) _flowContentCount++;
        }

        internal void RemoveAt(int index) {
            if (Segments[index].Run.RunningStringElement == null) _flowContentCount--;
            Width -= Segments[index].Width;
            Segments.RemoveAt(index);
        }

        internal double ResolveLineHeight(double fallback) {
            if (!HasFlowContent) return 0D;
            double height = fallback;
            for (int i = 0; i < Segments.Count; i++) {
                height = Math.Max(height, Segments[i].Run.AtomicBlock?.Height ?? Segments[i].Run.Style.LineHeight);
            }
            if (!HasReplacedImage) return Math.Max(0.01D, height);

            double ascent = 0D;
            double descent = 0D;
            for (int i = 0; i < Segments.Count; i++) {
                HtmlInlineRun run = Segments[i].Run;
                if (run.AtomicBlock != null) {
                    double atomicBaseline = Math.Min(run.AtomicBlock.Height, Math.Max(0D, run.AtomicBaseline ?? run.AtomicBlock.Height));
                    ascent = Math.Max(ascent, atomicBaseline);
                    descent = Math.Max(descent, run.AtomicBlock.Height - atomicBaseline);
                } else {
                    ascent = Math.Max(ascent, ResolveTextAscent(run.Style));
                    descent = Math.Max(descent, Math.Max(0D, run.Style.LineHeight - ResolveTextAscent(run.Style)));
                }
            }
            return Math.Max(0.01D, ascent + descent);
        }

        internal bool HasReplacedImage => Segments.Any(segment => segment.Run.IsReplacedImage);

        internal double ResolveBaseline(double fallback) {
            if (!HasReplacedImage) return ResolveLineHeight(fallback);
            double ascent = 0D;
            for (int i = 0; i < Segments.Count; i++) {
                HtmlInlineRun run = Segments[i].Run;
                ascent = Math.Max(ascent, run.AtomicBlock == null
                    ? ResolveTextAscent(run.Style)
                    : Math.Min(run.AtomicBlock.Height, Math.Max(0D, run.AtomicBaseline ?? run.AtomicBlock.Height)));
            }
            return ascent;
        }
    }

    private readonly struct HyphenationToken {
        internal HyphenationToken(
            string paintText,
            string logicalText,
            IReadOnlyList<int> primaryBreaks,
            IReadOnlyList<int> secondaryBreaks,
            IReadOnlyList<int> sourceBoundaries) {
            PaintText = paintText;
            LogicalText = logicalText;
            PrimaryBreaks = primaryBreaks;
            SecondaryBreaks = secondaryBreaks;
            SourceBoundaries = sourceBoundaries;
        }

        internal string PaintText { get; }
        internal string LogicalText { get; }
        internal IReadOnlyList<int> PrimaryBreaks { get; }
        internal IReadOnlyList<int> SecondaryBreaks { get; }
        internal bool HasBreaks => PrimaryBreaks.Count > 0 || SecondaryBreaks.Count > 0;
        internal IReadOnlyList<int> SourceBoundaries { get; }
    }

    private sealed class InlineSegment {
        internal InlineSegment(
            string text,
            double width,
            HtmlInlineRun run,
            string? logicalText = null,
            bool bidiResolved = false,
            int logicalEndProgress = 0) {
            Text = text;
            LogicalText = logicalText ?? text;
            Width = width;
            Run = run;
            BidiResolved = bidiResolved;
            LogicalEndProgress = logicalEndProgress;
        }

        internal string Text { get; }
        internal string LogicalText { get; }
        internal double Width { get; }
        internal HtmlInlineRun Run { get; }
        internal bool BidiResolved { get; }
        internal int LogicalEndProgress { get; }
    }

    private static double ResolveTextAscent(HtmlRenderBoxStyle style) {
        double leading = Math.Max(0D, style.LineHeight - style.Font.Size);
        return Math.Min(style.LineHeight, leading / 2D + style.Font.Size * 0.8D);
    }
}
