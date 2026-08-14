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

        if (formattingContainer != null && ShouldAssignNavigationNode(parentStyle)) {
            int semanticNodeId = GetSemanticNodeId(formattingContainer);
            foreach (HtmlInlineRun run in runs) run.AssignSemanticNode(parentStyle.SemanticRole, semanticNodeId);
        }
        AssignSemanticFragmentOrders(runs);

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
                var resolvedRun = new HtmlInlineRun(
                    OfficeArabicTextShaper.Shape(fallback.Text),
                    style,
                    run.LinkUri,
                    run.Source,
                    run.PaintOffsetX,
                    run.PaintOffsetY,
                    run.OwnerElement,
                    run.PositionedMarkerElement,
                    fallback.Text);
                if (run.SemanticNodeId.HasValue) {
                    resolvedRun.AssignSemanticNode(run.SemanticRole, run.SemanticNodeId.Value, run.BookmarkAnchorText, run.SemanticFragmentOrder);
                }
                if (run.InlineSemanticGroupRole.HasValue && run.InlineSemanticGroupKey != null) {
                    resolvedRun.AssignInlineSemanticGroup(run.InlineSemanticGroupRole.Value, run.InlineSemanticGroupKey);
                }
                resolvedRuns.Add(resolvedRun);
            }
        }

        return resolvedRuns;
    }

    private static void AssignSemanticFragmentOrders(IEnumerable<HtmlInlineRun> runs) {
        var nextOrders = new Dictionary<int, int>();
        foreach (HtmlInlineRun run in runs) {
            if (!run.SemanticNodeId.HasValue) continue;
            int nodeId = run.SemanticNodeId.Value;
            nextOrders.TryGetValue(nodeId, out int order);
            run.AssignSemanticNode(run.SemanticRole, nodeId, run.BookmarkAnchorText, order);
            nextOrders[nodeId] = order + 1;
        }
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
        RegisterInlineSemanticControls(element, style);
        string? link = inheritedLink;
        if (tag == "a") {
            link = ResolveSafeLink(element.GetAttribute("href"), element);
        }
        if (HtmlCssRunningElementParser.TryParsePosition(style.Position, out string runningElementName)) {
            runs.Add(new HtmlInlineRun(
                CaptureRunningElement(element, runningElementName, width, style, inheritedStyle, depth + 1),
                style,
                HtmlRenderStyleResolver.DescribeSource(element)));
            return;
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
            var controlRun = new HtmlInlineRun(
                control,
                style,
                link,
                HtmlRenderStyleResolver.DescribeSource(element),
                inheritedPaintOffsetX + controlOffsetX,
                inheritedPaintOffsetY + controlOffsetY,
                element,
                isReplacedImage: true);
            int controlNodeId = GetSemanticNodeId(element);
            if (ShouldAssignNavigationNode(style)) {
                controlRun.AssignSemanticNode(style.SemanticRole, controlNodeId, ResolveBookmarkAnchorText(element, style));
            }
            AssignInlineSemanticGroup(controlRun, style, controlNodeId);
            runs.Add(controlRun);
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

        if (style.Transform != "none"
            || style.OpacityWasSpecified && (style.Opacity < 1D || style.UnsupportedOpacity.Length > 0)
            || style.ClipPath != "none") {
            _inlineStackingElements.Add(element);
        }

        if (!string.IsNullOrWhiteSpace(style.StringSet)) {
            runs.Add(new HtmlInlineRun(
                element,
                style,
                HtmlRenderStyleResolver.DescribeSource(element)));
        }

        ResolvePositionPaintOffset(style, width, containingHeight, HtmlRenderStyleResolver.DescribeSource(element), out double elementPaintOffsetX, out double elementPaintOffsetY);
        double paintOffsetX = inheritedPaintOffsetX + elementPaintOffsetX;
        double paintOffsetY = inheritedPaintOffsetY + elementPaintOffsetY;

        List<HtmlInlineRun>? semanticRuns = ShouldCollectSemanticInlineRuns(style)
            ? new List<HtmlInlineRun>()
            : null;
        ICollection<HtmlInlineRun> targetRuns = semanticRuns ?? runs;
        AddGeneratedInlineRun(element, HtmlPseudoElementKind.Before, width, containingHeight, style, link, paintOffsetX, paintOffsetY, targetRuns);

        if (tag == "img") {
            AddInlineImageRun(element, style, link, paintOffsetX, paintOffsetY, targetRuns);
            AppendSemanticInlineRuns(element, style, semanticRuns, runs, link, paintOffsetX, paintOffsetY);
            return;
        }
        if (tag == "math" && TryAddInlineMathRun(element, width, style, link, paintOffsetX, paintOffsetY, targetRuns)) {
            AppendSemanticInlineRuns(element, style, semanticRuns, runs, link, paintOffsetX, paintOffsetY);
            return;
        }

        foreach (INode child in element.ChildNodes) {
            CollectInlineRuns(child, width, containingHeight, style, link, depth + 1, paintOffsetX, paintOffsetY, targetRuns);
        }

        AddGeneratedInlineRun(element, HtmlPseudoElementKind.After, width, containingHeight, style, link, paintOffsetX, paintOffsetY, targetRuns);
        AppendSemanticInlineRuns(element, style, semanticRuns, runs, link, paintOffsetX, paintOffsetY);
    }

    private static bool ShouldAssignNavigationNode(HtmlRenderBoxStyle style) =>
        HtmlRenderHeading.TryGetLevel(style.SemanticRole, out _)
        || style.BookmarkLevelSpecified
        || style.BookmarkLabel != null
        || style.BookmarkState != HtmlRenderBookmarkState.Default;

    private static bool ShouldCollectSemanticInlineRuns(HtmlRenderBoxStyle style) =>
        ShouldAssignNavigationNode(style)
        || style.SemanticArtifact
        || style.SemanticGroupRoleOverride.HasValue;

    private void AppendSemanticInlineRuns(
        IElement element,
        HtmlRenderBoxStyle style,
        IReadOnlyList<HtmlInlineRun>? semanticRuns,
        ICollection<HtmlInlineRun> destination,
        string? link,
        double paintOffsetX,
        double paintOffsetY) {
        if (semanticRuns == null) return;
        int nodeId = GetSemanticNodeId(element);
        string bookmarkAnchorText = ResolveBookmarkAnchorText(element, style);
        if (semanticRuns.Count == 0
            && ShouldAssignNavigationNode(style)
            && !style.BookmarkSuppressed
            && bookmarkAnchorText.Length > 0) {
            string source = HtmlRenderStyleResolver.DescribeSource(element);
            var anchor = new HtmlRenderBookmarkAnchor(nodeId, bookmarkAnchorText, 0D, 0D, 0.01D, 0.01D, 0, source);
            var markerBlock = new HtmlRenderFlowBlock(
                0.01D,
                0.01D,
                new[] { anchor },
                HtmlPageBreakTarget.None,
                HtmlPageBreakTarget.None,
                false,
                source);
            var markerRun = new HtmlInlineRun(markerBlock, style, link, source, paintOffsetX, paintOffsetY, element);
            markerRun.AssignSemanticNode(style.SemanticRole, nodeId, bookmarkAnchorText);
            AssignInlineSemanticGroup(markerRun, style, nodeId);
            destination.Add(markerRun);
            return;
        }
        foreach (HtmlInlineRun run in semanticRuns) {
            if (ShouldAssignNavigationNode(style)) run.AssignSemanticNode(style.SemanticRole, nodeId, bookmarkAnchorText);
            AssignInlineSemanticGroup(run, style, nodeId);
            destination.Add(run);
        }
    }

    private static void AssignInlineSemanticGroup(HtmlInlineRun run, HtmlRenderBoxStyle style, int semanticNodeId) {
        string structureElementKey = "html-inline:" + semanticNodeId.ToString(System.Globalization.CultureInfo.InvariantCulture);
        if (style.SemanticArtifact) run.AssignInlineSemanticGroup(HtmlRenderSemanticGroupRole.Artifact, structureElementKey);
        else if (style.SemanticGroupRoleOverride.HasValue) run.AssignInlineSemanticGroup(style.SemanticGroupRoleOverride.Value, structureElementKey);
    }

    private void ReportUnsupportedBidi(IText textNode, HtmlRenderBoxStyle style) {
        IElement? element = textNode.ParentElement;
        if (element == null || string.IsNullOrWhiteSpace(textNode.Data) || _reportedBidiElements.Contains(element)) return;
        bool joiningScript = OfficeTextElements.ContainsJoiningScript(textNode.Data)
            && !OfficeArabicTextShaper.CanShapeAllJoiningCharacters(textNode.Data);
        if (!joiningScript) return;
        IReadOnlyList<OfficeFontFallbackRun> fallbackRuns = _fonts.PlanFallbackRuns(
            textNode.Data,
            style.Font.FamilyName,
            style.Font.Style);
        bool allUnsupportedRunsShaped = true;
        foreach (OfficeFontFallbackRun fallback in fallbackRuns) {
            if (!OfficeTextElements.ContainsJoiningScript(fallback.Text)
                || OfficeArabicTextShaper.CanShapeAllJoiningCharacters(fallback.Text)) continue;
            HtmlRenderBoxStyle fallbackStyle = style.Clone();
            fallbackStyle.Font = fallbackStyle.Font.WithFamilyName(fallback.FamilyName);
            if (!TryShapeWithConfiguredProvider(fallback.Text, fallbackStyle)) {
                allUnsupportedRunsShaped = false;
                break;
            }
        }
        if (allUnsupportedRunsShaped) return;
        _reportedBidiElements.Add(element);
        _diagnostics.Add(
            ComponentName,
            HtmlRenderDiagnosticCodes.ComplexTextShapingUnsupported,
            "A joining script outside the bounded core-Arabic shaper used scalar glyphs.",
            HtmlDiagnosticSeverity.Warning,
            HtmlRenderStyleResolver.DescribeSource(element),
            "joining-script");
    }

    private bool TryShapeWithConfiguredProvider(string text, HtmlRenderBoxStyle style) {
        IOfficeTextShapingProvider? provider = _options.TextShapingProvider;
        if (provider == null) return false;

        OfficeTrueTypeFont? font = _fonts.ResolveForText(
            text,
            style.Font.FamilyName,
            style.Font.Style,
            out OfficeFontStyle _);
        if (font == null) return false;

        _cancellationToken.ThrowIfCancellationRequested();
        string logicalText = OfficeArabicTextShaper.ToLogicalText(text);
        OfficeTextShapingResult? result = provider.ShapeText(new OfficeTextShapingRequest(
            logicalText,
            font.DisplayName ?? style.Font.FamilyName,
            font.FontDataForShaping,
            isOpenTypeCff: false,
            font.UnitsPerEm,
            OfficeTextElements.ResolveBaseDirection(logicalText),
            _options.TextShapingLanguage,
            _cancellationToken,
            font.CollectionIndex,
            cloneFontData: false));
        if (result == null) return false;

        _ = font.CreateShapedTextRun(logicalText, result);
        return true;
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
            && run.RunningElementAssignment == null
            && run.Text.IndexOf('\u2028') < 0
            && run.Text.IndexOf('\n') < 0
            && run.Text.IndexOf('\r') < 0);
        int canonicalProgress = 0;
        bool canonicalHasContent = false;
        bool canonicalPreviousWasCollapsibleSpace = false;
        var lines = new List<InlineLine>();
        var line = new InlineLine();
        bool previousWasCollapsibleSpace = false;
        int noWrapRangeStart = -1;
        bool noWrapRangeStartedAfterContent = false;
        for (int runIndex = 0; runIndex < runs.Count; runIndex++) {
            HtmlInlineRun run = runs[runIndex];
            bool runPreventsWrapping = !paragraphStyle.PreventTextWrapping && run.Style.PreventTextWrapping;
            if (!runPreventsWrapping && noWrapRangeStart >= 0) {
                previousWasCollapsibleSpace = FinalizeNoWrapRange(
                    lines,
                    ref line,
                    noWrapRangeStart,
                    noWrapRangeStartedAfterContent,
                    width);
                noWrapRangeStart = -1;
                noWrapRangeStartedAfterContent = false;
            } else if (runPreventsWrapping && noWrapRangeStart < 0) {
                noWrapRangeStart = line.Segments.Count;
                noWrapRangeStartedAfterContent = line.HasFlowContent;
            }
            if (run.RunningStringElement != null) {
                line.Add(new InlineSegment(string.Empty, 0D, run));
                continue;
            }
            if (run.RunningElementAssignment != null) {
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
                if (!paragraphStyle.PreventTextWrapping
                    && !runPreventsWrapping
                    && line.HasFlowContent
                    && line.Width + atomicWidth > width) {
                    TrimTrailingWhitespace(line);
                    lines.Add(line);
                    line = new InlineLine();
                }

                line.Add(new InlineSegment(string.Empty, atomicWidth, run));
                continue;
            }

            int logicalOffset = 0;
            bool preserveWhitespace = run.Style.PreserveWhitespace;
            IReadOnlyList<string> tokens = Tokenize(run.Text, preserveWhitespace, run.Style.BreakSpaces).ToList();
            for (int tokenIndex = 0; tokenIndex < tokens.Count; tokenIndex++) {
                string token = tokens[tokenIndex];
                string logicalToken = SliceLogicalToken(run, token, ref logicalOffset);
                if (token == "\u2028" || preserveWhitespace && (token == "\n" || token == "\r\n")) {
                    if (noWrapRangeStart >= 0) {
                        FinalizeNoWrapRange(
                            lines,
                            ref line,
                            noWrapRangeStart,
                            noWrapRangeStartedAfterContent,
                            width);
                    }
                    lines.Add(line);
                    line = new InlineLine();
                    previousWasCollapsibleSpace = false;
                    noWrapRangeStart = runPreventsWrapping ? 0 : -1;
                    noWrapRangeStartedAfterContent = false;
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

                bool hasTabs = preserveWhitespace && normalizedToken.IndexOf('\t') >= 0;
                double tabExpandedWidth = hasTabs ? MeasureTabExpandedText(normalizedToken, run.Style, line.Width) : 0D;
                string paintToken = hasTabs ? normalizedToken.Replace("\t", string.Empty) : normalizedToken;
                HyphenationToken hyphenation = PrepareHyphenationToken(paintToken, normalizedLogicalToken, run.Style);
                paintToken = hyphenation.PaintText;
                string logicalPaintToken = hyphenation.LogicalText;
                double measured = hasTabs ? tabExpandedWidth : MeasureInlineText(paintToken, run.Style);
                bool preventTokenWrapping = paragraphStyle.PreventTextWrapping || runPreventsWrapping;
                if (!preventTokenWrapping
                    && !whitespace
                    && run.Style.WordBreak != "break-all"
                    && measured > Math.Max(0D, width - line.Width)
                    && TryAddHyphenatedToken(
                        lines,
                        ref line,
                        run,
                        hyphenation,
                        width,
                        visibleTokenStart,
                        tokenEnd,
                        !HasRemainingInlineFlowContent(runs, runIndex, tokens, tokenIndex))) {
                    continue;
                }
                if (!preventTokenWrapping
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
                if (!preventTokenWrapping
                    && !whitespace
                    && AllowsEmergencyTokenBreak(run.Style)
                    && (measured > width || breakAllIntoRemainingSpace)) {
                    AddBrokenToken(lines, ref line, run, paintToken, logicalPaintToken, width, visibleTokenStart);
                    continue;
                }

                if (!preventTokenWrapping && line.HasFlowContent && line.Width + measured > width) {
                    TrimTrailingWhitespace(line);
                    lines.Add(line);
                    line = new InlineLine();
                    if (whitespace && !preserveWhitespace) continue;
                }

                if (hasTabs) {
                    AddTabExpandedSegments(line, normalizedToken, normalizedLogicalToken, run, visibleTokenStart);
                    continue;
                }
                line.Add(new InlineSegment(paintToken, measured, run, logicalPaintToken, logicalEndProgress: tokenEnd));
            }
        }

        if (noWrapRangeStart >= 0) {
            FinalizeNoWrapRange(
                lines,
                ref line,
                noWrapRangeStart,
                noWrapRangeStartedAfterContent,
                width);
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
            && paragraphStyle.OverflowX != "visible") {
            foreach (InlineLine overflowingLine in lines.Where(candidate =>
                         candidate.Width > (candidate.HasExplicitPlacement ? candidate.AvailableWidth : width) + 0.0001D)) {
                int lineLogicalProgress = overflowingLine.Segments
                    .Select(segment => segment.LogicalEndProgress)
                    .DefaultIfEmpty(completeLogicalProgress)
                    .Max();
                ApplyEndEllipsis(overflowingLine, width, lineLogicalProgress);
            }
        }
        return RenderInlineLines(lines, width, paragraphStyle, formattingContainer, supportsContinuationReflow: supportsContinuationReflow);
    }

    private static bool HasRemainingInlineFlowContent(
        IReadOnlyList<HtmlInlineRun> runs,
        int runIndex,
        IReadOnlyList<string> tokens,
        int tokenIndex) {
        for (int index = tokenIndex + 1; index < tokens.Count; index++) {
            if (!string.IsNullOrWhiteSpace(tokens[index])) return true;
        }
        for (int index = runIndex + 1; index < runs.Count; index++) {
            HtmlInlineRun candidate = runs[index];
            if (candidate.AtomicBlock != null || candidate.FloatingBlock != null) return true;
            if (!string.IsNullOrWhiteSpace(candidate.Text)) return true;
        }
        return false;
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
        if (style.WordBreak != "break-all" && style.Hyphens == "auto" && _options.TextHyphenationCallback != null) {
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
        int logicalEndProgress,
        bool isFinalContentToken) {
        if (!token.HasBreaks || token.PaintText.Length != token.LogicalText.Length) return false;
        if (run.Style.HyphenateLimitLast == "always"
            && isFinalContentToken
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

    private double MeasureTabExpandedText(string value, HtmlRenderBoxStyle style, double currentWidth) {
        if (value.IndexOf('\t') < 0) return MeasureInlineText(value, style);
        double spaceWidth = Math.Max(0.01D, MeasureInlineText(" ", style));
        double stopWidth = style.TabSizeIsLength ? style.TabSize : style.TabSize * spaceWidth;
        double cursor = Math.Max(0D, currentWidth);
        foreach (char character in value) {
            if (character != '\t') {
                cursor += MeasureInlineText(character.ToString(), style);
                continue;
            }
            if (stopWidth <= 0D) continue;
            double nextStop = (Math.Floor(cursor / stopWidth) + 1D) * stopWidth;
            cursor = nextStop;
        }
        return Math.Max(0D, cursor - Math.Max(0D, currentWidth));
    }

    private void AddTabExpandedSegments(
        InlineLine line,
        string paintText,
        string logicalText,
        HtmlInlineRun run,
        int logicalStartProgress = 0) {
        double spaceWidth = Math.Max(0.01D, MeasureInlineText(" ", run.Style));
        double stopWidth = run.Style.TabSizeIsLength ? run.Style.TabSize : run.Style.TabSize * spaceWidth;
        int chunkStart = 0;
        int logicalProgress = logicalStartProgress;
        for (int index = 0; index < paintText.Length; index++) {
            if (paintText[index] != '\t') continue;
            if (index > chunkStart) {
                string paintChunk = paintText.Substring(chunkStart, index - chunkStart);
                string logicalChunk = logicalText.Substring(chunkStart, index - chunkStart);
                logicalProgress += logicalChunk.Length;
                line.Add(new InlineSegment(
                    paintChunk,
                    MeasureInlineText(paintChunk, run.Style),
                    run,
                    logicalChunk,
                    logicalEndProgress: logicalProgress));
            }

            double tabWidth = 0D;
            if (stopWidth > 0D) {
                double nextStop = (Math.Floor(line.Width / stopWidth) + 1D) * stopWidth;
                tabWidth = Math.Max(0D, nextStop - line.Width);
            }
            string logicalTab = logicalText.Substring(index, 1);
            logicalProgress += logicalTab.Length;
            line.Add(new InlineSegment(string.Empty, tabWidth, run, logicalTab, logicalEndProgress: logicalProgress));
            chunkStart = index + 1;
        }

        if (chunkStart >= paintText.Length) return;
        string finalPaint = paintText.Substring(chunkStart);
        string finalLogical = logicalText.Substring(chunkStart);
        logicalProgress += finalLogical.Length;
        line.Add(new InlineSegment(
            finalPaint,
            MeasureInlineText(finalPaint, run.Style),
            run,
            finalLogical,
            logicalEndProgress: logicalProgress));
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
        if (source.SemanticNodeId.HasValue) run.AssignSemanticNode(source.SemanticRole, source.SemanticNodeId.Value, source.BookmarkAnchorText, source.SemanticFragmentOrder);
        if (source.InlineSemanticGroupRole.HasValue && source.InlineSemanticGroupKey != null) {
            run.AssignInlineSemanticGroup(source.InlineSemanticGroupRole.Value, source.InlineSemanticGroupKey);
        }
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
            if (segment.Run.AtomicBlock == null
                && segment.Text.Length > 0
                && merged.Count > 0
                && merged[merged.Count - 1].Text.Length > 0
                && ReferenceEquals(merged[merged.Count - 1].Run, segment.Run)) {
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
        return measured;
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
            if (segment.Run.RunningStringElement != null || segment.Run.RunningElementAssignment != null) continue;
            if (!IsWhitespaceToken(segment.Text)) break;
            if (segment.Run.Style.BreakSpaces) break;
            line.RemoveAt(index);
        }
    }

    private static bool FinalizeNoWrapRange(
        ICollection<InlineLine> lines,
        ref InlineLine line,
        int rangeStart,
        bool startedAfterContent,
        double width) {
        if (startedAfterContent && rangeStart < line.Segments.Count && line.Width > width + 0.0001D) {
            var range = line.Segments.Skip(rangeStart).ToArray();
            while (line.Segments.Count > rangeStart) line.RemoveAt(rangeStart);
            TrimTrailingWhitespace(line);
            if (line.Segments.Count > 0) lines.Add(line);
            line = new InlineLine();
            foreach (InlineSegment segment in range) {
                if (!line.HasFlowContent
                    && segment.Run.RunningStringElement == null
                    && segment.Run.RunningElementAssignment == null
                    && IsWhitespaceToken(segment.Text)
                    && !segment.Run.Style.PreserveWhitespace) {
                    continue;
                }
                line.Add(segment);
            }
        }

        for (int index = line.Segments.Count - 1; index >= 0; index--) {
            InlineSegment segment = line.Segments[index];
            if (segment.Run.RunningStringElement != null || segment.Run.RunningElementAssignment != null) continue;
            return IsWhitespaceToken(segment.Text) && !segment.Run.Style.PreserveWhitespace;
        }
        return false;
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
            if (segment.Run.RunningStringElement == null && segment.Run.RunningElementAssignment == null) _flowContentCount++;
        }

        internal void RemoveAt(int index) {
            if (Segments[index].Run.RunningStringElement == null && Segments[index].Run.RunningElementAssignment == null) _flowContentCount--;
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
