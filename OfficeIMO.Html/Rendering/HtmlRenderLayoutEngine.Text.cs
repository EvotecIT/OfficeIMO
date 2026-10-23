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
        HtmlListMarker? marker,
        IElement? generatedContentOwner,
        int skipLogicalCharacters = 0,
        PagedFloatBoundary? pageBoundary = null,
        IReadOnlyList<HtmlFloatExclusion>? inheritedFloats = null,
        bool applyTextIndent = true, InlineFloatContext? floatContext = null,
        bool includeGeneratedBefore = true, bool includeGeneratedAfter = true) {
        var runs = new List<HtmlInlineRun>();
        IElement? formattingContainer = generatedContentOwner ?? nodes.FirstOrDefault()?.ParentElement;
        if (marker != null && !marker.IsOutside) {
            runs.Add(CreateListMarkerRun(marker, generatedContentOwner));
        }

        double? containingHeight = ResolveContainingBlockHeight(parentStyle);
        string? inheritedLink = formattingContainer == null ? null : ResolveAncestorLink(formattingContainer);
        if (includeGeneratedBefore && generatedContentOwner != null && IncludeColumnGeneratedContent(generatedContentOwner, HtmlPseudoElementKind.Before)) {
            AddGeneratedInlineRun(generatedContentOwner, HtmlPseudoElementKind.Before, width, containingHeight, parentStyle, inheritedLink, 0D, 0D, runs);
        }

        foreach (INode node in nodes) {
            CollectInlineRuns(node, width, containingHeight, parentStyle, inheritedLink, depth, 0D, 0D, runs);
        }

        if (includeGeneratedAfter && generatedContentOwner != null && IncludeColumnGeneratedContent(generatedContentOwner, HtmlPseudoElementKind.After)) {
            AddGeneratedInlineRun(generatedContentOwner, HtmlPseudoElementKind.After, width, containingHeight, parentStyle, inheritedLink, 0D, 0D, runs);
        }

        // Suppress indentation of later anonymous blocks without changing the
        // inheritance context used above by descendant inline-block elements.
        if (!applyTextIndent) parentStyle = WithoutTextIndent(parentStyle);
        ApplyPendingInlineTextTransforms(runs, retainSuppressedText: true);
        ApplyFirstLetterStyle(formattingContainer, width, parentStyle, runs);
        RemoveSuppressedInlineText(runs);
        ApplyFirstLineStyle(formattingContainer, width, parentStyle, runs);
        runs = ApplyScopedFontFallbacks(runs);

        if (formattingContainer != null && ShouldAssignNavigationNode(parentStyle)) {
            int semanticNodeId = GetSemanticNodeId(formattingContainer);
            foreach (HtmlInlineRun run in runs) run.AssignSemanticNode(parentStyle.SemanticRole, semanticNodeId);
        }
        AssignSemanticFragmentOrders(runs);

        if (marker == null || !marker.IsOutside || skipLogicalCharacters > 0) {
            return LayoutInlineRuns(runs, width, parentStyle, formattingContainer, skipLogicalCharacters, pageBoundary, inheritedFloats, floatContext: floatContext);
        }

        HtmlRenderBoxStyle outsideMarkerStyle = marker.Style.Clone();
        outsideMarkerStyle.TextIndent = null;
        outsideMarkerStyle.PreserveWhitespace = true;
        outsideMarkerStyle.BreakSpaces = true;
        var markerRuns = new List<HtmlInlineRun> {
            CreateListMarkerRun(marker, generatedContentOwner, outsideMarkerStyle)
        };
        ApplyPendingInlineTextTransforms(markerRuns);
        markerRuns = ApplyScopedFontFallbacks(markerRuns);
        double gap = Math.Max(2D, parentStyle.Font.Size * 0.25D);
        double markerAdvance = marker.Image?.Width ?? MeasureInlineText(marker.Content, outsideMarkerStyle);
        double gutter = Math.Min(Math.Max(1D, width * 0.5D), markerAdvance + gap);
        HtmlInlineLayout markerLayout = LayoutInlineRuns(markerRuns, gutter, outsideMarkerStyle, formattingContainer);
        HtmlInlineLayout bodyLayout = LayoutInlineRuns(runs, width, parentStyle, formattingContainer, skipLogicalCharacters, pageBoundary, inheritedFloats, floatContext: floatContext);
        return CombineOutsideListMarker(markerLayout, bodyLayout, width, gap, parentStyle);
    }

    private static HtmlInlineLayout CombineOutsideListMarker(
        HtmlInlineLayout marker,
        HtmlInlineLayout body,
        double width,
        double gap,
        HtmlRenderBoxStyle style) {
        var visuals = new List<HtmlRenderVisual>(marker.Visuals.Count + body.Visuals.Count);
        if (marker.Visuals.Count > 0) {
            (double markerX, _, double markerWidth, _) = ResolveSemanticBounds(marker.Visuals, width, Math.Max(marker.Height, style.LineHeight));
            double offsetX = string.Equals(style.Direction, "rtl", StringComparison.OrdinalIgnoreCase)
                ? width + gap - markerX
                : -gap - markerX - markerWidth;
            foreach (HtmlRenderVisual visual in marker.Visuals) {
                visuals.Add(visual.Translate(offsetX, 0D, visuals.Count));
            }
        }
        foreach (HtmlRenderVisual visual in body.Visuals) {
            visuals.Add(visual.Translate(0D, 0D, visuals.Count));
        }
        return new HtmlInlineLayout(
            visuals,
            Math.Max(marker.Height, body.Height),
            body.BreakOffsets,
            body.RunningStringAssignments,
            body.BreakProgress,
            body.SupportsContinuationReflow,
            lineBreakOffsets: body.LineBreakOffsets,
            floatExclusions: body.FloatExclusions,
            interruptedFlow: body.InterruptedFlow?.WithVisuals(visuals, width: width, height: Math.Max(marker.Height, body.Height)),
            forcedBreaks: body.ForcedBreaks,
            lineBreakGroups: body.LineBreakGroups,
            continuationGroups: body.ContinuationGroups,
            trailingGroups: body.TrailingGroups,
            pagedPaintExtent: Math.Max(marker.Height, body.PagedPaintExtent),
            avoidBreakRanges: body.AvoidBreakRanges,
            safeBreakRanges: body.SafeBreakRanges);
    }

    /// <summary>Keeps default marker text at its list item's source entry across every marker layout route.</summary>
    private HtmlInlineRun CreateListMarkerRun(HtmlListMarker marker, IElement? owner, HtmlRenderBoxStyle? style = null) {
        HtmlRenderBoxStyle effectiveStyle = style ?? marker.Style;
        HtmlInlineRun run = marker.IsImage
            ? new HtmlInlineRun(marker.Image!, effectiveStyle, null, "list-marker", ownerElement: owner, isReplacedImage: true)
            : new HtmlInlineRun(marker.Content, effectiveStyle, null, "list-marker", ownerElement: owner);
        run.SemanticSourceOrder = owner == null ? null : GetSemanticSourceOrder(owner);
        return run;
    }

    private void AddOutsideMarkerForBlockChildren(
        ICollection<HtmlRenderVisual> visuals,
        HtmlListMarker marker,
        double contentWidth,
        HtmlRenderBoxStyle itemStyle,
        IElement owner) {
        HtmlRenderBoxStyle markerStyle = marker.Style.Clone();
        markerStyle.PreserveWhitespace = true;
        markerStyle.BreakSpaces = true;
        var runs = new List<HtmlInlineRun> { CreateListMarkerRun(marker, owner) };
        ApplyPendingInlineTextTransforms(runs);
        runs = ApplyScopedFontFallbacks(runs);
        double markerAdvance = marker.Image?.Width ?? MeasureInlineText(marker.Content, markerStyle);
        double gutter = Math.Max(1D, markerAdvance + Math.Max(2D, itemStyle.Font.Size * 0.25D));
        HtmlInlineLayout layout = LayoutInlineRuns(runs, gutter, markerStyle, owner);
        if (layout.Visuals.Count == 0) return;

        (double markerX, double markerY, double markerWidth, _) = ResolveSemanticBounds(
            layout.Visuals, gutter, Math.Max(layout.Height, markerStyle.LineHeight));
        double gap = Math.Max(2D, itemStyle.Font.Size * 0.25D);
        double offsetX = string.Equals(itemStyle.Direction, "rtl", StringComparison.OrdinalIgnoreCase)
            ? contentWidth + gap - markerX
            : -gap - markerX - markerWidth;
        double? firstTextY = FindFirstBodyTextY(visuals);
        double offsetY = (firstTextY ?? 0D) - markerY;
        foreach (HtmlRenderVisual visual in layout.Visuals) {
            visuals.Add(visual.Translate(offsetX, offsetY, visuals.Count));
        }
    }

    private static double? FindFirstBodyTextY(IEnumerable<HtmlRenderVisual> visuals) {
        foreach (HtmlRenderVisual visual in visuals) {
            if (visual is HtmlRenderText text && text.Source != "list-marker") return text.Y;
            IReadOnlyList<HtmlRenderVisual>? children = GetGroupChildren(visual);
            if (children == null) continue;
            double? childY = FindFirstBodyTextY(children);
            if (childY.HasValue) return childY;
        }
        return null;
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

    private void AssignLogicalTextOrders(IEnumerable<HtmlInlineRun> runs) {
        foreach (HtmlInlineRun run in runs) {
            if (run.IsTextTransformContextOnly || run.Text.Length == 0 || run.LogicalTextOrder.HasValue) continue;
            run.AssignLogicalTextOrder(_nextLogicalTextOrder++);
        }
    }

    private void ReportUnsupportedComplexTextShaping(string text, IElement? element, HtmlRenderBoxStyle style, string? source = null) {
        if (element == null || string.IsNullOrWhiteSpace(text) || _reportedComplexTextShapingElements.Contains(element)) return;
        bool featureShaping = !style.TextFeatureSettings.IsDefault;
        if (!featureShaping && !RequiresConfiguredTextShaping(text, style.TextFeatureSettings)) return;
        if (featureShaping && _options.TextShapingProvider != null && !RequiresConfiguredTextShaping(text, style.TextFeatureSettings)) return;
        IReadOnlyList<OfficeFontFallbackRun> fallbackRuns = _fonts.PlanFallbackRuns(
            text,
            style.Font.FamilyName,
            style.FontDescriptor);
        bool allUnsupportedRunsShaped = true;
        foreach (OfficeFontFallbackRun fallback in fallbackRuns) {
            if (!featureShaping && !RequiresConfiguredTextShaping(fallback.Text, style.TextFeatureSettings)) continue;
            HtmlRenderBoxStyle fallbackStyle = style.Clone();
            fallbackStyle.Font = fallbackStyle.Font.WithFamilyName(InlineFallbackFamilyRequest(fallback));
            if (!TryMeasureWithConfiguredProvider(fallback.Text, fallbackStyle, out _)) {
                allUnsupportedRunsShaped = false;
                break;
            }
        }
        if (allUnsupportedRunsShaped) return;
        _reportedComplexTextShapingElements.Add(element);
        string diagnosticCode = featureShaping
            ? HtmlRenderDiagnosticCodes.OpenTypeFeatureUnsupported
            : HtmlRenderDiagnosticCodes.ComplexTextShapingUnsupported;
        string message = featureShaping
            ? "A requested OpenType feature required a lookup outside the bounded managed shaping subset; scalar glyphs were used."
            : "The selected shaping provider did not accept this complex-script run; diagnosed scalar fallback was used.";
        _diagnostics.Add(
            ComponentName,
            diagnosticCode,
            message,
            HtmlDiagnosticSeverity.Warning,
            source ?? HtmlRenderStyleResolver.DescribeSource(element),
            "provider-declined",
            OfficeConversionLossKind.Approximation);
    }

    private static bool RequiresConfiguredTextShaping(string text, OfficeTextFeatureSettings featureSettings) =>
        OfficeTextElements.ContainsCombiningMark(text)
        || OfficeArabicTextShaper.RequiresLamAlefLigatureShaping(text, featureSettings)
        || OfficeTextElements.ContainsShapingRequiredScript(text)
        || OfficeTextElements.ContainsVariationSelector(text)
        || OfficeTextElements.ContainsJoiningScript(text);

    private bool TryMeasureWithConfiguredProvider(string text, HtmlRenderBoxStyle style, out double measured) {
        IOfficeTextShapingProvider? provider = _options.TextShapingProvider;
        if (provider == null && (!style.TextFeatureSettings.IsDefault || OfficeManagedTextShaper.RequiresComplexLayout(text)))
            provider = OfficeManagedTextShapingProvider.Instance;
        if (provider == null || provider == OfficeManagedTextShapingProvider.Instance && style.WritingMode != "horizontal-tb") {
            measured = 0D;
            return false;
        }
        ChargeIntrinsicTextMeasurement(text);
        OfficeFontInfo effectiveFont = GetEffectiveTextFont(style);
        // Match the exact language forwarded by ShapeConfiguredText, including the
        // global override. Custom providers may distinguish any supplied values.
        string? shapingLanguage = _options.TextShapingLanguage ?? style.Language;
        if (_shapedTextMeasurementCache.TryGet(text, effectiveFont, style.TextFeatureSettings, shapingLanguage, out measured)) return true;

        IOfficeFontProgram? font = _fonts.ResolveForText(
            text,
            effectiveFont.FamilyName,
            effectiveFont.Style,
            effectiveFont.Size,
            out OfficeFontStyle _);
        if (font == null) {
            measured = 0D;
            return false;
        }

        _cancellationToken.ThrowIfCancellationRequested();
        string logicalText = OfficeArabicTextShaper.ToLogicalText(text);
        OfficeTextShapingResult? result = ShapeConfiguredText(logicalText, font, effectiveFont, style,
            OfficeTextElements.ResolveBaseDirection(logicalText), provider);
        if (result == null) {
            measured = 0D;
            return false;
        }

        measured = font.MeasureShapedText(logicalText, result, effectiveFont.Size);
        _shapedTextMeasurementCache.Store(text, effectiveFont, style.TextFeatureSettings, shapingLanguage, measured);
        return true;
    }

    private static bool HasInlineBoxPaint(HtmlRenderBoxStyle style) =>
        style.BackgroundColor.HasValue && style.BackgroundColor.Value.A > 0
        || style.BackgroundImageLayerCount > 0
        || style.Borders.HasPaint
        || style.BoxShadowLayerCount > 0
        || style.OutlineWidth > 0D && style.OutlineStyle != "none";

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
            if ((candidate.AtomicBlock != null && !candidate.IsFlowMarker) || candidate.FloatingBlock != null) return true;
            if (!string.IsNullOrWhiteSpace(candidate.Text)) return true;
        }
        return false;
    }

    private HyphenationToken PrepareHyphenationToken(string paintToken, string logicalToken, HtmlRenderBoxStyle style,
        bool reportUnsupportedLanguage = true) {
        if (reportUnsupportedLanguage && style.Hyphens == "auto" && style.WordBreak != "break-all"
            && _options.TextHyphenationCallback == null
            && !OfficeTextHyphenationPatterns.SupportsLanguage(style.Language)
            && logicalToken.Any(char.IsLetter)
            && _reportedHyphenationLanguages.Add(style.Language ?? string.Empty)) {
            _diagnostics.Add(ComponentName, "HyphenationLanguageUnsupported",
                "Automatic hyphenation has no embedded patterns for the effective language; only authored soft hyphens remain available.",
                HtmlDiagnosticSeverity.Warning, style.Language ?? string.Empty, null,
                OfficeConversionLossKind.Approximation);
        }
        if (paintToken.IndexOf('\u00AD') < 0
            && (style.Hyphens != "auto" || (_options.TextHyphenationCallback == null
                && !OfficeTextHyphenationPatterns.SupportsLanguage(style.Language)))) {
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
        if (style.WordBreak != "break-all" && style.Hyphens == "auto") {
            IReadOnlyList<int>? automatic = _options.TextHyphenationCallback != null
                ? _options.TextHyphenationCallback(logical.ToString())
                : OfficeTextHyphenationPatterns.GetBreakpoints(logical.ToString(), style.Language);
            if (automatic != null) {
                foreach (int point in automatic) automaticBreaks.Add(point);
            }
        }

        string logicalText = logical.ToString();
        int minimumWordLength = Math.Max(1, style.HyphenateMinimumWordLength);
        int minimumPrefix = Math.Max(1, style.HyphenateMinimumPrefixLength);
        int minimumSuffix = Math.Max(1, style.HyphenateMinimumSuffixLength);
        int[] graphemeStarts = StringInfo.ParseCombiningCharacters(logicalText);
        int[] characterCounts = new int[graphemeStarts.Length + 1];
        for (int index = 0; index < graphemeStarts.Length; index++) {
            UnicodeCategory category = CharUnicodeInfo.GetUnicodeCategory(logicalText, graphemeStarts[index]);
            characterCounts[index + 1] = characterCounts[index] +
                (category == UnicodeCategory.NonSpacingMark || IsPunctuationCategory(category) ? 0 : 1);
        }
        int totalCharacters = characterCounts[graphemeStarts.Length];
        int[] FilterBreaks(IEnumerable<int> candidates) => totalCharacters < minimumWordLength
            ? Array.Empty<int>()
            : candidates
                .Where(point => point > 0 && point < logicalText.Length)
                .Select(point => Array.BinarySearch(graphemeStarts, point))
                .Where(index => index > 0 && characterCounts[index] >= minimumPrefix
                    && totalCharacters - characterCounts[index] >= minimumSuffix)
                .Select(index => graphemeStarts[index])
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

    private static bool IsPunctuationCategory(UnicodeCategory category) => category == UnicodeCategory.ConnectorPunctuation
        || category == UnicodeCategory.DashPunctuation
        || category == UnicodeCategory.OpenPunctuation
        || category == UnicodeCategory.ClosePunctuation
        || category == UnicodeCategory.InitialQuotePunctuation
        || category == UnicodeCategory.FinalQuotePunctuation
        || category == UnicodeCategory.OtherPunctuation;

    private double MeasureText(string value, OfficeFontInfo font, OfficeFontFaceDescriptor descriptor) {
        ChargeIntrinsicTextMeasurement(value);
        _fontUsage?.Observe(value, font.FamilyName, descriptor);
        if (_fonts.TryMeasureText(value, font.Size, font.FamilyName, font.Style, out double scopedWidth)) {
            return scopedWidth;
        }

        OfficeTextMeasurer measurer = OfficeTextMeasurer.Create(font);
        if (_options.FallbackTextMeasurement != null) {
            double? measured = _options.FallbackTextMeasurement(value, font, descriptor);
            if (measured.HasValue) return measured.Value;
        }
        OfficeTextMeasurementStyle style = measurer.CreateStyle(font, 72D);
        return measurer.MeasureWidth(value, style);
    }

    private double MeasureInlineText(string value, HtmlRenderBoxStyle style) {
        if (style.Font.Size <= 0D) return 0D;
        _fontUsage?.Observe(value, style.Font.FamilyName, style.FontDescriptor);
        double measured = TryMeasureWithConfiguredProvider(value, style, out double shapedWidth)
            ? shapedWidth
            : MeasureText(value, GetEffectiveTextFont(style), style.FontDescriptor);
        if (Math.Abs(style.LetterSpacing) <= 0.000001D && Math.Abs(style.WordSpacing) <= 0.000001D) {
            return Math.Max(0.01D, measured);
        }
        IReadOnlyList<string> elements = OfficeTextElements.Split(value);
        if (elements.Count == 0) return measured;
        measured += style.LetterSpacing * elements.Count;
        // Word spacing includes no-break separators independently of collapsing.
        measured += style.WordSpacing * elements.Count(IsWordSpacingSeparator);
        return measured;
    }

    private double MeasureInlineText(string value, HtmlRenderBoxStyle style, HtmlInlineRun run) =>
        run.GetPreparedShaping(value, style) is OfficePreparedTextShaping prepared
            ? MeasurePreparedInlineText(value, style, prepared)
            : MeasureInlineText(value, style);

    private double MeasureInlineText(string value, HtmlRenderBoxStyle style, IntrinsicTextRun run) =>
        run.PreparedShaping != null && string.Equals(value, run.PreparedShaping.Text, StringComparison.Ordinal)
            ? MeasurePreparedInlineText(value, style, run.PreparedShaping)
            : MeasureInlineText(value, style);

    private double MeasurePreparedInlineText(string value, HtmlRenderBoxStyle style, OfficePreparedTextShaping prepared) {
        ChargeIntrinsicTextMeasurement(value);
        _fontUsage?.Observe(value, style.Font.FamilyName, style.FontDescriptor);
        return prepared.Measure(GetEffectiveTextFont(style).Size);
    }

    private string? ResolveSafeLink(string? rawHref, IElement element) {
        if (rawHref == null) return null;
        string candidate = rawHref.Trim();
        if (candidate.Length == 0) return null;
        if (candidate.Length > 1 && candidate[0] == '#') {
            if (!TryDecodeFragment(candidate.Substring(1), out string fragment)) {
                return ReportUnavailableFragment(element, candidate);
            }
            if (string.Equals(fragment, "top", StringComparison.OrdinalIgnoreCase)) return "#top";
            if (!_namedDestinationTargets.ContainsKey(fragment)) return ReportUnavailableFragment(element, candidate);
            return "#" + fragment;
        }
        string resolved = HtmlUrlPolicyEvaluator.ResolveUrl(rawHref, _baseUri, _options.UrlPolicy);
        if (resolved.Length > 0) return resolved;
        _diagnostics.Add(ComponentName, "HyperlinkRejectedByPolicy", "A hyperlink target was rejected before entering the rendered document.", HtmlDiagnosticSeverity.Warning, HtmlRenderStyleResolver.DescribeSource(element), rawHref);
        return null;
    }

    private string? ReportUnavailableFragment(IElement element, string candidate) {
        _diagnostics.Add(
            ComponentName,
            HtmlRenderDiagnosticCodes.HyperlinkTargetUnavailable,
            "A document-internal hyperlink target was not found and the link annotation was omitted.",
            HtmlDiagnosticSeverity.Warning,
            HtmlRenderStyleResolver.DescribeSource(element),
            candidate,
            OfficeConversionLossKind.Omission);
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

            bool currentWhitespace = IsCollapsibleCssWhitespace(current);
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

    private static bool IsWhitespaceToken(string token) => token.Length > 0 && token.All(IsCollapsibleCssWhitespace);

    // Word spacing applies to separators independently of CSS collapsing and
    // line-breaking rules, including nonbreaking spaces but excluding fixed-width spaces.
    private static bool IsWordSpacingSeparator(string textElement) => textElement.Length > 0
        && textElement.All(static value => char.IsWhiteSpace(value)
            && (value < '\u2000' || value > '\u200a')
            && value != '\u3000');

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
        int strutRangeStart,
        bool startedAfterContent,
        double width) {
        if (startedAfterContent && rangeStart < line.Segments.Count && line.Width > line.ResolveAvailableWidth(width) + 0.0001D) {
            var range = line.Segments.Skip(rangeStart).ToArray();
            IReadOnlyList<HtmlRenderBoxStyle> rangeStruts = line.TakeEmptyInlineStruts(strutRangeStart);
            line.RemoveRange(rangeStart, line.Segments.Count - rangeStart);
            TrimTrailingWhitespace(line);
            if (line.Segments.Count > 0) lines.Add(line);
            line = new InlineLine();
            line.RecordEmptyInlineStruts(rangeStruts);
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


    internal readonly struct HyphenationToken {
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
        internal HyphenationToken SliceSource(int sourceStart, int sourceLength) {
            int sourceEnd = sourceStart + sourceLength;
            IReadOnlyList<int> boundaries = SourceBoundaries;
            int start = boundaries.Count == 0 ? sourceStart
                : Enumerable.Range(0, boundaries.Count).First(index => boundaries[index] >= sourceStart);
            int end = boundaries.Count == 0 ? sourceEnd
                : Enumerable.Range(start, boundaries.Count - start).First(index => boundaries[index] >= sourceEnd);
            int[] SliceBreaks(IReadOnlyList<int> points) => points.Where(point => point > start && point < end)
                .Select(point => point - start).ToArray();
            return new HyphenationToken(PaintText.Substring(start, end - start), LogicalText.Substring(start, end - start),
                SliceBreaks(PrimaryBreaks), SliceBreaks(SecondaryBreaks),
                SourceBoundaries.Count == 0 ? Array.Empty<int>()
                    : SourceBoundaries.Skip(start).Take(end - start + 1).Select(point => point - sourceStart).ToArray());
        }

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
            int logicalEndProgress = 0,
            OfficeTextDirection? resolvedDirection = null) {
            Text = text;
            LogicalText = logicalText ?? text;
            Width = width;
            Run = run;
            BidiResolved = bidiResolved;
            LogicalEndProgress = logicalEndProgress;
            ResolvedDirection = resolvedDirection;
            PreparedShaping = run.GetPreparedShaping(LogicalText, run.Style);
        }

        internal string Text { get; }
        internal string LogicalText { get; }
        internal double Width { get; private set; }
        internal double LeadingAdvance { get; set; }
        internal double TrailingAdvance { get; set; }
        internal double Advance => LeadingAdvance + Width + TrailingAdvance;
        internal HtmlInlineRun Run { get; }
        internal bool BidiResolved { get; }
        internal OfficeTextDirection? ResolvedDirection { get; }
        internal OfficePreparedTextShaping? PreparedShaping { get; }
        internal int LogicalEndProgress { get; }
        internal void SetWidth(double width) => Width = Math.Max(0D, width);
    }

}
