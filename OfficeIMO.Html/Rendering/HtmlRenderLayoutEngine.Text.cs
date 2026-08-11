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

        if (tag != "img" && style.Display == "inline-block") {
            AddInlineBlockRun(element, width, inheritedStyle, depth, style, link, inheritedPaintOffsetX, inheritedPaintOffsetY, runs);
            return;
        }
        if (tag != "img" && style.Display == "inline-flex") {
            AddInlineFlexRun(element, width, inheritedStyle, depth, style, link, inheritedPaintOffsetX, inheritedPaintOffsetY, runs);
            return;
        }
        if (tag != "img" && style.Display == "inline-grid") {
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
            foreach (string token in Tokenize(run.Text, paragraphStyle.PreserveWhitespace)) {
                string logicalToken = SliceLogicalToken(run, token, ref logicalOffset);
                if (token == "\u2028" || paragraphStyle.PreserveWhitespace && (token == "\n" || token == "\r\n")) {
                    lines.Add(line);
                    line = new InlineLine();
                    previousWasCollapsibleSpace = false;
                    continue;
                }

                bool whitespace = IsWhitespaceToken(token);
                string normalizedToken = !paragraphStyle.PreserveWhitespace && whitespace ? " " : token;
                string normalizedLogicalToken = !paragraphStyle.PreserveWhitespace && whitespace ? " " : logicalToken;
                bool contributesCanonicalProgress = paragraphStyle.PreserveWhitespace
                    || !whitespace
                    || canonicalHasContent && !canonicalPreviousWasCollapsibleSpace;
                int tokenStart = canonicalProgress;
                if (contributesCanonicalProgress) canonicalProgress += normalizedLogicalToken.Length;
                int tokenEnd = canonicalProgress;
                if (!paragraphStyle.PreserveWhitespace) {
                    if (whitespace) {
                        canonicalPreviousWasCollapsibleSpace = true;
                    } else {
                        canonicalHasContent = true;
                        canonicalPreviousWasCollapsibleSpace = false;
                    }
                }

                if (!paragraphStyle.PreserveWhitespace && whitespace) {
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

                string paintToken = paragraphStyle.PreserveWhitespace && normalizedToken.IndexOf('\t') >= 0
                    ? ExpandTabs(normalizedToken, run.Style, line.Width)
                    : normalizedToken;
                double measured = MeasureText(paintToken, run.Style.Font);
                if (!paragraphStyle.PreventTextWrapping && !whitespace && measured > width && AllowsEmergencyTokenBreak(run.Style)) {
                    AddBrokenToken(lines, ref line, run, paintToken, normalizedLogicalToken, width, visibleTokenStart);
                    continue;
                }

                if (!paragraphStyle.PreventTextWrapping && line.HasFlowContent && line.Width + measured > width) {
                    TrimTrailingWhitespace(line);
                    lines.Add(line);
                    line = new InlineLine();
                    if (whitespace && !paragraphStyle.PreserveWhitespace) continue;
                }

                line.Add(new InlineSegment(paintToken, measured, run, normalizedLogicalToken, logicalEndProgress: tokenEnd));
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

    private string ExpandTabs(string value, HtmlRenderBoxStyle style, double currentWidth) {
        if (value.IndexOf('\t') < 0) return value;
        double spaceWidth = Math.Max(0.01D, MeasureText(" ", style.Font));
        double stopWidth = Math.Max(spaceWidth, style.TabSize * spaceWidth);
        double cursor = Math.Max(0D, currentWidth);
        var expanded = new StringBuilder();
        foreach (char character in value) {
            if (character != '\t') {
                expanded.Append(character);
                cursor += MeasureText(character.ToString(), style.Font);
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
        while (line.Segments.Count > 0) {
            InlineSegment segment = line.Segments[line.Segments.Count - 1];
            if (segment.Run.AtomicBlock != null || segment.Text.Length == 0) {
                line.RemoveAt(line.Segments.Count - 1);
                continue;
            }

            ellipsisRun = segment.Run;
            line.RemoveAt(line.Segments.Count - 1);
            double remainingWidth = Math.Max(0D, availableWidth - line.Width);
            double ellipsisWidth = MeasureText("\u2026", segment.Run.Style.Font);
            if (ellipsisWidth > remainingWidth + 0.0001D) continue;

            var paint = new StringBuilder();
            var logical = new StringBuilder();
            IReadOnlyList<string> paintElements = OfficeTextElements.Split(segment.Text);
            IReadOnlyList<string> logicalElements = OfficeTextElements.Split(segment.LogicalText);
            for (int index = 0; index < paintElements.Count; index++) {
                string candidate = paint.ToString() + paintElements[index];
                if (MeasureText(candidate, segment.Run.Style.Font) + ellipsisWidth > remainingWidth + 0.0001D) break;
                paint.Append(paintElements[index]);
                if (index < logicalElements.Count) logical.Append(logicalElements[index]);
            }

            string text = paint.ToString() + "\u2026";
            string logicalText = logical.ToString() + "\u2026";
            line.Add(new InlineSegment(
                text,
                MeasureText(text, segment.Run.Style.Font),
                segment.Run,
                logicalText,
                logicalEndProgress: completeLogicalProgress));
            return;
        }

        if (ellipsisRun != null) {
            double ellipsisWidth = MeasureText("\u2026", ellipsisRun.Style.Font);
            if (ellipsisWidth <= availableWidth + 0.0001D) {
                line.Add(new InlineSegment("\u2026", ellipsisWidth, ellipsisRun, "\u2026", logicalEndProgress: completeLogicalProgress));
            }
        }
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
        for (int index = 0; index < paintElements.Count; index++) {
            string value = paintElements[index];
            string logicalValue = index < logicalElements.Count ? logicalElements[index] : OfficeArabicTextShaper.ToLogicalText(value);
            double charWidth = MeasureText(value, run.Style.Font);
            if (part.Length > 0 && partWidth + charWidth > width) {
                if (line.HasFlowContent) {
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
                lines.Add(line);
                line = new InlineLine();
                part.Clear();
                logicalPart.Clear();
                partWidth = 0D;
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

    private string? ResolveSafeLink(string? rawHref, IElement element) {
        if (string.IsNullOrWhiteSpace(rawHref)) return null;
        string resolved = HtmlUrlPolicyEvaluator.ResolveUrl(rawHref, _baseUri, _options.UrlPolicy);
        if (resolved.Length > 0) return resolved;
        _diagnostics.Add(ComponentName, "HyperlinkRejectedByPolicy", "A hyperlink target was rejected before entering the rendered document.", HtmlDiagnosticSeverity.Warning, HtmlRenderStyleResolver.DescribeSource(element), rawHref);
        return null;
    }

    private static IEnumerable<string> Tokenize(string text, bool preserveWhitespace) {
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
                    ascent = Math.Max(ascent, run.AtomicBlock.Height);
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
                ascent = Math.Max(ascent, run.AtomicBlock?.Height ?? ResolveTextAscent(run.Style));
            }
            return ascent;
        }
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
