using System.Text;
using AngleSharp.Dom;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private HtmlInlineLayout LayoutInlineNodes(IEnumerable<INode> nodes, double width, HtmlRenderBoxStyle parentStyle, int depth, string? prefix, IElement? generatedContentOwner) {
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

        return LayoutInlineRuns(runs, width, parentStyle, formattingContainer);
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
        bool bidiControl = OfficeTextElements.ContainsBidiControl(textNode.Data);
        if (!joiningScript && !bidiControl) return;
        _reportedBidiElements.Add(element);
        _diagnostics.Add(
            ComponentName,
            joiningScript ? HtmlRenderDiagnosticCodes.ComplexTextShapingUnsupported : HtmlRenderDiagnosticCodes.BidiLayoutUnsupported,
            joiningScript
                ? "A joining script outside the bounded core-Arabic shaper used scalar glyphs."
                : "Explicit Unicode bidi controls require an embedding or isolate stage that is not active yet.",
            HtmlDiagnosticSeverity.Warning,
            HtmlRenderStyleResolver.DescribeSource(element),
            joiningScript ? "joining-script" : "bidi-control");
    }

    private HtmlInlineLayout LayoutInlineRuns(IReadOnlyList<HtmlInlineRun> runs, double width, HtmlRenderBoxStyle paragraphStyle, IElement? formattingContainer = null) {
        if (runs.Count == 0 || width <= 0D) return new HtmlInlineLayout(Array.Empty<HtmlRenderVisual>(), 0D);
        if (runs.Any(run => run.FloatingBlock != null)) {
            return LayoutInlineRunsWithFloats(runs, width, paragraphStyle, formattingContainer);
        }
        var lines = new List<InlineLine>();
        var line = new InlineLine();
        bool previousWasCollapsibleSpace = false;
        foreach (HtmlInlineRun run in runs) {
            if (run.PositionedMarkerElement != null) {
                line.Add(new InlineSegment(string.Empty, 0D, run));
                previousWasCollapsibleSpace = false;
                continue;
            }
            if (run.AtomicBlock != null) {
                previousWasCollapsibleSpace = false;
                double atomicWidth = run.AtomicBlock.Width;
                if (line.Segments.Count > 0 && line.Width + atomicWidth > width) {
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
                if (!paragraphStyle.PreserveWhitespace && whitespace) {
                    if (line.Segments.Count == 0 || previousWasCollapsibleSpace) continue;
                    previousWasCollapsibleSpace = true;
                } else {
                    previousWasCollapsibleSpace = false;
                }

                double measured = MeasureText(normalizedToken, run.Style.Font);
                if (!whitespace && measured > width && AllowsEmergencyTokenBreak(run.Style)) {
                    AddBrokenToken(lines, ref line, run, normalizedToken, normalizedLogicalToken, width);
                    continue;
                }

                if (line.Segments.Count > 0 && line.Width + measured > width) {
                    TrimTrailingWhitespace(line);
                    lines.Add(line);
                    line = new InlineLine();
                    if (whitespace && !paragraphStyle.PreserveWhitespace) continue;
                }

                line.Add(new InlineSegment(normalizedToken, measured, run, normalizedLogicalToken));
            }
        }

        TrimTrailingWhitespace(line);
        if (line.Segments.Count > 0 || lines.Count == 0) lines.Add(line);
        return RenderInlineLines(lines, width, paragraphStyle, formattingContainer);
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
                merged[merged.Count - 1] = new InlineSegment(previous.Text + segment.Text, previous.Width + segment.Width, previous.Run, previous.LogicalText + segment.LogicalText);
            } else {
                merged.Add(segment);
            }
        }

        return merged;
    }

    private void AddBrokenToken(ICollection<InlineLine> lines, ref InlineLine line, HtmlInlineRun run, string token, string logicalToken, double width) {
        var part = new StringBuilder();
        var logicalPart = new StringBuilder();
        double partWidth = 0D;
        IReadOnlyList<string> paintElements = OfficeTextElements.Split(token);
        IReadOnlyList<string> logicalElements = OfficeTextElements.Split(logicalToken);
        for (int index = 0; index < paintElements.Count; index++) {
            string value = paintElements[index];
            string logicalValue = index < logicalElements.Count ? logicalElements[index] : OfficeArabicTextShaper.ToLogicalText(value);
            double charWidth = MeasureText(value, run.Style.Font);
            if (part.Length > 0 && partWidth + charWidth > width) {
                if (line.Segments.Count > 0) {
                    TrimTrailingWhitespace(line);
                    lines.Add(line);
                    line = new InlineLine();
                }

                line.Add(new InlineSegment(part.ToString(), partWidth, run, logicalPart.ToString()));
                lines.Add(line);
                line = new InlineLine();
                part.Clear();
                logicalPart.Clear();
                partWidth = 0D;
            }

            part.Append(value);
            logicalPart.Append(logicalValue);
            partWidth += charWidth;
        }

        if (part.Length > 0) {
            if (line.Segments.Count > 0 && line.Width + partWidth > width) {
                TrimTrailingWhitespace(line);
                lines.Add(line);
                line = new InlineLine();
            }

            line.Add(new InlineSegment(part.ToString(), partWidth, run, logicalPart.ToString()));
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
        while (line.Segments.Count > 0 && IsWhitespaceToken(line.Segments[line.Segments.Count - 1].Text)) {
            line.RemoveLast();
        }
    }

    private static double ResolveLineOffset(OfficeTextAlignment alignment, double width, double lineWidth) {
        if (alignment == OfficeTextAlignment.Center) return Math.Max(0D, (width - lineWidth) / 2D);
        if (alignment == OfficeTextAlignment.Right) return Math.Max(0D, width - lineWidth);
        return 0D;
    }

    private sealed class InlineLine {
        internal List<InlineSegment> Segments { get; } = new List<InlineSegment>();
        internal double Width { get; private set; }
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
        }

        internal void RemoveLast() {
            int index = Segments.Count - 1;
            Width -= Segments[index].Width;
            Segments.RemoveAt(index);
        }

        internal double ResolveLineHeight(double fallback) {
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
        internal InlineSegment(string text, double width, HtmlInlineRun run, string? logicalText = null) {
            Text = text;
            LogicalText = logicalText ?? text;
            Width = width;
            Run = run;
        }

        internal string Text { get; }
        internal string LogicalText { get; }
        internal double Width { get; }
        internal HtmlInlineRun Run { get; }
    }

    private static double ResolveTextAscent(HtmlRenderBoxStyle style) {
        double leading = Math.Max(0D, style.LineHeight - style.Font.Size);
        return Math.Min(style.LineHeight, leading / 2D + style.Font.Size * 0.8D);
    }
}
