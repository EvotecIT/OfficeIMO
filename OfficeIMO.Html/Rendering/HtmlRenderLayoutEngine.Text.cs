using System.Text;
using AngleSharp.Dom;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private HtmlInlineLayout LayoutInlineNodes(IEnumerable<INode> nodes, double width, HtmlRenderBoxStyle parentStyle, int depth, string? prefix, IElement? generatedContentOwner) {
        var runs = new List<HtmlInlineRun>();
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

        return LayoutInlineRuns(runs, width, parentStyle);
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
                runs.Add(new HtmlInlineRun(ApplyTextTransform(textNode.Data, inheritedStyle.TextTransform), inheritedStyle, inheritedLink, inheritedStyle.SemanticRole, inheritedPaintOffsetX, inheritedPaintOffsetY));
            }

            return;
        }

        if (!(node is IElement element) || ShouldSkipElement(element)) return;
        string tag = element.TagName.ToLowerInvariant();
        if (tag == "br") {
            runs.Add(new HtmlInlineRun("\u2028", inheritedStyle, inheritedLink, HtmlRenderStyleResolver.DescribeSource(element), inheritedPaintOffsetX, inheritedPaintOffsetY));
            return;
        }

        HtmlRenderBoxStyle style = _styleResolver.Resolve(element, width, inheritedStyle);
        if (style.Display == "none") return;
        ResolvePositionPaintOffset(style, width, containingHeight, HtmlRenderStyleResolver.DescribeSource(element), out double elementPaintOffsetX, out double elementPaintOffsetY);
        double paintOffsetX = inheritedPaintOffsetX + elementPaintOffsetX;
        double paintOffsetY = inheritedPaintOffsetY + elementPaintOffsetY;
        string? link = inheritedLink;
        if (tag == "a") {
            link = ResolveSafeLink(element.GetAttribute("href"), element);
        }

        AddGeneratedInlineRun(element, HtmlPseudoElementKind.Before, width, containingHeight, style, link, paintOffsetX, paintOffsetY, runs);

        if (tag == "img") {
            string alternativeText = element.GetAttribute("alt") ?? string.Empty;
            if (alternativeText.Length > 0) runs.Add(new HtmlInlineRun(alternativeText, style, link, HtmlRenderStyleResolver.DescribeSource(element), paintOffsetX, paintOffsetY));
            AddUnsupported(HtmlRenderDiagnosticCodes.InlineImageFallback, "An inline image was represented by its alternative text; block image layout is supported separately.", element);
            return;
        }

        foreach (INode child in element.ChildNodes) {
            CollectInlineRuns(child, width, containingHeight, style, link, depth + 1, paintOffsetX, paintOffsetY, runs);
        }

        AddGeneratedInlineRun(element, HtmlPseudoElementKind.After, width, containingHeight, style, link, paintOffsetX, paintOffsetY, runs);
    }

    private HtmlInlineLayout LayoutInlineRuns(IReadOnlyList<HtmlInlineRun> runs, double width, HtmlRenderBoxStyle paragraphStyle) {
        if (runs.Count == 0 || width <= 0D) return new HtmlInlineLayout(Array.Empty<HtmlRenderVisual>(), 0D);
        var lines = new List<InlineLine>();
        var line = new InlineLine();
        bool previousWasCollapsibleSpace = false;
        foreach (HtmlInlineRun run in runs) {
            foreach (string token in Tokenize(run.Text, paragraphStyle.PreserveWhitespace)) {
                if (token == "\u2028" || paragraphStyle.PreserveWhitespace && (token == "\n" || token == "\r\n")) {
                    lines.Add(line);
                    line = new InlineLine();
                    previousWasCollapsibleSpace = false;
                    continue;
                }

                bool whitespace = IsWhitespaceToken(token);
                string normalizedToken = !paragraphStyle.PreserveWhitespace && whitespace ? " " : token;
                if (!paragraphStyle.PreserveWhitespace && whitespace) {
                    if (line.Segments.Count == 0 || previousWasCollapsibleSpace) continue;
                    previousWasCollapsibleSpace = true;
                } else {
                    previousWasCollapsibleSpace = false;
                }

                double measured = MeasureText(normalizedToken, run.Style.Font);
                if (!whitespace && measured > width) {
                    AddBrokenToken(lines, ref line, run, normalizedToken, width);
                    continue;
                }

                if (line.Segments.Count > 0 && line.Width + measured > width) {
                    TrimTrailingWhitespace(line);
                    lines.Add(line);
                    line = new InlineLine();
                    if (whitespace && !paragraphStyle.PreserveWhitespace) continue;
                }

                line.Add(new InlineSegment(normalizedToken, measured, run));
            }
        }

        TrimTrailingWhitespace(line);
        if (line.Segments.Count > 0 || lines.Count == 0) lines.Add(line);

        var visuals = new List<HtmlRenderVisual>();
        var breakOffsets = new List<double>();
        double y = 0D;
        foreach (InlineLine current in lines) {
            double lineHeight = current.ResolveLineHeight(paragraphStyle.LineHeight);
            double offsetX = ResolveLineOffset(paragraphStyle.Alignment, width, current.Width);
            double x = offsetX;
            foreach (InlineSegment segment in MergeAdjacentInlineSegments(current.Segments)) {
                if (segment.Text.Length > 0 && segment.Width > 0D) {
                    double frameTolerance = Math.Max(1D, segment.Run.Style.Font.Size * 0.35D);
                    double frameWidth = Math.Min(Math.Max(0.01D, width - x), segment.Width + frameTolerance);
                    HtmlRenderVisual visual = new HtmlRenderText(
                        segment.Text,
                        x,
                        y,
                        Math.Max(0.01D, frameWidth),
                        Math.Max(0.01D, lineHeight),
                        segment.Run.Style.Font,
                        segment.Run.Style.Color,
                        OfficeTextAlignment.Left,
                        lineHeight,
                        visuals.Count,
                        segment.Run.LinkUri,
                        segment.Run.Source,
                        segment.Run.Style.SemanticRole);
                    visuals.Add(visual.TranslatePaint(segment.Run.PaintOffsetX, segment.Run.PaintOffsetY, visuals.Count));
                }

                x += segment.Width;
            }

            y += lineHeight;
            breakOffsets.Add(y);
        }

        return new HtmlInlineLayout(visuals, y, breakOffsets);
    }

    private static IReadOnlyList<InlineSegment> MergeAdjacentInlineSegments(IReadOnlyList<InlineSegment> segments) {
        var merged = new List<InlineSegment>(segments.Count);
        foreach (InlineSegment segment in segments) {
            if (merged.Count > 0 && ReferenceEquals(merged[merged.Count - 1].Run, segment.Run)) {
                InlineSegment previous = merged[merged.Count - 1];
                merged[merged.Count - 1] = new InlineSegment(previous.Text + segment.Text, previous.Width + segment.Width, previous.Run);
            } else {
                merged.Add(segment);
            }
        }

        return merged;
    }

    private void AddBrokenToken(ICollection<InlineLine> lines, ref InlineLine line, HtmlInlineRun run, string token, double width) {
        var part = new StringBuilder();
        double partWidth = 0D;
        foreach (string value in OfficeTextElements.Enumerate(token)) {
            double charWidth = MeasureText(value, run.Style.Font);
            if (part.Length > 0 && partWidth + charWidth > width) {
                if (line.Segments.Count > 0) {
                    TrimTrailingWhitespace(line);
                    lines.Add(line);
                    line = new InlineLine();
                }

                line.Add(new InlineSegment(part.ToString(), partWidth, run));
                lines.Add(line);
                line = new InlineLine();
                part.Clear();
                partWidth = 0D;
            }

            part.Append(value);
            partWidth += charWidth;
        }

        if (part.Length > 0) {
            if (line.Segments.Count > 0 && line.Width + partWidth > width) {
                TrimTrailingWhitespace(line);
                lines.Add(line);
                line = new InlineLine();
            }

            line.Add(new InlineSegment(part.ToString(), partWidth, run));
        }
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
            for (int i = 0; i < Segments.Count; i++) height = Math.Max(height, Segments[i].Run.Style.LineHeight);
            return Math.Max(0.01D, height);
        }
    }

    private sealed class InlineSegment {
        internal InlineSegment(string text, double width, HtmlInlineRun run) {
            Text = text;
            Width = width;
            Run = run;
        }

        internal string Text { get; }
        internal double Width { get; }
        internal HtmlInlineRun Run { get; }
    }
}
