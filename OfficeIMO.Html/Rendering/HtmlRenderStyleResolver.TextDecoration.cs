using AngleSharp.Dom;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderStyleResolver {
    private readonly HashSet<IElement> _reportedDecorationGeometry = new();
    private readonly Dictionary<IElement, bool> _decorationComplexText = new();

    private void ResolveTextDecorationGeometry(IElement element, HtmlComputedStyle computed,
        HtmlRenderBoxStyle style, HtmlRenderBoxStyle? parent, bool propagatedUnderline) {
        bool originatesDecoration = style.Display != "contents";
        bool overline = originatesDecoration && HtmlRenderCssValues.SplitWhitespace(computed.GetValue("text-decoration-line"))
            .Contains("overline", StringComparer.OrdinalIgnoreCase);
        if (overline) style.OverlineStyle = ResolveTextDecorationStyle(computed.GetValue("text-decoration-style"));
        bool propagated = parent != null && CanReceiveTextDecoration(style.Display, style.Position, style.FloatSide)
            && (parent.VectorTextDecoration || propagatedUnderline);
        if (propagated && parent != null) {
            if (style.OverlineStyle == OfficeTextDecorationStyle.None) style.OverlineStyle = parent.OverlineStyle;
            if (style.StrikethroughStyle == OfficeTextDecorationStyle.None) style.StrikethroughStyle = parent.StrikethroughStyle;
        }
        bool decorated = style.UnderlineStyle != OfficeTextDecorationStyle.None
            || style.StrikethroughStyle != OfficeTextDecorationStyle.None || style.OverlineStyle != OfficeTextDecorationStyle.None;
        if (!decorated) return;
        string thickness = computed.GetValue("text-decoration-thickness").Trim().ToLowerInvariant();
        string offset = computed.GetValue("text-underline-offset").Trim().ToLowerInvariant();
        if (propagated && parent != null) {
            style.DecorationThickness = parent.DecorationThickness;
            style.UnderlineOffset = parent.UnderlineOffset;
            style.VectorTextDecoration = parent.VectorTextDecoration;
            style.DecorationColor = parent.DecorationColor;
        } else {
            if (thickness.Length > 0 && thickness != "auto" && thickness != "from-font"
                && TryResolveLength(thickness, style.Font.Size, style.Font.Size, _rootFontSize, out double parsed))
                style.DecorationThickness = Math.Max(1D, parsed);
            if (offset.Length > 0 && offset != "auto"
                && TryResolveLength(offset, style.Font.Size, style.Font.Size, _rootFontSize, out double parsedOffset))
                style.UnderlineOffset = parsedOffset;
        }
        style.VectorTextDecoration |= style.OverlineStyle != OfficeTextDecorationStyle.None || style.DecorationThickness.HasValue || style.UnderlineOffset.HasValue;
        string skipInk = computed.GetValue("text-decoration-skip-ink").Trim().ToLowerInvariant();
        string position = computed.GetValue("text-underline-position").Trim().ToLowerInvariant();
        var boundaries = new List<string>();
        if (originatesDecoration && thickness == "from-font") boundaries.Add("from-font thickness has no qualified selected-face decoration metrics");
        if (style.VectorTextDecoration && (style.WritingMode != "horizontal-tb" || style.Direction == "rtl"
            || style.UnicodeBidi != "normal" || style.BaselineLevel != 0
            || ContainsDecorationComplexText(element)
            || element.LocalName is "input" or "textarea" or "select")) {
            boundaries.Add("vertical, bidi and shifted-baseline decoration geometry uses existing automatic paint");
            style.VectorTextDecoration = false;
            style.OverlineStyle = OfficeTextDecorationStyle.None;
        }
        if (style.VectorTextDecoration && (style.UnderlineStyle != OfficeTextDecorationStyle.None
            || style.OverlineStyle != OfficeTextDecorationStyle.None) && (skipInk == "auto" || skipInk == "all"))
            boundaries.Add("glyph ink skipping is not applied to the continuous vector decoration");
        if (style.VectorTextDecoration && style.UnderlineStyle != OfficeTextDecorationStyle.None && position.Length > 0 && position != "auto")
            boundaries.Add("special underline positioning uses the alphabetic baseline");
        if (style.VectorTextDecoration && style.TextShadows.Count > 0)
            boundaries.Add("decoration shadow paint is not projected with explicit vector geometry");
        if (style.VectorTextDecoration && propagated && parent != null && Math.Abs(style.Font.Size - parent.Font.Size) > 0.000001D)
            boundaries.Add("mixed-size descendants use their own em box for propagated decoration placement");
        if (style.VectorTextDecoration && propagated && originatesDecoration
            && HasAuthoredValue(computed, "text-decoration-line")
            && HtmlRenderCssValues.SplitWhitespace(computed.GetValue("text-decoration-line"))
                .Any(line => string.Equals(line, "underline", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(line, "overline", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(line, "line-through", StringComparison.OrdinalIgnoreCase)))
            boundaries.Add("nested independently authored decorations share one resolved decoration band");
        if (style.VectorTextDecoration && propagated && style.Display == "inline"
            && (style.MarginLeft != 0D || style.MarginRight != 0D || style.HorizontalInsets != 0D))
            boundaries.Add("propagated decoration continuity across descendant inline margins, borders and padding is not qualified");
        if (style.VectorTextDecoration && (IsPatterned(style.UnderlineStyle)
            || IsPatterned(style.OverlineStyle) || IsPatterned(style.StrikethroughStyle)))
            boundaries.Add("non-solid pattern phase, spacing and ink use bounded vector approximations");
        if (boundaries.Count > 0 && _reportedDecorationGeometry.Add(element)) {
            string boundary = string.Join("; ", boundaries);
            _diagnostics.Add("OfficeIMO.Html.Renderer", HtmlRenderDiagnosticCodes.TextDecorationThicknessApproximated,
                "Text decoration geometry is partially represented: " + boundary + ".",
                HtmlDiagnosticSeverity.Warning, DescribeSource(element), boundary, OfficeConversionLossKind.Approximation);
        }
        if (overline && !style.VectorTextDecoration)
            _diagnostics.Add("OfficeIMO.Html.Renderer", HtmlRenderDiagnosticCodes.TextDecorationLineUnsupported,
                "Overline is omitted in this specialized text context.", HtmlDiagnosticSeverity.Warning,
                DescribeSource(element), "text-decoration-line=overline", OfficeConversionLossKind.Omission);
    }

    private static bool IsPatterned(OfficeTextDecorationStyle style) =>
        style != OfficeTextDecorationStyle.None && style != OfficeTextDecorationStyle.Single;

    private bool CanReceiveTextDecoration(string display, string position, string floatSide) =>
        display is "inline" or "contents"
        && position.Trim().ToLowerInvariant() is not ("absolute" or "fixed")
        && (string.IsNullOrWhiteSpace(floatSide) || string.Equals(floatSide.Trim(), "none", StringComparison.OrdinalIgnoreCase)
            // Continuous output keeps footnotes in flow rather than extracting them.
            || _options.Mode != HtmlRenderMode.Paged && string.Equals(floatSide.Trim(), "footnote", StringComparison.OrdinalIgnoreCase));

    private bool ContainsDecorationComplexText(IElement root) {
        if (_decorationComplexText.TryGetValue(root, out bool cached)) return cached;
        // Cache a bottom-up flag rather than materializing each ancestor's TextContent.
        // The DOM stays fixed during layout, including continuation and intrinsic passes.
        var pending = new Stack<(IElement Element, bool Exit)>();
        _operationBudget.ChargeLayoutOperations(1L, _options.MaxLayoutOperations, "text decoration subtree scan");
        pending.Push((root, false));
        while (pending.Count > 0) {
            _cancellationToken.ThrowIfCancellationRequested();
            (IElement element, bool exit) = pending.Pop();
            if (exit) {
                foreach (IElement child in element.Children) {
                    _cancellationToken.ThrowIfCancellationRequested();
                    if (_decorationComplexText[child]) {
                        _decorationComplexText[element] = true;
                        break;
                    }
                }
                continue;
            }
            if (_decorationComplexText.ContainsKey(element)) continue;
            bool complex = false;
            foreach (INode node in element.ChildNodes) {
                _cancellationToken.ThrowIfCancellationRequested();
                if (node is not IText text) continue;
                for (int offset = 0; offset < text.Data.Length;) {
                    _cancellationToken.ThrowIfCancellationRequested();
                    _operationBudget.ChargeLayoutOperations(1L, _options.MaxLayoutOperations, "text decoration character scan");
                    int count = Math.Min(256, text.Data.Length - offset);
                    if (offset + count < text.Data.Length && char.IsHighSurrogate(text.Data[offset + count - 1])
                        && char.IsLowSurrogate(text.Data[offset + count])) count++;
                    string part = text.Data.Substring(offset, count);
                    offset += count;
                    if (OfficeTextElements.ContainsRightToLeft(part) || OfficeTextElements.ContainsBidiControl(part)) {
                        complex = true;
                        break;
                    }
                }
                if (complex) break;
            }
            _decorationComplexText[element] = complex;
            // Complete child summaries even if local text is complex, so later child
            // style resolutions reuse the same bounded traversal.
            pending.Push((element, true));
            foreach (IElement child in element.Children) {
                _cancellationToken.ThrowIfCancellationRequested();
                if (_decorationComplexText.ContainsKey(child)) continue;
                _operationBudget.ChargeLayoutOperations(1L, _options.MaxLayoutOperations, "text decoration subtree scan");
                pending.Push((child, false));
            }
        }
        return _decorationComplexText[root];
    }
}
