using AngleSharp.Dom;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderStyleResolver {
    private readonly HashSet<IElement> _reportedDecorationGeometry = new();

    private void ResolveTextDecorationGeometry(IElement element, HtmlComputedStyle computed,
        HtmlRenderBoxStyle style, HtmlRenderBoxStyle? parent, bool propagatedUnderline) {
        bool overline = HtmlRenderCssValues.SplitWhitespace(computed.GetValue("text-decoration-line"))
            .Contains("overline", StringComparer.OrdinalIgnoreCase);
        if (overline) style.OverlineStyle = ResolveTextDecorationStyle(computed.GetValue("text-decoration-style"));
        bool propagated = parent != null && (style.Display == "inline" || style.Display == "contents")
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
        if (thickness == "from-font") boundaries.Add("from-font thickness has no qualified selected-face decoration metrics");
        if (style.VectorTextDecoration && (style.WritingMode != "horizontal-tb" || style.Direction == "rtl"
            || style.UnicodeBidi != "normal" || style.BaselineLevel != 0
            || OfficeTextElements.ContainsRightToLeft(element.TextContent)
            || OfficeTextElements.ContainsBidiControl(element.TextContent)
            || element.LocalName is "input" or "textarea" or "select")) {
            boundaries.Add("vertical, bidi and shifted-baseline decoration geometry uses existing automatic paint");
            style.VectorTextDecoration = false;
            style.OverlineStyle = OfficeTextDecorationStyle.None;
        }
        if (style.VectorTextDecoration && (skipInk == "auto" || skipInk == "all"))
            boundaries.Add("glyph ink skipping is not applied to the continuous vector decoration");
        if (style.VectorTextDecoration && position.Length > 0 && position != "auto")
            boundaries.Add("special underline positioning uses the alphabetic baseline");
        if (style.VectorTextDecoration && style.TextShadows.Count > 0)
            boundaries.Add("decoration shadow paint is not projected with explicit vector geometry");
        if (style.VectorTextDecoration && propagated && parent != null && Math.Abs(style.Font.Size - parent.Font.Size) > 0.000001D)
            boundaries.Add("mixed-size descendants use their own em box for propagated decoration placement");
        if (style.VectorTextDecoration && propagated && HasAuthoredValue(computed, "text-decoration-line"))
            boundaries.Add("nested independently authored decorations share one resolved decoration band");
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
}
