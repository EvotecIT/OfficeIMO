using AngleSharp.Dom;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private HtmlRenderFlowBlock ApplyElementPositioning(
        HtmlRenderFlowBlock block,
        HtmlRenderBoxStyle style,
        double containingWidth,
        double? containingHeight,
        IElement element) {
        return ApplyPositioning(block, style, containingWidth, containingHeight, HtmlRenderStyleResolver.DescribeSource(element));
    }

    private HtmlRenderFlowBlock ApplyPositioning(
        HtmlRenderFlowBlock block,
        HtmlRenderBoxStyle style,
        double containingWidth,
        double? containingHeight,
        string source) {
        ResolvePositionPaintOffset(style, containingWidth, containingHeight, source, out double offsetX, out double offsetY);
        return Math.Abs(offsetX) <= 0.0001D && Math.Abs(offsetY) <= 0.0001D
            ? block
            : block.TranslatePaint(offsetX, offsetY);
    }

    private void ResolvePositionPaintOffset(
        HtmlRenderBoxStyle style,
        double containingWidth,
        double? containingHeight,
        string source,
        out double offsetX,
        out double offsetY) {
        offsetX = 0D;
        offsetY = 0D;
        if (style.Position == "static") return;
        if (style.Position != "relative") {
            _diagnostics.Add(
                ComponentName,
                HtmlRenderDiagnosticCodes.PositioningModeUnsupported,
                "CSS " + style.Position + " positioning is not yet active; the element used normal flow.",
                HtmlDiagnosticSeverity.Warning,
                source,
                "position=" + style.Position);
            return;
        }

        offsetX = ResolvePositionAxis(style.Left, style.Right, containingWidth, style, source, "left", "right");
        offsetY = ResolvePositionAxis(style.Top, style.Bottom, containingHeight, style, source, "top", "bottom");
        if (style.ZIndex != "auto") {
            _diagnostics.Add(
                ComponentName,
                HtmlRenderDiagnosticCodes.PositionZIndexPending,
                "CSS z-index is not yet active; the relatively positioned element retained source paint order.",
                HtmlDiagnosticSeverity.Warning,
                source,
                "z-index=" + style.ZIndex);
        }
    }

    private double ResolvePositionAxis(
        string leadingValue,
        string trailingValue,
        double? reference,
        HtmlRenderBoxStyle style,
        string source,
        string leadingProperty,
        string trailingProperty) {
        if (!IsAutoInset(leadingValue)) {
            return TryResolvePositionInset(leadingValue, reference, style, source, leadingProperty, out double leading)
                ? leading
                : 0D;
        }

        if (!IsAutoInset(trailingValue)) {
            return TryResolvePositionInset(trailingValue, reference, style, source, trailingProperty, out double trailing)
                ? -trailing
                : 0D;
        }

        return 0D;
    }

    private bool TryResolvePositionInset(
        string value,
        double? reference,
        HtmlRenderBoxStyle style,
        string source,
        string property,
        out double resolved) {
        resolved = 0D;
        bool percentage = value.EndsWith("%", StringComparison.Ordinal);
        if (percentage && !reference.HasValue) {
            ReportUnsupportedPositionInset(source, property, value, "the containing block has automatic height");
            return false;
        }

        double lengthReference = reference ?? 0D;
        if (HtmlRenderCssValues.TryLength(value, lengthReference, style.Font.Size, _options.DefaultFontSize, out resolved)) {
            return true;
        }

        ReportUnsupportedPositionInset(source, property, value, "the inset length is outside the active CSS length model");
        return false;
    }

    private void ReportUnsupportedPositionInset(string source, string property, string value, string reason) {
        _diagnostics.Add(
            ComponentName,
            HtmlRenderDiagnosticCodes.PositionInsetUnsupported,
            "A relative-position inset used zero offset because " + reason + ".",
            HtmlDiagnosticSeverity.Warning,
            source,
            property + "=" + value);
    }

    private static bool IsAutoInset(string value) =>
        string.IsNullOrWhiteSpace(value) || string.Equals(value, "auto", StringComparison.OrdinalIgnoreCase);

    private static double? ResolveContainingBlockHeight(HtmlRenderBoxStyle style) {
        if (!style.ExplicitHeight.HasValue) return null;
        return style.BorderBox
            ? Math.Max(0D, style.ExplicitHeight.Value - style.VerticalInsets)
            : style.ExplicitHeight.Value;
    }
}
