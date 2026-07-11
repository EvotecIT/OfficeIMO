using AngleSharp.Dom;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private HtmlRenderFlowBlock ApplyElementPaintEffects(
        HtmlRenderFlowBlock block,
        HtmlRenderBoxStyle style,
        double containingWidth,
        IElement element,
        out bool createsStackingContext) {
        createsStackingContext = false;
        string source = HtmlRenderStyleResolver.DescribeSource(element);
        if (style.UnsupportedOpacity.Length > 0) {
            _diagnostics.Add(
                ComponentName,
                HtmlRenderDiagnosticCodes.OpacityValueUnsupported,
                "A CSS opacity value used the opaque fallback.",
                HtmlDiagnosticSeverity.Warning,
                source,
                "opacity=" + style.UnsupportedOpacity);
        }

        bool hasTransform = style.Transform != "none";
        OfficeTransform transform = OfficeTransform.Identity;
        if (hasTransform) {
            double availableWidth = Math.Max(1D, containingWidth - style.MarginLeft - style.MarginRight);
            double boxWidth = ResolveBoxWidth(availableWidth, style);
            double boxHeight = Math.Max(0.01D, block.Height - style.MarginTop - style.MarginBottom);
            if (!HtmlCssTransformParser.TryParse(
                    style.Transform,
                    style.TransformOrigin,
                    style.MarginLeft,
                    style.MarginTop,
                    boxWidth,
                    boxHeight,
                    style.Font.Size,
                    _options.DefaultFontSize,
                    out transform,
                    out string detail)) {
                _diagnostics.Add(
                    ComponentName,
                    HtmlRenderDiagnosticCodes.TransformValueUnsupported,
                    "A CSS transform or transform-origin value used the identity fallback.",
                    HtmlDiagnosticSeverity.Warning,
                    source,
                    detail);
                hasTransform = false;
                transform = OfficeTransform.Identity;
            }
        }

        bool hasOpacity = style.OpacityWasSpecified && style.UnsupportedOpacity.Length == 0 && style.Opacity < 1D;
        createsStackingContext = hasTransform || hasOpacity;
        if (!createsStackingContext || block.Visuals.Count == 0) return block;
        var group = new HtmlRenderEffectGroup(
            0D,
            0D,
            Math.Max(0.01D, block.Width),
            Math.Max(0.01D, block.Height),
            transform,
            hasOpacity ? style.Opacity : 1D,
            block.Visuals,
            0,
            source);
        return block.WithVisuals(new[] { group });
    }

    private void ReportUnsupportedInlinePaintEffects(IElement element, HtmlRenderBoxStyle style) {
        bool transform = style.Transform != "none";
        bool opacity = style.OpacityWasSpecified && (style.Opacity < 1D || style.UnsupportedOpacity.Length > 0);
        if (!transform && !opacity) return;
        var details = new List<string>(2);
        if (transform) details.Add("transform=" + style.Transform);
        if (opacity) details.Add("opacity=" + (style.UnsupportedOpacity.Length > 0 ? style.UnsupportedOpacity : style.Opacity.ToString(System.Globalization.CultureInfo.InvariantCulture)));
        _diagnostics.Add(
            ComponentName,
            HtmlRenderDiagnosticCodes.InlinePaintEffectUnsupported,
            "A non-atomic inline paint effect used normal inline paint.",
            HtmlDiagnosticSeverity.Warning,
            HtmlRenderStyleResolver.DescribeSource(element),
            string.Join(";", details));
    }
}
