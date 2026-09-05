using AngleSharp.Dom;
using OfficeIMO.Drawing;
using System.Globalization;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private void AddTextShadowVisuals(
        ICollection<HtmlRenderVisual> rootVisuals,
        IDictionary<IElement, List<HtmlRenderVisual>> ownedVisuals,
        HtmlInlineRun run,
        IElement? formattingContainer,
        IReadOnlyList<HtmlRenderVisual> textVisuals) {
        HtmlRenderBoxStyle style = run.Style;
        string source = run.Source ?? "text";
        if (!ValidateTextShadows(style, run.OwnerElement, source) || textVisuals.Count == 0) return;

        for (int layerIndex = style.TextShadows.Count - 1; layerIndex >= 0; layerIndex--) {
            HtmlCssTextShadow layer = style.TextShadows[layerIndex];
            if (layer.Opacity <= 0D) continue;
            IReadOnlyList<OfficePoint> samples = CreateTextShadowSamples(layer);
            var sampleGroups = new List<HtmlRenderVisual>(samples.Count);
            double left = double.MaxValue;
            double top = double.MaxValue;
            double right = double.MinValue;
            double bottom = double.MinValue;
            string layerSource = TextShadowSource(source, layerIndex, style.TextShadowLayerCount);
            for (int sampleIndex = 0; sampleIndex < samples.Count; sampleIndex++) {
                OfficePoint sample = samples[sampleIndex];
                var painted = new List<HtmlRenderVisual>(textVisuals.Count);
                foreach (HtmlRenderVisual visual in textVisuals) {
                    if (visual is not HtmlRenderText text) continue;
                    double x = text.X + layer.OffsetX + sample.X;
                    double y = text.Y + layer.OffsetY + sample.Y;
                    painted.Add(CloneTextShadow(text, layer.Color, x, y, painted.Count));
                    left = Math.Min(left, x);
                    top = Math.Min(top, y);
                    right = Math.Max(right, x + text.Width);
                    bottom = Math.Max(bottom, y + text.Height);
                }
                if (painted.Count == 0) continue;
                double sampleLeft = painted.Min(visual => visual.X);
                double sampleTop = painted.Min(visual => visual.Y);
                double sampleRight = painted.Max(visual => visual.X + visual.Width);
                double sampleBottom = painted.Max(visual => visual.Y + visual.Height);
                sampleGroups.Add(new HtmlRenderEffectGroup(
                    sampleLeft,
                    sampleTop,
                    Math.Max(0.01D, sampleRight - sampleLeft),
                    Math.Max(0.01D, sampleBottom - sampleTop),
                    OfficeTransform.Identity,
                    ResolveTextShadowSampleOpacity(layer.Opacity, samples.Count, sampleIndex),
                    painted,
                    sampleGroups.Count,
                    layerSource + ":sample[" + sampleIndex.ToString(CultureInfo.InvariantCulture) + "]"));
                ChargeLayoutOperations(painted.Count, layerSource);
            }
            if (sampleGroups.Count == 0) continue;
            var artifact = new HtmlRenderSemanticGroup(
                HtmlRenderSemanticGroupRole.Artifact,
                left,
                top,
                Math.Max(0.01D, right - left),
                Math.Max(0.01D, bottom - top),
                sampleGroups,
                rootVisuals.Count,
                layerSource);
            AddInlineOwnedVisual(rootVisuals, ownedVisuals, artifact, run.OwnerElement, formattingContainer);
        }
    }

    private bool ValidateTextShadows(HtmlRenderBoxStyle style, IElement? element, string source) {
        if (style.UnsupportedTextShadow.Length > 0) {
            if (_reportedTextShadowFallbacks.Add(source)) {
                _diagnostics.Add(
                    ComponentName,
                    HtmlRenderDiagnosticCodes.TextShadowValueUnsupported,
                    "A CSS text shadow was omitted.",
                    HtmlDiagnosticSeverity.Warning,
                    element == null ? source : HtmlRenderStyleResolver.DescribeSource(element),
                    "text-shadow=" + style.UnsupportedTextShadow,
                    OfficeConversionLossKind.Omission);
            }
            return false;
        }

        if (style.TextShadowLayerCount > _options.MaxTextShadowLayers
            && _reportedTextShadowFallbacks.Add(source + ":limit")) {
            _diagnostics.Add(
                ComponentName,
                HtmlRenderDiagnosticCodes.TextShadowLayerLimit,
                "CSS text-shadow layers beyond the configured per-element limit were omitted.",
                HtmlDiagnosticSeverity.Warning,
                element == null ? source : HtmlRenderStyleResolver.DescribeSource(element),
                "layers=" + style.TextShadowLayerCount.ToString(CultureInfo.InvariantCulture)
                    + ";limit=" + _options.MaxTextShadowLayers.ToString(CultureInfo.InvariantCulture),
                OfficeConversionLossKind.Omission);
        }
        return style.TextShadows.Count > 0;
    }

    private static IReadOnlyList<OfficePoint> CreateTextShadowSamples(HtmlCssTextShadow shadow) {
        if (shadow.BlurRadius <= 0.0001D) return new[] { new OfficePoint(0D, 0D) };
        double radius = shadow.BlurRadius * 0.65D;
        double diagonal = radius * 0.7071067811865476D;
        return new[] {
            new OfficePoint(0D, 0D),
            new OfficePoint(radius, 0D),
            new OfficePoint(-radius, 0D),
            new OfficePoint(0D, radius),
            new OfficePoint(0D, -radius),
            new OfficePoint(diagonal, diagonal),
            new OfficePoint(diagonal, -diagonal),
            new OfficePoint(-diagonal, diagonal),
            new OfficePoint(-diagonal, -diagonal)
        };
    }

    private static double ResolveTextShadowSampleOpacity(double opacity, int sampleCount, int sampleIndex) {
        if (sampleCount <= 1) return opacity;
        return Math.Max(0D, Math.Min(1D, opacity * (sampleIndex == 0 ? 0.7D : 0.18D)));
    }

    private static HtmlRenderText CloneTextShadow(
        HtmlRenderText text,
        OfficeColor color,
        double x,
        double y,
        int paintOrder) =>
        new HtmlRenderText(
            text.Text,
            x,
            y,
            text.Width,
            text.Height,
            text.Font,
            color,
            text.Alignment,
            text.LineHeight,
            paintOrder,
            null,
            text.Source,
            null,
            layoutY: null,
            semanticNodeId: null,
            textAdvanceWidth: text.TextAdvanceWidth,
            bidiVisualOrderResolved: text.BidiVisualOrderResolved,
            semanticFragmentOrder: null,
            logicalTextOrder: null,
            underlineStyle: text.UnderlineStyle,
            strikethroughStyle: text.StrikethroughStyle,
            baseline: text.Baseline,
            baselineLevel: text.BaselineLevel,
            baselineScale: text.BaselineScale,
            baselineOffset: text.BaselineOffset,
            textPaintWidth: text.TextPaintWidth,
            decorationColor: color,
            featureSettings: text.FeatureSettings,
            fontPalette: text.FontPalette);

    private static string TextShadowSource(string source, int index, int count) =>
        count == 1
            ? source + ":text-shadow"
            : source + ":text-shadow[" + index.ToString(CultureInfo.InvariantCulture) + "]";
}
