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
        double availableWidth = Math.Max(1D, containingWidth - style.MarginLeft - style.MarginRight);
        double boxWidth = ResolveBoxWidth(availableWidth, style);
        double boxHeight = Math.Max(0.01D, block.Height - style.MarginTop - style.MarginBottom);
        HtmlCssResolvedClipPath? clipPath = null;
        bool hasClipPath = style.ClipPath != "none";
        if (hasClipPath && !HtmlCssClipPathParser.TryResolve(
                style.ClipPath,
                boxWidth,
                boxHeight,
                style.Font.Size,
                _options.DefaultFontSize,
                _options.Mode == HtmlRenderMode.Paged ? _activePageGeometry.Width : _options.ViewportWidth,
                _options.Mode == HtmlRenderMode.Paged ? _activePageGeometry.Height : _options.ViewportHeight ?? 1056D,
                style.ContainerUnitWidth ?? double.NaN,
                style.ContainerUnitHeight ?? double.NaN,
                style,
                out clipPath,
                out string clipDetail)) {
            _diagnostics.Add(
                ComponentName,
                HtmlRenderDiagnosticCodes.ClipPathValueUnsupported,
                "A CSS clip-path value used no clipping.",
                HtmlDiagnosticSeverity.Warning,
                source,
                clipDetail,
                OfficeConversionLossKind.Omission);
            hasClipPath = false;
        }
        OfficeTransform transform = OfficeTransform.Identity;
        if (hasTransform) {
            if (!HtmlCssTransformParser.TryParse(
                    style.Transform,
                    style.TransformOrigin,
                    style.MarginLeft,
                    style.MarginTop,
                    boxWidth,
                    boxHeight,
                    style.Font.Size,
                    _options.DefaultFontSize,
                    _options.Mode == HtmlRenderMode.Paged ? _activePageGeometry.Width : _options.ViewportWidth,
                    _options.Mode == HtmlRenderMode.Paged ? _activePageGeometry.Height : _options.ViewportHeight ?? 1056D,
                    style.ContainerUnitWidth ?? double.NaN,
                    style.ContainerUnitHeight ?? double.NaN,
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
        createsStackingContext = hasTransform || hasOpacity || hasClipPath;
        if (!createsStackingContext || block.Visuals.Count == 0) return block;
        IReadOnlyList<HtmlRenderVisual> effectVisuals = hasTransform || hasOpacity || hasClipPath
            ? ReplaceDescendantFormFieldsForPaintEffect(
                block.Visuals,
                hasTransform ? "ancestor-transform=" + source
                    : hasOpacity ? "ancestor-opacity=" + style.Opacity.ToString(System.Globalization.CultureInfo.InvariantCulture)
                    : "ancestor-clip-path=" + style.ClipPath)
            : block.Visuals;
        if (clipPath != null) {
            effectVisuals = new[] {
                new HtmlRenderPathClipGroup(
                    style.MarginLeft + clipPath.X,
                    style.MarginTop + clipPath.Y,
                    clipPath.ClipPath,
                    effectVisuals,
                    0,
                    source)
            };
        }
        if (!hasTransform && !hasOpacity) return block.WithVisuals(effectVisuals);
        var group = new HtmlRenderEffectGroup(
            0D,
            0D,
            Math.Max(0.01D, block.Width),
            Math.Max(0.01D, block.Height),
            transform,
            hasOpacity ? style.Opacity : 1D,
            effectVisuals,
            0,
            source);
        return block.WithVisuals(new[] { group });
    }

    private IReadOnlyList<HtmlRenderVisual> ReplaceDescendantFormFieldsForPaintEffect(
        IReadOnlyList<HtmlRenderVisual> visuals,
        string detail) {
        var replaced = new List<HtmlRenderVisual>(visuals.Count);
        foreach (HtmlRenderVisual visual in visuals) {
            if (visual is HtmlRenderFormField field) {
                ReportTransformedFormFieldFallback(field.Source ?? "form-control", detail);
                replaced.AddRange(field.Visuals);
                continue;
            }

            IReadOnlyList<HtmlRenderVisual>? children = GetGroupChildren(visual);
            if (children == null) {
                replaced.Add(visual);
                continue;
            }

            IReadOnlyList<HtmlRenderVisual> transformedChildren = ReplaceDescendantFormFieldsForPaintEffect(children, detail);
            replaced.Add(CloneGroupWithChildren(visual, transformedChildren));
        }
        return replaced;
    }

    private IReadOnlyList<HtmlRenderVisual> ApplyInlineElementPaintEffects(
        IElement element,
        HtmlRenderBoxStyle style,
        InlineContainingRect? bounds,
        IReadOnlyList<HtmlRenderVisual> visuals) {
        if (bounds == null || visuals.Count == 0) return visuals;
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

        HtmlCssResolvedClipPath? clipPath = null;
        bool hasClipPath = style.ClipPath != "none";
        if (hasClipPath && !HtmlCssClipPathParser.TryResolve(
                style.ClipPath,
                bounds.Width,
                bounds.Height,
                style.Font.Size,
                _options.DefaultFontSize,
                _options.Mode == HtmlRenderMode.Paged ? _activePageGeometry.Width : _options.ViewportWidth,
                _options.Mode == HtmlRenderMode.Paged ? _activePageGeometry.Height : _options.ViewportHeight ?? 1056D,
                style.ContainerUnitWidth ?? double.NaN,
                style.ContainerUnitHeight ?? double.NaN,
                style,
                out clipPath,
                out string clipDetail)) {
            _diagnostics.Add(
                ComponentName,
                HtmlRenderDiagnosticCodes.ClipPathValueUnsupported,
                "A CSS clip-path value used no clipping.",
                HtmlDiagnosticSeverity.Warning,
                source,
                clipDetail,
                OfficeConversionLossKind.Omission);
            hasClipPath = false;
        }

        OfficeTransform transform = OfficeTransform.Identity;
        bool hasTransform = style.Transform != "none";
        if (hasTransform && !HtmlCssTransformParser.TryParse(
                style.Transform,
                style.TransformOrigin,
                bounds.X,
                bounds.Y,
                bounds.Width,
                bounds.Height,
                style.Font.Size,
                _options.DefaultFontSize,
                _options.Mode == HtmlRenderMode.Paged ? _activePageGeometry.Width : _options.ViewportWidth,
                _options.Mode == HtmlRenderMode.Paged ? _activePageGeometry.Height : _options.ViewportHeight ?? 1056D,
                style.ContainerUnitWidth ?? double.NaN,
                style.ContainerUnitHeight ?? double.NaN,
                out transform,
                out string transformDetail)) {
            _diagnostics.Add(
                ComponentName,
                HtmlRenderDiagnosticCodes.TransformValueUnsupported,
                "A CSS transform or transform-origin value used the identity fallback.",
                HtmlDiagnosticSeverity.Warning,
                source,
                transformDetail);
            hasTransform = false;
        }

        bool hasOpacity = style.OpacityWasSpecified && style.UnsupportedOpacity.Length == 0 && style.Opacity < 1D;
        if (!hasTransform && !hasOpacity && !hasClipPath) return visuals;
        IReadOnlyList<HtmlRenderVisual> effectVisuals = ReplaceDescendantFormFieldsForPaintEffect(
            visuals,
            hasTransform ? "ancestor-transform=" + source
                : hasOpacity ? "ancestor-opacity=" + style.Opacity.ToString(System.Globalization.CultureInfo.InvariantCulture)
                : "ancestor-clip-path=" + style.ClipPath);
        if (clipPath != null) {
            effectVisuals = new[] {
                new HtmlRenderPathClipGroup(
                    bounds.X + clipPath.X,
                    bounds.Y + clipPath.Y,
                    clipPath.ClipPath,
                    effectVisuals,
                    0,
                    source)
            };
        }
        if (!hasTransform && !hasOpacity) return effectVisuals;
        return new[] {
            new HtmlRenderEffectGroup(
                bounds.X,
                bounds.Y,
                bounds.Width,
                bounds.Height,
                transform,
                hasOpacity ? style.Opacity : 1D,
                effectVisuals,
                0,
                source)
        };
    }
}
