using AngleSharp.Dom;
using OfficeIMO.Drawing;
using System.Globalization;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private void AddOuterBoxShadows(
        ICollection<HtmlRenderVisual> visuals,
        HtmlRenderBoxStyle style,
        double x,
        double y,
        double width,
        double height,
        HtmlResolvedBorderRadii radii,
        IElement source,
        string sourceDescription) {
        if (!ValidateBoxShadows(style, source, sourceDescription)) return;

        for (int index = style.BoxShadows.Count - 1; index >= 0; index--) {
            HtmlCssBoxShadow layer = style.BoxShadows[index];
            OfficeShadow shadow = layer.Shadow;
            if (layer.Inset || shadow.Opacity <= 0D) continue;

            double shadowWidth = width + layer.SpreadRadius * 2D;
            double shadowHeight = height + layer.SpreadRadius * 2D;
            if (shadowWidth <= 0.0001D || shadowHeight <= 0.0001D) continue;
            HtmlResolvedBorderRadii shadowRadii = radii.Expand(layer.SpreadRadius, shadowWidth, shadowHeight);
            OfficeShape carrier = CreateBoxShape(shadowWidth, shadowHeight, shadowRadii);
            carrier.FillColor = null;
            carrier.StrokeColor = null;
            carrier.StrokeWidth = 0D;
            carrier.Shadow = shadow.Clone();
            visuals.Add(new HtmlRenderShape(
                carrier,
                x - layer.SpreadRadius,
                y - layer.SpreadRadius,
                visuals.Count,
                source: BoxShadowSource(sourceDescription, index, style.BoxShadowLayerCount, inset: false)));
        }
    }

    private void AddInsetBoxShadows(
        ICollection<HtmlRenderVisual> visuals,
        HtmlRenderBoxStyle style,
        double x,
        double y,
        double width,
        double height,
        HtmlResolvedBorderRadii radii,
        IElement source,
        string sourceDescription) {
        if (!ValidateBoxShadows(style, source, sourceDescription)) return;

        for (int index = style.BoxShadows.Count - 1; index >= 0; index--) {
            HtmlCssBoxShadow layer = style.BoxShadows[index];
            if (!layer.Inset || layer.Shadow.Opacity <= 0D) continue;
            IReadOnlyList<HtmlRenderVisual> layers = CreateInsetShadowLayers(
                layer,
                x,
                y,
                width,
                height,
                radii,
                BoxShadowSource(sourceDescription, index, style.BoxShadowLayerCount, inset: true));
            if (layers.Count == 0) continue;

            HtmlResolvedBorderRadii normalized = radii.Normalize(width, height);
            if (normalized.IsZero) {
                visuals.Add(new HtmlRenderClipGroup(
                    x,
                    y,
                    width,
                    height,
                    clipHorizontal: true,
                    clipVertical: true,
                    layers,
                    visuals.Count,
                    source: BoxShadowSource(sourceDescription, index, style.BoxShadowLayerCount, inset: true)));
            } else {
                visuals.Add(new HtmlRenderPathClipGroup(
                    x,
                    y,
                    CreateBoxClipPath(width, height, normalized),
                    layers,
                    visuals.Count,
                    source: BoxShadowSource(sourceDescription, index, style.BoxShadowLayerCount, inset: true)));
            }
        }
    }

    private bool ValidateBoxShadows(HtmlRenderBoxStyle style, IElement source, string sourceDescription) {
        if (style.UnsupportedBoxShadow.Length > 0) {
            if (_reportedBoxShadowFallbacks.Add(sourceDescription)) {
                _diagnostics.Add(
                    ComponentName,
                    HtmlRenderDiagnosticCodes.BoxShadowValueUnsupported,
                    "A CSS box shadow was omitted.",
                    HtmlDiagnosticSeverity.Warning,
                    HtmlRenderStyleResolver.DescribeSource(source),
                    "box-shadow=" + style.UnsupportedBoxShadow,
                    HtmlConversionLossKind.Omission);
            }
            return false;
        }

        if (style.BoxShadowLayerCount > _options.MaxBoxShadowLayers
            && _reportedBoxShadowFallbacks.Add(sourceDescription + ":limit")) {
            _diagnostics.Add(
                ComponentName,
                HtmlRenderDiagnosticCodes.BoxShadowLayerLimit,
                "CSS box-shadow layers beyond the configured per-element limit were omitted.",
                HtmlDiagnosticSeverity.Warning,
                HtmlRenderStyleResolver.DescribeSource(source),
                "layers=" + style.BoxShadowLayerCount.ToString(CultureInfo.InvariantCulture)
                    + ";limit=" + _options.MaxBoxShadowLayers.ToString(CultureInfo.InvariantCulture),
                HtmlConversionLossKind.Omission);
        }
        return style.BoxShadows.Count > 0;
    }

    private static IReadOnlyList<HtmlRenderVisual> CreateInsetShadowLayers(
        HtmlCssBoxShadow layer,
        double x,
        double y,
        double width,
        double height,
        HtmlResolvedBorderRadii radii,
        string source) {
        var visuals = new List<HtmlRenderVisual>();
        OfficeShadow shadow = layer.Shadow;
        if (shadow.BlurRadius > 0D) {
            const int blurLayers = 4;
            for (int index = blurLayers; index >= 1; index--) {
                double factor = index / (double)blurLayers;
                double opacity = shadow.Opacity * (0.04D + (blurLayers - index + 1) * 0.05D);
                AddInsetShadowRing(
                    visuals,
                    x,
                    y,
                    width,
                    height,
                    radii,
                    shadow,
                    layer.SpreadRadius + shadow.BlurRadius * factor,
                    opacity,
                    source + ":blur");
            }
        }
        AddInsetShadowRing(visuals, x, y, width, height, radii, shadow, layer.SpreadRadius, shadow.Opacity, source);
        return visuals;
    }

    private static void AddInsetShadowRing(
        ICollection<HtmlRenderVisual> visuals,
        double x,
        double y,
        double width,
        double height,
        HtmlResolvedBorderRadii radii,
        OfficeShadow shadow,
        double inset,
        double opacity,
        string source) {
        if (opacity <= 0D) return;
        double innerX = inset + shadow.OffsetX;
        double innerY = inset + shadow.OffsetY;
        double innerWidth = width - inset * 2D;
        double innerHeight = height - inset * 2D;

        var commands = new List<OfficePathCommand>();
        AppendRoundedContour(commands, 0D, 0D, width, height, radii.Normalize(width, height));
        if (innerWidth > 0.0001D && innerHeight > 0.0001D) {
            HtmlResolvedBorderRadii innerRadii = radii.Inset(inset, inset, inset, inset, innerWidth, innerHeight);
            AppendRoundedContour(commands, innerX, innerY, innerWidth, innerHeight, innerRadii);
        }

        double minX = Math.Min(0D, innerX);
        double minY = Math.Min(0D, innerY);
        OfficeShape ring = OfficeShape.Path(commands);
        ring.FillColor = shadow.Color;
        ring.FillOpacity = opacity;
        ring.StrokeWidth = 0D;
        ring.FillRule = OfficeFillRule.EvenOdd;
        visuals.Add(new HtmlRenderShape(ring, x + minX, y + minY, visuals.Count, source: source));
    }

    private static void AppendRoundedContour(
        ICollection<OfficePathCommand> commands,
        double x,
        double y,
        double width,
        double height,
        HtmlResolvedBorderRadii radii) {
        foreach (OfficePathCommand command in radii.CreatePathCommands(width, height)) {
            switch (command.Kind) {
                case OfficePathCommandKind.MoveTo:
                    commands.Add(OfficePathCommand.MoveTo(x + command.Point.X, y + command.Point.Y));
                    break;
                case OfficePathCommandKind.LineTo:
                    commands.Add(OfficePathCommand.LineTo(x + command.Point.X, y + command.Point.Y));
                    break;
                case OfficePathCommandKind.CubicBezierTo:
                    commands.Add(OfficePathCommand.CubicBezierTo(
                        x + command.ControlPoint1.X,
                        y + command.ControlPoint1.Y,
                        x + command.ControlPoint2.X,
                        y + command.ControlPoint2.Y,
                        x + command.Point.X,
                        y + command.Point.Y));
                    break;
                case OfficePathCommandKind.Close:
                    commands.Add(OfficePathCommand.Close());
                    break;
            }
        }
    }

    private static string BoxShadowSource(string source, int index, int count, bool inset) {
        if (count == 1 && !inset) return source + ":box-shadow";
        return source + ":box-shadow[" + index.ToString(CultureInfo.InvariantCulture) + "]" + (inset ? ":inset" : string.Empty);
    }
}
