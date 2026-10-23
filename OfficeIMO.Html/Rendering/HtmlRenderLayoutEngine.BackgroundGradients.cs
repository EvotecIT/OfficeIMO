using OfficeIMO.Drawing;
using System.Globalization;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    /// <summary>Paints a generated image with the same concrete geometry and tile budget as a URL image.</summary>
    private void AddGradientBackground(
        ICollection<HtmlRenderVisual> visuals, HtmlRenderBoxStyle style, HtmlRenderBackgroundLayer layer,
        double areaX, double areaY, double areaWidth, double areaHeight,
        HtmlResolvedBorderRadii radii, BackgroundPaintBox clipBox, string source, string visualSourceDescription) {
        CheckCancellation();
        BackgroundImageSize size = ResolveBackgroundImageSize(layer.Size, areaWidth, areaHeight,
            areaWidth, areaHeight, style, out bool usedFallback, hasIntrinsicRatio: false);
        if (usedFallback) {
            AddUnsupported(HtmlRenderDiagnosticCodes.BackgroundImageValueUnsupported,
                "A CSS background-size value used a contained deterministic fallback.", source, layer.Size);
        }
        // A zero image axis paints nothing and must not consume the operation tile budget.
        if (size.Width <= 0D || size.Height <= 0D) return;
        ResolveBackgroundRepeatModes(layer.Repeat, source, out BackgroundRepeatMode repeatX, out BackgroundRepeatMode repeatY);
        size = ApplyRoundBackgroundSize(size, layer.Size, areaWidth, areaHeight, repeatX, repeatY, hasIntrinsicRatio: false);
        var origin = new BackgroundPaintBox(areaX, areaY, areaWidth, areaHeight, radii);
        OfficeImagePatternLayout pattern = ResolveBackgroundPattern(layer, style, origin, clipBox, size, repeatX, repeatY);
        bool repeated = pattern.RepeatX || pattern.RepeatY;
        if (repeated && pattern.EstimatedTileCount == 0L) return;
        if (!TryCreateGradientBackground(style, layer, size.Width, size.Height, source,
                out OfficeDrawing? drawing, out OfficeShape? shape) || drawing == null) return;
        bool reserved = !repeated || TryReserveBackgroundTiles(pattern.EstimatedTileCount, source);
        var layerVisuals = new List<HtmlRenderVisual>();
        bool originTileCoversClip = pattern.Tile.X <= clipBox.X && pattern.Tile.Y <= clipBox.Y
            && pattern.Tile.X + size.Width >= clipBox.X + clipBox.Width
            && pattern.Tile.Y + size.Height >= clipBox.Y + clipBox.Height;
        if (repeated && reserved && (pattern.EstimatedTileCount > 1L || !originTileCoversClip)) {
            AddBackgroundDrawingPattern(layerVisuals, drawing, pattern, _options.MaxBackgroundImageTiles, visualSourceDescription);
        } else if (shape != null) {
            // Preserve the direct shape route for a single gradient, including PDF alpha handling.
            layerVisuals.Add(new HtmlRenderShape(shape, pattern.Tile.X, pattern.Tile.Y, 0, source: visualSourceDescription));
        } else {
            AddVisibleBackgroundDrawing(layerVisuals, drawing, pattern.Tile.X, pattern.Tile.Y,
                size.Width, size.Height, clipBox.X, clipBox.Y, clipBox.Width, clipBox.Height, visualSourceDescription);
        }
        AddBoxClipVisuals(visuals, layerVisuals, clipBox.X, clipBox.Y, clipBox.Width, clipBox.Height,
            clipBox.Radii, visualSourceDescription + ":clip");
    }

    /// <summary>Resolves gradient stops and endpoints in the final tile, independently of the outer box clip.</summary>
    private bool TryCreateGradientBackground(
        HtmlRenderBoxStyle style, HtmlRenderBackgroundLayer layer, double areaWidth, double areaHeight,
        string source, out OfficeDrawing? drawing, out OfficeShape? shape) {
        drawing = null;
        shape = null;
        if (layer.ConicGradient != null) {
            if (!layer.ConicGradient.TryResolve(areaWidth, areaHeight, style.Font.Size, _styleResolver.RootFontSize, ActiveSurfaceWidth, _options.Mode == HtmlRenderMode.Paged ? _activePageGeometry.Height : _options.ViewportHeight ?? 1056D, style.ContainerUnitWidth ?? double.NaN, style.ContainerUnitHeight ?? double.NaN, out OfficeConicGradient? conicGradient, out bool conicStopLimitExceeded, style.CharacterAdvance)
                || conicGradient == null) {
                if (conicStopLimitExceeded) {
                    _diagnostics.Add(ComponentName, HtmlRenderDiagnosticCodes.GradientStopLimitExceeded,
                        "A repeating CSS conic gradient exceeded the configured materialized color-stop limit and was omitted.",
                        HtmlDiagnosticSeverity.Error, source,
                        "limit=" + _options.MaxGradientStops.ToString(CultureInfo.InvariantCulture), OfficeConversionLossKind.Omission);
                } else {
                    AddUnsupported(HtmlRenderDiagnosticCodes.BackgroundImageValueUnsupported,
                        "A CSS conic gradient could not be resolved against its paint area and was omitted.",
                        source, lossKind: OfficeConversionLossKind.Omission);
                }
                return false;
            }
            drawing = conicGradient.CreateDrawing(areaWidth, areaHeight, _options.ConicGradientQualitySegments);
            return true;
        }

        OfficeShape fill = OfficeShape.Rectangle(areaWidth, areaHeight);
        fill.FillColor = null;
        bool stopLimitExceeded = false;
        OfficeLinearGradient? linearGradient = null;
        if (layer.LinearGradient != null
            && !layer.LinearGradient.TryResolve(areaWidth, areaHeight, style.Font.Size, _styleResolver.RootFontSize, ActiveSurfaceWidth, _options.Mode == HtmlRenderMode.Paged ? _activePageGeometry.Height : _options.ViewportHeight ?? 1056D, style.ContainerUnitWidth ?? double.NaN, style.ContainerUnitHeight ?? double.NaN, out linearGradient, out stopLimitExceeded, style.CharacterAdvance)) {
            if (stopLimitExceeded) {
                _diagnostics.Add(ComponentName, HtmlRenderDiagnosticCodes.GradientStopLimitExceeded,
                    "A repeating CSS linear gradient exceeded the configured materialized color-stop limit and was omitted.",
                    HtmlDiagnosticSeverity.Error, source,
                    "limit=" + _options.MaxGradientStops.ToString(CultureInfo.InvariantCulture), OfficeConversionLossKind.Omission);
                return false;
            }
            AddUnsupported(
                HtmlRenderDiagnosticCodes.BackgroundImageValueUnsupported,
                "A CSS linear gradient could not be resolved against its paint area and was omitted.",
                source,
                lossKind: OfficeConversionLossKind.Omission);
            return false;
        }
        fill.FillGradient = linearGradient;
        OfficeRadialGradient? radialGradient = null;
        if (layer.RadialGradient != null
            && !layer.RadialGradient.TryResolve(areaWidth, areaHeight, style.Font.Size, _styleResolver.RootFontSize, ActiveSurfaceWidth, _options.Mode == HtmlRenderMode.Paged ? _activePageGeometry.Height : _options.ViewportHeight ?? 1056D, style.ContainerUnitWidth ?? double.NaN, style.ContainerUnitHeight ?? double.NaN, out radialGradient, out stopLimitExceeded, style.CharacterAdvance)) {
            if (stopLimitExceeded) {
                _diagnostics.Add(ComponentName, HtmlRenderDiagnosticCodes.GradientStopLimitExceeded,
                    "A repeating CSS radial gradient exceeded the configured materialized color-stop limit and was omitted.",
                    HtmlDiagnosticSeverity.Error, source,
                    "limit=" + _options.MaxGradientStops.ToString(CultureInfo.InvariantCulture), OfficeConversionLossKind.Omission);
                return false;
            }
            AddUnsupported(
                HtmlRenderDiagnosticCodes.BackgroundImageValueUnsupported,
                "A CSS radial gradient could not be resolved against its paint area and was omitted.",
                source,
                lossKind: OfficeConversionLossKind.Omission);
            return false;
        }

        fill.FillRadialGradient = radialGradient;
        fill.StrokeWidth = 0D;
        shape = fill;
        drawing = new OfficeDrawing(areaWidth, areaHeight);
        drawing.AddShape(fill, 0D, 0D);
        return true;
    }
}
