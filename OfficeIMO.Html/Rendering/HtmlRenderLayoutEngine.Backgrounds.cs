using AngleSharp.Dom;
using OfficeIMO.Drawing;
using System.Globalization;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private void AddBoxPaint(
        ICollection<HtmlRenderVisual> visuals,
        HtmlRenderBoxStyle style,
        double x,
        double y,
        double width,
        double height,
        IElement source,
        bool paintBorders = true) {
        // Record the finalized border box even when it has no background or border paint.
        // Descendant geometry then follows the normal overflow/effect/fragmentation path.
        if (CapturePrintLayoutBoxes && width > 0.0001D && height > 0.0001D) {
            visuals.Add(new HtmlRenderLayoutBox(x, y, width, height, visuals.Count,
                HtmlRenderStyleResolver.DescribeSource(source)));
        }
        if (!style.PaintVisible) return;
        int backgroundStart = visuals.Count;
        string sourceDescription = HtmlRenderStyleResolver.DescribeSource(source);
        HtmlResolvedBorderRadii radii = ResolveBoxRadii(style, width, height, source, sourceDescription);
        AddOuterBoxShadows(visuals, style, x, y, width, height, radii, source, sourceDescription);
        AddBoxBackgroundCore(visuals, style, x, y, width, height, style.BorderInsets, radii, sourceDescription, sourceDescription, sourceDescription);
        AddInsetBoxShadows(visuals, style, x, y, width, height, style.BorderInsets, radii, source, sourceDescription);

        if (paintBorders || style.UnsupportedBorderPaint.Length > 0) {
            AddBorderPaint(visuals, style, x, y, width, height, radii, source, sourceDescription);
        }
        foreach (HtmlRenderVisual visual in visuals.Skip(backgroundStart)) visual.PaintPhase = HtmlRenderPaintPhase.BlockBackground;
    }

    private void AddBoxOutlinePaint(
        ICollection<HtmlRenderVisual> visuals,
        HtmlRenderBoxStyle style,
        double x,
        double y,
        double width,
        double height,
        IElement source) {
        if (!style.PaintVisible) return;
        string sourceDescription = HtmlRenderStyleResolver.DescribeSource(source);
        HtmlResolvedBorderRadii radii = ResolveBoxRadii(style, width, height, source, sourceDescription);
        AddOutlinePaint(visuals, style, x, y, width, height, radii, source, sourceDescription);
    }

    private void AddBoxBackground(
        ICollection<HtmlRenderVisual> visuals,
        HtmlRenderBoxStyle style,
        double x,
        double y,
        double width,
        double height,
        double borderWidth,
        IElement source,
        string diagnosticSourceDescription,
        string visualSourceDescription) {
        if (!style.PaintVisible) return;
        int backgroundStart = visuals.Count;
        HtmlResolvedBorderRadii radii = ResolveBoxRadii(style, width, height, source, diagnosticSourceDescription);
        AddOuterBoxShadows(visuals, style, x, y, width, height, radii, source, diagnosticSourceDescription);
        AddBoxBackgroundCore(visuals, style, x, y, width, height, HtmlRenderBorderInsets.Uniform(borderWidth), radii, HtmlRenderStyleResolver.DescribeSource(source), diagnosticSourceDescription, visualSourceDescription);
        AddInsetBoxShadows(visuals, style, x, y, width, height, HtmlRenderBorderInsets.Uniform(borderWidth), radii, source, diagnosticSourceDescription);
        foreach (HtmlRenderVisual visual in visuals.Skip(backgroundStart)) visual.PaintPhase = HtmlRenderPaintPhase.BlockBackground;
    }

    private void AddBoxBackgroundCore(
        ICollection<HtmlRenderVisual> visuals,
        HtmlRenderBoxStyle style,
        double x,
        double y,
        double width,
        double height,
        HtmlRenderBorderInsets borderInsets,
        HtmlResolvedBorderRadii radii,
        string source,
        string diagnosticSourceDescription,
        string visualSourceDescription) {
        // Outer shadows and outlines can expand a zero-layout box. Its own
        // background still has no paint area, for ordinary and generated boxes.
        if (width <= 0.0001D || height <= 0.0001D) return;
        AddSolidBoxBackground(visuals, style, x, y, width, height, borderInsets, radii, visualSourceDescription);

        AddBackgroundImages(visuals, style, x, y, width, height, borderInsets, radii, source, diagnosticSourceDescription, visualSourceDescription);
    }

    private void AddBackgroundImages(
        ICollection<HtmlRenderVisual> visuals,
        HtmlRenderBoxStyle style,
        double x,
        double y,
        double width,
        double height,
        HtmlRenderBorderInsets borderInsets,
        HtmlResolvedBorderRadii radii,
        string source,
        string diagnosticSourceDescription,
        string visualSourceDescription) {
        if (style.BackgroundImageLayerCount == 0) return;

        if (style.BackgroundImageLayerCount > _options.MaxBackgroundImageLayers) {
            AddUnsupported(
                HtmlRenderDiagnosticCodes.BackgroundImageLayerLimit,
                "CSS background-image layers beyond the configured per-element limit were omitted.",
                source,
                "layers=" + style.BackgroundImageLayerCount.ToString(CultureInfo.InvariantCulture)
                    + ";limit=" + _options.MaxBackgroundImageLayers.ToString(CultureInfo.InvariantCulture),
                OfficeConversionLossKind.Omission);
        }

        if (style.UnsupportedBackgroundImageLayerCount > 0) {
            AddUnsupported(
                HtmlRenderDiagnosticCodes.BackgroundImageValueUnsupported,
                "Unsupported or invalid CSS background image functions were omitted.",
                source,
                "layers=" + style.UnsupportedBackgroundImageLayerCount.ToString(CultureInfo.InvariantCulture),
                OfficeConversionLossKind.Omission);
        }

        if (style.GradientStopLimitExceededCount > 0) {
            _diagnostics.Add(
                ComponentName,
                HtmlRenderDiagnosticCodes.GradientStopLimitExceeded,
                "CSS gradients exceeded the configured color-stop limit and were omitted.",
                HtmlDiagnosticSeverity.Error,
                diagnosticSourceDescription,
                "gradients=" + style.GradientStopLimitExceededCount.ToString(CultureInfo.InvariantCulture)
                    + ";limit=" + _options.MaxGradientStops.ToString(CultureInfo.InvariantCulture),
                OfficeConversionLossKind.Omission);
        }

        for (int layerIndex = style.BackgroundImageLayers.Count - 1; layerIndex >= 0; layerIndex--) {
            AddBackgroundImageLayer(
                visuals,
                style,
                style.BackgroundImageLayers[layerIndex],
                layerIndex,
                x,
                y,
                width,
                height,
                borderInsets,
                radii,
                source,
                diagnosticSourceDescription,
                visualSourceDescription);
        }
    }

    private void AddBackgroundImageLayer(
        ICollection<HtmlRenderVisual> visuals,
        HtmlRenderBoxStyle style,
        HtmlRenderBackgroundLayer layer,
        int layerIndex,
        double x,
        double y,
        double width,
        double height,
        HtmlRenderBorderInsets borderInsets,
        HtmlResolvedBorderRadii radii,
        string source,
        string diagnosticSourceDescription,
        string visualSourceDescription) {
        string layerSuffix = style.BackgroundImageLayers.Count > 1
            ? "[" + layerIndex.ToString(CultureInfo.InvariantCulture) + "]"
            : string.Empty;
        string layerVisualSource = visualSourceDescription + ":background-image" + layerSuffix;
        BackgroundPaintBox originBox = ResolveBackgroundBox(layer.Origin, x, y, width, height, borderInsets, style, radii);
        BackgroundPaintBox clipBox = ResolveBackgroundBox(layer.Clip, x, y, width, height, borderInsets, style, radii);
        if (clipBox.Width <= 0D || clipBox.Height <= 0D) return;
        double areaX = originBox.X;
        double areaY = originBox.Y;
        double areaWidth = originBox.Width;
        double areaHeight = originBox.Height;
        if (layer.Attachment == "fixed") {
            AddUnsupported(
                HtmlRenderDiagnosticCodes.BackgroundImageValueUnsupported,
                "CSS background-attachment:fixed used the element-local static-page positioning area.",
                source,
                "background-attachment=fixed",
                OfficeConversionLossKind.Approximation);
        }
        if (layer.LinearGradient != null || layer.RadialGradient != null || layer.ConicGradient != null) {
            AddGradientBackground(
                visuals,
                style,
                layer,
                areaX,
                areaY,
                areaWidth,
                areaHeight,
                originBox.Radii,
                clipBox,
                source,
                visualSourceDescription + ":background-gradient" + layerSuffix);
            return;
        }

        if (string.IsNullOrWhiteSpace(layer.Source)
            || !TryResolveImageSource(layer.Source!, diagnosticSourceDescription + ":background-image", out byte[]? bytes, out string contentType, out OfficeImageInfo? imageInfo)
            || bytes == null) {
            return;
        }
        var layerVisuals = new List<HtmlRenderVisual>();
        double intrinsicWidth = imageInfo != null && imageInfo.Width > 0
            ? imageInfo.Width * HtmlRenderOptions.CssPixelsPerInch / Math.Max(1D, imageInfo.DpiX)
            : areaWidth;
        double intrinsicHeight = imageInfo != null && imageInfo.Height > 0
            ? imageInfo.Height * HtmlRenderOptions.CssPixelsPerInch / Math.Max(1D, imageInfo.DpiY)
            : areaHeight;
        BackgroundImageSize imageSize = ResolveBackgroundImageSize(layer.Size, areaWidth, areaHeight, intrinsicWidth, intrinsicHeight, style, out bool usedSizeFallback);
        if (usedSizeFallback) {
            AddUnsupported(
                HtmlRenderDiagnosticCodes.BackgroundImageValueUnsupported,
                "A CSS background-size value used a contained deterministic fallback.",
                source,
                layer.Size);
        }

        ResolveBackgroundRepeatModes(layer.Repeat, source, out BackgroundRepeatMode repeatX, out BackgroundRepeatMode repeatY);
        if (imageSize.Width <= 0D || imageSize.Height <= 0D) return;
        imageSize = ApplyRoundBackgroundSize(imageSize, layer.Size, areaWidth, areaHeight, repeatX, repeatY);
        if (imageSize.Width <= 0D || imageSize.Height <= 0D) return;

        // A viewBox fits to the concrete tile, including repeat:round sizing.
        // Without one, retain the intrinsic scene and existing tile projection.
        OfficeDrawing? svgDrawing = null;
        if (string.Equals(contentType, "image/svg+xml", StringComparison.OrdinalIgnoreCase)
            && !TryReadSvgDrawing(
                bytes,
                intrinsicWidth,
                intrinsicHeight,
                diagnosticSourceDescription + ":background-image",
                out svgDrawing,
                viewportWidth: imageInfo?.HasSvgViewBox == true ? imageSize.Width : null,
                viewportHeight: imageInfo?.HasSvgViewBox == true ? imageSize.Height : null)) {
            return;
        }

        OfficeImagePatternLayout pattern = ResolveBackgroundPattern(layer, style, originBox, clipBox, imageSize, repeatX, repeatY);
        double tileX = pattern.Tile.X;
        double tileY = pattern.Tile.Y;
        if (pattern.RepeatX || pattern.RepeatY) {
            long tileCount = pattern.EstimatedTileCount;
            if (TryReserveBackgroundTiles(tileCount, diagnosticSourceDescription)) {
                if (svgDrawing != null) {
                    AddBackgroundDrawingPattern(layerVisuals, svgDrawing, pattern, _options.MaxBackgroundImageTiles, layerVisualSource);
                } else {
                    layerVisuals.Add(new HtmlRenderImagePattern(
                        bytes,
                        contentType,
                        pattern,
                        _options.MaxBackgroundImageTiles,
                        layerVisuals.Count,
                        layerVisualSource));
                }
            } else if (tileCount > 0L) {
                if (svgDrawing != null) AddVisibleBackgroundDrawing(layerVisuals, svgDrawing, tileX, tileY, imageSize.Width, imageSize.Height, clipBox.X, clipBox.Y, clipBox.Width, clipBox.Height, layerVisualSource);
                else AddVisibleBackgroundImage(layerVisuals, bytes, contentType, tileX, tileY, imageSize.Width, imageSize.Height, clipBox.X, clipBox.Y, clipBox.Width, clipBox.Height, layerVisualSource);
            }
        } else {
            if (svgDrawing != null) AddVisibleBackgroundDrawing(layerVisuals, svgDrawing, tileX, tileY, imageSize.Width, imageSize.Height, clipBox.X, clipBox.Y, clipBox.Width, clipBox.Height, layerVisualSource);
            else AddVisibleBackgroundImage(layerVisuals, bytes, contentType, tileX, tileY, imageSize.Width, imageSize.Height, clipBox.X, clipBox.Y, clipBox.Width, clipBox.Height, layerVisualSource);
        }
        AddBoxClipVisuals(
            visuals,
            layerVisuals,
            clipBox.X,
            clipBox.Y,
            clipBox.Width,
            clipBox.Height,
            clipBox.Radii,
            layerVisualSource + ":clip");

    }

    private static BackgroundPaintBox ResolveBackgroundBox(
        string box,
        double x,
        double y,
        double width,
        double height,
        HtmlRenderBorderInsets borderInsets,
        HtmlRenderBoxStyle style,
        HtmlResolvedBorderRadii radii) {
        double top = 0D;
        double right = 0D;
        double bottom = 0D;
        double left = 0D;
        if (box == "padding-box" || box == "content-box") {
            top += borderInsets.Top;
            right += borderInsets.Right;
            bottom += borderInsets.Bottom;
            left += borderInsets.Left;
        }
        if (box == "content-box") {
            top += style.PaddingTop;
            right += style.PaddingRight;
            bottom += style.PaddingBottom;
            left += style.PaddingLeft;
        }

        double boxWidth = Math.Max(0D, width - left - right);
        double boxHeight = Math.Max(0D, height - top - bottom);
        return new BackgroundPaintBox(
            x + left,
            y + top,
            boxWidth,
            boxHeight,
            radii.Inset(left, top, right, bottom, boxWidth, boxHeight));
    }

    private HtmlResolvedBorderRadii ResolveBoxRadii(
        HtmlRenderBoxStyle style,
        double width,
        double height,
        IElement source,
        string sourceDescription) {
        if (HtmlCssBorderRadiusParser.TryResolve(
            style,
            width,
            height,
            _styleResolver.RootFontSize,
            ActiveSurfaceWidth,
            _options.Mode == HtmlRenderMode.Paged ? _activePageGeometry.Height : _options.ViewportHeight ?? 1056D,
            style.ContainerUnitWidth ?? double.NaN,
            style.ContainerUnitHeight ?? double.NaN,
            out HtmlResolvedBorderRadii radii,
            out string detail)) {
            return radii;
        }
        if (_reportedBorderRadiusFallbacks.Add(sourceDescription)) {
            _diagnostics.Add(
                ComponentName,
                HtmlRenderDiagnosticCodes.BorderRadiusValueUnsupported,
                "A CSS border radius used square-corner fallback.",
                HtmlDiagnosticSeverity.Warning,
                HtmlRenderStyleResolver.DescribeSource(source),
                detail);
        }
        return default;
    }

    private static OfficeShape CreateBoxShape(double width, double height, HtmlResolvedBorderRadii radii) {
        HtmlResolvedBorderRadii normalized = radii.Normalize(width, height);
        if (normalized.IsZero) return OfficeShape.Rectangle(width, height);
        return normalized.IsUniformCircular
            ? OfficeShape.RoundedRectangle(width, height, normalized.BoundedUniformRadius(width, height))
            : OfficeShape.Path(normalized.CreatePathCommands(width, height));
    }

    private static void AddVisibleBackgroundImage(
        ICollection<HtmlRenderVisual> visuals,
        byte[] bytes,
        string contentType,
        double tileX,
        double tileY,
        double tileWidth,
        double tileHeight,
        double areaX,
        double areaY,
        double areaWidth,
        double areaHeight,
        string visualSourceDescription) {
        double visibleLeft = Math.Max(tileX, areaX);
        double visibleTop = Math.Max(tileY, areaY);
        double visibleRight = Math.Min(tileX + tileWidth, areaX + areaWidth);
        double visibleBottom = Math.Min(tileY + tileHeight, areaY + areaHeight);
        if (visibleRight <= visibleLeft || visibleBottom <= visibleTop) return;

        OfficeImageSourceCrop crop = OfficeImageSourceCrop.FromStrictFractions(
            Math.Max(0D, (visibleLeft - tileX) / tileWidth),
            Math.Max(0D, (visibleTop - tileY) / tileHeight),
            Math.Max(0D, (tileX + tileWidth - visibleRight) / tileWidth),
            Math.Max(0D, (tileY + tileHeight - visibleBottom) / tileHeight));
        visuals.Add(new HtmlRenderImage(
            bytes,
            contentType,
            visibleLeft,
            visibleTop,
            visibleRight - visibleLeft,
            visibleBottom - visibleTop,
            visuals.Count,
            source: visualSourceDescription,
            sourceCrop: crop));
    }

}
