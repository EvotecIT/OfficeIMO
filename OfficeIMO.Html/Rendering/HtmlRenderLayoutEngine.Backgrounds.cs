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
        IElement source) {
        if (!style.PaintVisible || width <= 0.0001D || height <= 0.0001D) return;
        string sourceDescription = HtmlRenderStyleResolver.DescribeSource(source);
        HtmlResolvedBorderRadii radii = ResolveBoxRadii(style, width, height, source, sourceDescription);
        AddOuterBoxShadows(visuals, style, x, y, width, height, radii, source, sourceDescription);
        AddBoxBackgroundCore(visuals, style, x, y, width, height, style.BorderInsets, radii, source, sourceDescription, sourceDescription);
        AddInsetBoxShadows(visuals, style, x, y, width, height, radii, source, sourceDescription);

        AddBorderPaint(visuals, style, x, y, width, height, radii, source, sourceDescription);
    }

    private void AddBoxOutlinePaint(
        ICollection<HtmlRenderVisual> visuals,
        HtmlRenderBoxStyle style,
        double x,
        double y,
        double width,
        double height,
        IElement source) {
        if (!style.PaintVisible || width <= 0.0001D || height <= 0.0001D) return;
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
        HtmlResolvedBorderRadii radii = ResolveBoxRadii(style, width, height, source, diagnosticSourceDescription);
        AddOuterBoxShadows(visuals, style, x, y, width, height, radii, source, diagnosticSourceDescription);
        AddBoxBackgroundCore(visuals, style, x, y, width, height, HtmlRenderBorderInsets.Uniform(borderWidth), radii, source, diagnosticSourceDescription, visualSourceDescription);
        AddInsetBoxShadows(visuals, style, x, y, width, height, radii, source, diagnosticSourceDescription);
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
        IElement source,
        string diagnosticSourceDescription,
        string visualSourceDescription) {
        if (style.BackgroundColor.HasValue && style.BackgroundColor.Value.A > 0) {
            OfficeShape fill = CreateBoxShape(width, height, radii);
            fill.FillColor = style.BackgroundColor;
            fill.StrokeWidth = 0D;
            visuals.Add(new HtmlRenderShape(fill, x, y, visuals.Count, source: visualSourceDescription));
        }

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
        IElement source,
        string diagnosticSourceDescription,
        string visualSourceDescription) {
        if (style.BackgroundImageLayerCount == 0) return;

        if (style.BackgroundImageLayerCount > _options.MaxBackgroundImageLayers) {
            AddUnsupported(
                HtmlRenderDiagnosticCodes.BackgroundImageLayerLimit,
                "CSS background-image layers beyond the configured per-element limit were omitted.",
                source,
                "layers=" + style.BackgroundImageLayerCount.ToString(CultureInfo.InvariantCulture)
                    + ";limit=" + _options.MaxBackgroundImageLayers.ToString(CultureInfo.InvariantCulture));
        }

        if (style.UnsupportedBackgroundImageLayerCount > 0) {
            AddUnsupported(
                HtmlRenderDiagnosticCodes.BackgroundImageValueUnsupported,
                "Unsupported or invalid CSS background image functions were omitted.",
                source,
                "layers=" + style.UnsupportedBackgroundImageLayerCount.ToString(CultureInfo.InvariantCulture));
        }

        if (style.GradientStopLimitExceededCount > 0) {
            _diagnostics.Add(
                ComponentName,
                HtmlRenderDiagnosticCodes.GradientStopLimitExceeded,
                "CSS gradients exceeded the configured color-stop limit and were omitted.",
                HtmlDiagnosticSeverity.Error,
                diagnosticSourceDescription,
                "gradients=" + style.GradientStopLimitExceededCount.ToString(CultureInfo.InvariantCulture)
                    + ";limit=" + _options.MaxGradientStops.ToString(CultureInfo.InvariantCulture));
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
        IElement source,
        string diagnosticSourceDescription,
        string visualSourceDescription) {
        string layerSuffix = style.BackgroundImageLayers.Count > 1
            ? "[" + layerIndex.ToString(CultureInfo.InvariantCulture) + "]"
            : string.Empty;
        string layerVisualSource = visualSourceDescription + ":background-image" + layerSuffix;
        double areaX = x + borderInsets.Left;
        double areaY = y + borderInsets.Top;
        double areaWidth = Math.Max(0.01D, width - borderInsets.Horizontal);
        double areaHeight = Math.Max(0.01D, height - borderInsets.Vertical);
        HtmlResolvedBorderRadii innerRadii = radii.Inset(borderInsets.Left, borderInsets.Top, borderInsets.Right, borderInsets.Bottom, areaWidth, areaHeight);
        if (layer.LinearGradient != null || layer.RadialGradient != null) {
            AddGradientBackground(
                visuals,
                style,
                layer,
                areaX,
                areaY,
                areaWidth,
                areaHeight,
                innerRadii,
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
        OfficeDrawing? svgDrawing = null;
        if (string.Equals(contentType, "image/svg+xml", StringComparison.OrdinalIgnoreCase)
            && !TryReadSvgDrawing(bytes, diagnosticSourceDescription + ":background-image", out svgDrawing)) {
            return;
        }

        double intrinsicWidth = imageInfo != null && imageInfo.Width > 0
            ? imageInfo.Width * HtmlRenderOptions.CssPixelsPerInch / Math.Max(1D, imageInfo.DpiX)
            : areaWidth;
        double intrinsicHeight = imageInfo != null && imageInfo.Height > 0
            ? imageInfo.Height * HtmlRenderOptions.CssPixelsPerInch / Math.Max(1D, imageInfo.DpiY)
            : areaHeight;
        BackgroundImageSize imageSize = ResolveBackgroundImageSize(layer.Size, areaWidth, areaHeight, intrinsicWidth, intrinsicHeight, style.Font.Size, out bool usedSizeFallback);
        if (usedSizeFallback) {
            AddUnsupported(
                HtmlRenderDiagnosticCodes.BackgroundImageValueUnsupported,
                "A CSS background-size value used a contained deterministic fallback.",
                source,
                layer.Size);
        }

        bool repeatSupported = TryResolveBackgroundRepeat(layer.Repeat, out BackgroundRepeatMode repeatX, out BackgroundRepeatMode repeatY);
        if (!repeatSupported) {
            AddUnsupported(
                HtmlRenderDiagnosticCodes.BackgroundImageRepeatUnsupported,
                "A CSS background-repeat value used a single-image fallback.",
                source,
                layer.Repeat);
        }

        if (!repeatSupported) {
            repeatX = BackgroundRepeatMode.NoRepeat;
            repeatY = BackgroundRepeatMode.NoRepeat;
        }

        imageSize = ApplyRoundBackgroundSize(imageSize, layer.Size, areaWidth, areaHeight, repeatX, repeatY);

        (double offsetX, double offsetY) = ResolveBackgroundPosition(
            layer.Position,
            areaWidth - imageSize.Width,
            areaHeight - imageSize.Height,
            style.Font.Size);
        double tileX = areaX + offsetX;
        double tileY = areaY + offsetY;
        BackgroundRepeatAxis horizontal = ResolveBackgroundRepeatAxis(repeatX, areaX, areaWidth, tileX, imageSize.Width);
        BackgroundRepeatAxis vertical = ResolveBackgroundRepeatAxis(repeatY, areaY, areaHeight, tileY, imageSize.Height);
        if (horizontal.Repeat || vertical.Repeat) {
            var pattern = new OfficeImagePatternLayout(
                new OfficeImagePlacement(areaX, areaY, areaWidth, areaHeight),
                new OfficeImagePlacement(horizontal.Origin, vertical.Origin, imageSize.Width, imageSize.Height),
                horizontal.Repeat,
                vertical.Repeat,
                horizontal.Step,
                vertical.Step);
            long tileCount = pattern.EstimatedTileCount;
            if (tileCount > 0L && tileCount <= _options.MaxBackgroundImageTiles - _backgroundImageTileCount) {
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
                _backgroundImageTileCount += tileCount;
            } else if (tileCount > 0L) {
                _diagnostics.Add(
                    ComponentName,
                    HtmlRenderDiagnosticCodes.BackgroundImageTileLimitExceeded,
                    "Repeated CSS background images exceeded the configured operation-wide tile limit and used one clipped origin tile.",
                    HtmlDiagnosticSeverity.Error,
                    diagnosticSourceDescription,
                    "tiles=" + tileCount.ToString(CultureInfo.InvariantCulture) + ";limit=" + _options.MaxBackgroundImageTiles.ToString(CultureInfo.InvariantCulture));
                if (svgDrawing != null) AddVisibleBackgroundDrawing(layerVisuals, svgDrawing, tileX, tileY, imageSize.Width, imageSize.Height, areaX, areaY, areaWidth, areaHeight, layerVisualSource);
                else AddVisibleBackgroundImage(layerVisuals, bytes, contentType, tileX, tileY, imageSize.Width, imageSize.Height, areaX, areaY, areaWidth, areaHeight, layerVisualSource);
            }
        } else {
            if (svgDrawing != null) AddVisibleBackgroundDrawing(layerVisuals, svgDrawing, tileX, tileY, imageSize.Width, imageSize.Height, areaX, areaY, areaWidth, areaHeight, layerVisualSource);
            else AddVisibleBackgroundImage(layerVisuals, bytes, contentType, tileX, tileY, imageSize.Width, imageSize.Height, areaX, areaY, areaWidth, areaHeight, layerVisualSource);
        }
        AddBoxClipVisuals(
            visuals,
            layerVisuals,
            areaX,
            areaY,
            areaWidth,
            areaHeight,
            innerRadii,
            layerVisualSource + ":clip");

    }

    private void AddGradientBackground(
        ICollection<HtmlRenderVisual> visuals,
        HtmlRenderBoxStyle style,
        HtmlRenderBackgroundLayer layer,
        double areaX,
        double areaY,
        double areaWidth,
        double areaHeight,
        HtmlResolvedBorderRadii radii,
        IElement source,
        string visualSourceDescription) {
        if (!string.Equals(layer.Size.Trim(), "auto", StringComparison.OrdinalIgnoreCase)) {
            AddUnsupported(
                HtmlRenderDiagnosticCodes.BackgroundImageValueUnsupported,
                "An explicit CSS background-size on a gradient used the full paint area.",
                source,
                layer.Size);
        }

        OfficeShape fill = CreateBoxShape(areaWidth, areaHeight, radii);
        fill.FillColor = null;
        OfficeLinearGradient? linearGradient = null;
        if (layer.LinearGradient != null
            && !layer.LinearGradient.TryResolve(areaWidth, areaHeight, style.Font.Size, _options.DefaultFontSize, out linearGradient)) {
            AddUnsupported(
                HtmlRenderDiagnosticCodes.BackgroundImageValueUnsupported,
                "A CSS linear gradient could not be resolved against its paint area and was omitted.",
                source);
            return;
        }
        fill.FillGradient = linearGradient;
        OfficeRadialGradient? radialGradient = null;
        if (layer.RadialGradient != null
            && !layer.RadialGradient.TryResolve(areaWidth, areaHeight, style.Font.Size, _options.DefaultFontSize, out radialGradient)) {
            AddUnsupported(
                HtmlRenderDiagnosticCodes.BackgroundImageValueUnsupported,
                "A CSS radial gradient could not be resolved against its paint area and was omitted.",
                source);
            return;
        }

        fill.FillRadialGradient = radialGradient;
        fill.StrokeWidth = 0D;
        visuals.Add(new HtmlRenderShape(fill, areaX, areaY, visuals.Count, source: visualSourceDescription));
    }

    private HtmlResolvedBorderRadii ResolveBoxRadii(
        HtmlRenderBoxStyle style,
        double width,
        double height,
        IElement source,
        string sourceDescription) {
        if (HtmlCssBorderRadiusParser.TryResolve(style, width, height, _options.DefaultFontSize, out HtmlResolvedBorderRadii radii, out string detail)) {
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
            ? OfficeShape.RoundedRectangle(width, height, normalized.UniformRadius)
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

    private static bool TryResolveBackgroundRepeat(string value, out BackgroundRepeatMode repeatX, out BackgroundRepeatMode repeatY) {
        repeatX = BackgroundRepeatMode.Repeat;
        repeatY = BackgroundRepeatMode.Repeat;
        IReadOnlyList<string> values = HtmlRenderCssValues.SplitWhitespace(value ?? string.Empty)
            .Select(token => token.Trim().ToLowerInvariant())
            .Where(token => token.Length > 0)
            .ToList()
            .AsReadOnly();
        if (values.Count == 0) return true;
        if (values.Count > 2) return false;
        if (values.Count == 1 && values[0] == "repeat-x") {
            repeatX = BackgroundRepeatMode.Repeat;
            repeatY = BackgroundRepeatMode.NoRepeat;
            return true;
        }

        if (values.Count == 1 && values[0] == "repeat-y") {
            repeatX = BackgroundRepeatMode.NoRepeat;
            repeatY = BackgroundRepeatMode.Repeat;
            return true;
        }

        if (!TryParseBackgroundRepeatMode(values[0], out repeatX)) return false;
        if (values.Count == 1) {
            repeatY = repeatX;
            return true;
        }

        return TryParseBackgroundRepeatMode(values[1], out repeatY);
    }

    private static bool TryParseBackgroundRepeatMode(string value, out BackgroundRepeatMode mode) {
        switch (value) {
            case "repeat":
                mode = BackgroundRepeatMode.Repeat;
                return true;
            case "no-repeat":
                mode = BackgroundRepeatMode.NoRepeat;
                return true;
            case "space":
                mode = BackgroundRepeatMode.Space;
                return true;
            case "round":
                mode = BackgroundRepeatMode.Round;
                return true;
            default:
                mode = BackgroundRepeatMode.NoRepeat;
                return false;
        }
    }

    private static BackgroundRepeatAxis ResolveBackgroundRepeatAxis(
        BackgroundRepeatMode mode,
        double areaStart,
        double areaLength,
        double positionedTileStart,
        double tileLength) {
        if (mode == BackgroundRepeatMode.NoRepeat) return new BackgroundRepeatAxis(positionedTileStart, tileLength, repeat: false);
        if (mode == BackgroundRepeatMode.Repeat) return new BackgroundRepeatAxis(positionedTileStart, tileLength, repeat: true);
        if (mode == BackgroundRepeatMode.Round) return new BackgroundRepeatAxis(areaStart, tileLength, repeat: true);

        long count = Math.Max(0L, (long)Math.Floor((areaLength + 0.0000001D) / tileLength));
        if (count < 2L) return new BackgroundRepeatAxis(positionedTileStart, tileLength, repeat: false);
        double step = (areaLength - tileLength) / (count - 1L);
        return new BackgroundRepeatAxis(areaStart, Math.Max(tileLength, step), repeat: true);
    }

    private static BackgroundImageSize ApplyRoundBackgroundSize(
        BackgroundImageSize size,
        string sizeValue,
        double areaWidth,
        double areaHeight,
        BackgroundRepeatMode repeatX,
        BackgroundRepeatMode repeatY) {
        double width = size.Width;
        double height = size.Height;
        double originalWidth = width;
        double originalHeight = height;
        ResolveBackgroundSizeAutoAxes(sizeValue, out bool widthAuto, out bool heightAuto);
        if (repeatX == BackgroundRepeatMode.Round) {
            width = ResolveRoundedTileLength(areaWidth, width);
            if (repeatY != BackgroundRepeatMode.Round && heightAuto) height *= width / Math.Max(0.01D, originalWidth);
        }

        if (repeatY == BackgroundRepeatMode.Round) {
            height = ResolveRoundedTileLength(areaHeight, height);
            if (repeatX != BackgroundRepeatMode.Round && widthAuto) width *= height / Math.Max(0.01D, originalHeight);
        }

        return new BackgroundImageSize(Math.Max(0.01D, width), Math.Max(0.01D, height));
    }

    private static double ResolveRoundedTileLength(double areaLength, double tileLength) {
        double ratio = areaLength / Math.Max(0.01D, tileLength);
        long count = Math.Max(1L, (long)Math.Round(ratio, MidpointRounding.AwayFromZero));
        return areaLength / count;
    }

    private static void ResolveBackgroundSizeAutoAxes(string value, out bool widthAuto, out bool heightAuto) {
        IReadOnlyList<string> parts = HtmlRenderCssValues.SplitWhitespace(value ?? string.Empty);
        if (parts.Count == 0 || (parts.Count == 1 && string.Equals(parts[0], "auto", StringComparison.OrdinalIgnoreCase))) {
            widthAuto = true;
            heightAuto = true;
            return;
        }

        if (parts.Count == 1
            && (string.Equals(parts[0], "cover", StringComparison.OrdinalIgnoreCase)
                || string.Equals(parts[0], "contain", StringComparison.OrdinalIgnoreCase))) {
            widthAuto = false;
            heightAuto = false;
            return;
        }

        widthAuto = string.Equals(parts[0], "auto", StringComparison.OrdinalIgnoreCase);
        heightAuto = parts.Count == 1 || string.Equals(parts[1], "auto", StringComparison.OrdinalIgnoreCase);
    }

    private BackgroundImageSize ResolveBackgroundImageSize(
        string value,
        double areaWidth,
        double areaHeight,
        double intrinsicWidth,
        double intrinsicHeight,
        double fontSize,
        out bool usedFallback) {
        usedFallback = false;
        string normalized = (value ?? string.Empty).Trim().ToLowerInvariant();
        if (normalized.Length == 0 || normalized == "auto") {
            return new BackgroundImageSize(intrinsicWidth, intrinsicHeight);
        }

        if (normalized == "cover" || normalized == "contain") {
            double scaleX = areaWidth / Math.Max(0.01D, intrinsicWidth);
            double scaleY = areaHeight / Math.Max(0.01D, intrinsicHeight);
            double scale = normalized == "cover" ? Math.Max(scaleX, scaleY) : Math.Min(scaleX, scaleY);
            return new BackgroundImageSize(intrinsicWidth * scale, intrinsicHeight * scale);
        }

        IReadOnlyList<string> parts = HtmlRenderCssValues.SplitWhitespace(normalized).ToList().AsReadOnly();
        if (parts.Count == 0 || parts.Count > 2) {
            usedFallback = true;
            return Contain(areaWidth, areaHeight, intrinsicWidth, intrinsicHeight);
        }

        bool widthAuto = parts[0] == "auto";
        bool heightAuto = parts.Count == 1 || parts[1] == "auto";
        double resolvedWidth = intrinsicWidth;
        double resolvedHeight = intrinsicHeight;
        if (!widthAuto && !HtmlRenderCssValues.TryLength(parts[0], areaWidth, fontSize, _options.DefaultFontSize, out resolvedWidth)) {
            usedFallback = true;
            return Contain(areaWidth, areaHeight, intrinsicWidth, intrinsicHeight);
        }

        if (!heightAuto && !HtmlRenderCssValues.TryLength(parts[1], areaHeight, fontSize, _options.DefaultFontSize, out resolvedHeight)) {
            usedFallback = true;
            return Contain(areaWidth, areaHeight, intrinsicWidth, intrinsicHeight);
        }

        double ratio = intrinsicWidth / Math.Max(0.01D, intrinsicHeight);
        if (widthAuto && !heightAuto) resolvedWidth = resolvedHeight * ratio;
        if (!widthAuto && heightAuto) resolvedHeight = resolvedWidth / Math.Max(0.01D, ratio);
        return new BackgroundImageSize(Math.Max(0.01D, resolvedWidth), Math.Max(0.01D, resolvedHeight));
    }

    private static BackgroundImageSize Contain(double areaWidth, double areaHeight, double intrinsicWidth, double intrinsicHeight) {
        double scale = Math.Min(areaWidth / Math.Max(0.01D, intrinsicWidth), areaHeight / Math.Max(0.01D, intrinsicHeight));
        return new BackgroundImageSize(intrinsicWidth * scale, intrinsicHeight * scale);
    }

    private (double X, double Y) ResolveBackgroundPosition(string value, double availableX, double availableY, double fontSize) {
        IReadOnlyList<string> parts = HtmlRenderCssValues.SplitWhitespace(value ?? string.Empty).ToList().AsReadOnly();
        string first = parts.Count > 0 ? parts[0].ToLowerInvariant() : "0%";
        string second = parts.Count > 1 ? parts[1].ToLowerInvariant() : "center";
        if (IsVerticalPosition(first) && !IsVerticalPosition(second)) {
            (first, second) = (second, first);
        }

        return (
            ResolveBackgroundAxis(first, availableX, fontSize, horizontal: true),
            ResolveBackgroundAxis(second, availableY, fontSize, horizontal: false));
    }

    private double ResolveBackgroundAxis(string value, double available, double fontSize, bool horizontal) {
        if (value == "center") return available / 2D;
        if (value == (horizontal ? "right" : "bottom")) return available;
        if (value == (horizontal ? "left" : "top")) return 0D;
        if (value.EndsWith("%", StringComparison.Ordinal)
            && double.TryParse(value.Substring(0, value.Length - 1), NumberStyles.Float, CultureInfo.InvariantCulture, out double percentage)) {
            return available * percentage / 100D;
        }

        return HtmlRenderCssValues.TryLength(value, available, fontSize, _options.DefaultFontSize, out double length)
            ? length
            : 0D;
    }

    private static bool IsVerticalPosition(string value) => value == "top" || value == "bottom";

    private readonly struct BackgroundImageSize {
        internal BackgroundImageSize(double width, double height) {
            Width = width;
            Height = height;
        }

        internal double Width { get; }
        internal double Height { get; }
    }

    private readonly struct BackgroundRepeatAxis {
        internal BackgroundRepeatAxis(double origin, double step, bool repeat) {
            Origin = origin;
            Step = step;
            Repeat = repeat;
        }

        internal double Origin { get; }
        internal double Step { get; }
        internal bool Repeat { get; }
    }

    private enum BackgroundRepeatMode {
        NoRepeat,
        Repeat,
        Space,
        Round
    }
}
