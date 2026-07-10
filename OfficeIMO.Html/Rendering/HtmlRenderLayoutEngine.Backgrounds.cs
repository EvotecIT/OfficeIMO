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
        string sourceDescription = HtmlRenderStyleResolver.DescribeSource(source);
        AddBoxBackground(visuals, style, x, y, width, height, style.BorderWidth, source, sourceDescription, sourceDescription);

        if (style.BorderWidth > 0D) {
            OfficeShape border = OfficeShape.Rectangle(width, height);
            border.FillColor = null;
            border.StrokeColor = style.BorderColor;
            border.StrokeWidth = style.BorderWidth;
            visuals.Add(new HtmlRenderShape(border, x, y, visuals.Count, source: sourceDescription));
        }
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
        if (style.BackgroundColor.HasValue && style.BackgroundColor.Value.A > 0) {
            OfficeShape fill = OfficeShape.Rectangle(width, height);
            fill.FillColor = style.BackgroundColor;
            fill.StrokeWidth = 0D;
            visuals.Add(new HtmlRenderShape(fill, x, y, visuals.Count, source: visualSourceDescription));
        }

        AddBackgroundImage(visuals, style, x, y, width, height, borderWidth, source, diagnosticSourceDescription, visualSourceDescription);
    }

    private void AddBackgroundImage(
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
        if (string.IsNullOrWhiteSpace(style.BackgroundImageSource)) {
            return;
        }

        if (style.BackgroundImageLayerCount > 1) {
            AddUnsupported(
                HtmlRenderDiagnosticCodes.BackgroundImageLayerLimit,
                "Only the first CSS background-image layer was painted.",
                source,
                "layers=" + style.BackgroundImageLayerCount.ToString(CultureInfo.InvariantCulture));
        }

        if (!TryResolveImageSource(style.BackgroundImageSource, diagnosticSourceDescription + ":background-image", out byte[]? bytes, out string contentType, out OfficeImageInfo? imageInfo)
            || bytes == null) {
            return;
        }

        double areaX = x + borderWidth;
        double areaY = y + borderWidth;
        double areaWidth = Math.Max(0.01D, width - (borderWidth * 2D));
        double areaHeight = Math.Max(0.01D, height - (borderWidth * 2D));
        double intrinsicWidth = imageInfo != null && imageInfo.Width > 0
            ? imageInfo.Width * HtmlRenderOptions.CssPixelsPerInch / Math.Max(1D, imageInfo.DpiX)
            : areaWidth;
        double intrinsicHeight = imageInfo != null && imageInfo.Height > 0
            ? imageInfo.Height * HtmlRenderOptions.CssPixelsPerInch / Math.Max(1D, imageInfo.DpiY)
            : areaHeight;
        BackgroundImageSize imageSize = ResolveBackgroundImageSize(style.BackgroundSize, areaWidth, areaHeight, intrinsicWidth, intrinsicHeight, style.Font.Size, out bool usedSizeFallback);
        if (usedSizeFallback) {
            AddUnsupported(
                HtmlRenderDiagnosticCodes.BackgroundImageValueUnsupported,
                "A CSS background-size value used a contained deterministic fallback.",
                source,
                style.BackgroundSize);
        }

        bool repeatSupported = TryResolveBackgroundRepeat(style.BackgroundRepeat, out bool repeatX, out bool repeatY);
        if (!repeatSupported) {
            AddUnsupported(
                HtmlRenderDiagnosticCodes.BackgroundImageRepeatUnsupported,
                "A CSS background-repeat value used a single-image fallback.",
                source,
                style.BackgroundRepeat);
        }

        (double offsetX, double offsetY) = ResolveBackgroundPosition(
            style.BackgroundPosition,
            areaWidth - imageSize.Width,
            areaHeight - imageSize.Height,
            style.Font.Size);
        double tileX = areaX + offsetX;
        double tileY = areaY + offsetY;
        if (repeatSupported && (repeatX || repeatY)) {
            var pattern = new OfficeImagePatternLayout(
                new OfficeImagePlacement(areaX, areaY, areaWidth, areaHeight),
                new OfficeImagePlacement(tileX, tileY, imageSize.Width, imageSize.Height),
                repeatX,
                repeatY);
            long tileCount = pattern.EstimatedTileCount;
            if (tileCount > 0L && tileCount <= _options.MaxBackgroundImageTiles - _backgroundImageTileCount) {
                visuals.Add(new HtmlRenderImagePattern(
                    bytes,
                    contentType,
                    pattern,
                    _options.MaxBackgroundImageTiles,
                    visuals.Count,
                    visualSourceDescription + ":background-image"));
                _backgroundImageTileCount += tileCount;
            } else if (tileCount > 0L) {
                _diagnostics.Add(
                    ComponentName,
                    HtmlRenderDiagnosticCodes.BackgroundImageTileLimitExceeded,
                    "Repeated CSS background images exceeded the configured operation-wide tile limit and used one clipped origin tile.",
                    HtmlDiagnosticSeverity.Error,
                    diagnosticSourceDescription,
                    "tiles=" + tileCount.ToString(CultureInfo.InvariantCulture) + ";limit=" + _options.MaxBackgroundImageTiles.ToString(CultureInfo.InvariantCulture));
                AddVisibleBackgroundImage(visuals, bytes, contentType, tileX, tileY, imageSize.Width, imageSize.Height, areaX, areaY, areaWidth, areaHeight, visualSourceDescription);
            }
        } else {
            AddVisibleBackgroundImage(visuals, bytes, contentType, tileX, tileY, imageSize.Width, imageSize.Height, areaX, areaY, areaWidth, areaHeight, visualSourceDescription);
        }

        if (!OfficeRasterImageDecoder.TryDecode(bytes, out _)
            && !string.Equals(contentType, "image/svg+xml", StringComparison.OrdinalIgnoreCase)) {
            _diagnostics.Add(
                ComponentName,
                HtmlRenderDiagnosticCodes.RasterDecoderUnavailable,
                "The background image can be retained for SVG/PDF but the dependency-free PNG backend cannot decode it.",
                HtmlDiagnosticSeverity.Warning,
                diagnosticSourceDescription,
                contentType);
        }
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
        string sourceDescription) {
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
            source: sourceDescription + ":background-image",
            sourceCrop: crop));
    }

    private static bool TryResolveBackgroundRepeat(string value, out bool repeatX, out bool repeatY) {
        repeatX = false;
        repeatY = false;
        IReadOnlyList<string> values = HtmlRenderCssValues.SplitWhitespace(value ?? string.Empty)
            .Select(token => token.Trim().ToLowerInvariant())
            .Where(token => token.Length > 0)
            .ToList()
            .AsReadOnly();
        if (values.Count > 2) return false;
        if (values.Count == 0 || values[0] == "repeat") {
            repeatX = true;
            repeatY = values.Count < 2 || values[1] == "repeat";
            return values.Count < 2 || values[1] == "repeat" || values[1] == "no-repeat";
        }

        if (values[0] == "no-repeat") {
            repeatX = false;
            repeatY = values.Count > 1 && values[1] == "repeat";
            return values.Count < 2 || values[1] == "repeat" || values[1] == "no-repeat";
        }

        if (values.Count == 1 && values[0] == "repeat-x") {
            repeatX = true;
            return true;
        }

        if (values.Count == 1 && values[0] == "repeat-y") {
            repeatY = true;
            return true;
        }

        return false;
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
}
