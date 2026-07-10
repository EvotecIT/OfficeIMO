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
        if (style.BackgroundColor.HasValue && style.BackgroundColor.Value.A > 0) {
            OfficeShape fill = OfficeShape.Rectangle(width, height);
            fill.FillColor = style.BackgroundColor;
            fill.StrokeWidth = 0D;
            visuals.Add(new HtmlRenderShape(fill, x, y, visuals.Count, source: sourceDescription));
        }

        AddBackgroundImage(visuals, style, x, y, width, height, source, sourceDescription);

        if (style.BorderWidth > 0D) {
            OfficeShape border = OfficeShape.Rectangle(width, height);
            border.FillColor = null;
            border.StrokeColor = style.BorderColor;
            border.StrokeWidth = style.BorderWidth;
            visuals.Add(new HtmlRenderShape(border, x, y, visuals.Count, source: sourceDescription));
        }
    }

    private void AddBackgroundImage(
        ICollection<HtmlRenderVisual> visuals,
        HtmlRenderBoxStyle style,
        double x,
        double y,
        double width,
        double height,
        IElement source,
        string sourceDescription) {
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

        if (!TryResolveImageSource(style.BackgroundImageSource, sourceDescription + ":background-image", out byte[]? bytes, out string contentType, out OfficeImageInfo? imageInfo)
            || bytes == null) {
            return;
        }

        double areaX = x + style.BorderWidth;
        double areaY = y + style.BorderWidth;
        double areaWidth = Math.Max(0.01D, width - (style.BorderWidth * 2D));
        double areaHeight = Math.Max(0.01D, height - (style.BorderWidth * 2D));
        double intrinsicWidth = imageInfo != null && imageInfo.Width > 0
            ? imageInfo.Width * HtmlRenderOptions.CssPixelsPerInch / Math.Max(1D, imageInfo.DpiX)
            : areaWidth;
        double intrinsicHeight = imageInfo != null && imageInfo.Height > 0
            ? imageInfo.Height * HtmlRenderOptions.CssPixelsPerInch / Math.Max(1D, imageInfo.DpiY)
            : areaHeight;
        BackgroundImageSize imageSize = ResolveBackgroundImageSize(style.BackgroundSize, areaWidth, areaHeight, intrinsicWidth, intrinsicHeight, style.Font.Size, out bool usedSizeFallback);
        if (imageSize.Width > areaWidth + 0.0001D || imageSize.Height > areaHeight + 0.0001D) {
            double fit = Math.Min(areaWidth / imageSize.Width, areaHeight / imageSize.Height);
            imageSize = new BackgroundImageSize(imageSize.Width * fit, imageSize.Height * fit);
            usedSizeFallback = true;
        }

        if (usedSizeFallback) {
            AddUnsupported(
                HtmlRenderDiagnosticCodes.BackgroundImageValueUnsupported,
                "A CSS background-size value used a contained deterministic fallback.",
                source,
                style.BackgroundSize);
        }

        if (!string.Equals(style.BackgroundRepeat, "no-repeat", StringComparison.OrdinalIgnoreCase)) {
            AddUnsupported(
                HtmlRenderDiagnosticCodes.BackgroundImageRepeatUnsupported,
                "A repeating CSS background image used a single-image fallback until clipped pattern paint is available.",
                source,
                style.BackgroundRepeat);
        }

        (double offsetX, double offsetY) = ResolveBackgroundPosition(
            style.BackgroundPosition,
            areaWidth - imageSize.Width,
            areaHeight - imageSize.Height,
            style.Font.Size);
        visuals.Add(new HtmlRenderImage(
            bytes,
            contentType,
            areaX + offsetX,
            areaY + offsetY,
            imageSize.Width,
            imageSize.Height,
            visuals.Count,
            source: sourceDescription + ":background-image"));
        if (!OfficeRasterImageDecoder.TryDecode(bytes, out _)
            && !string.Equals(contentType, "image/svg+xml", StringComparison.OrdinalIgnoreCase)) {
            _diagnostics.Add(
                ComponentName,
                HtmlRenderDiagnosticCodes.RasterDecoderUnavailable,
                "The background image can be retained for SVG/PDF but the dependency-free PNG backend cannot decode it.",
                HtmlDiagnosticSeverity.Warning,
                sourceDescription,
                contentType);
        }
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
            ResolveBackgroundAxis(first, Math.Max(0D, availableX), fontSize, horizontal: true),
            ResolveBackgroundAxis(second, Math.Max(0D, availableY), fontSize, horizontal: false));
    }

    private double ResolveBackgroundAxis(string value, double available, double fontSize, bool horizontal) {
        if (value == "center") return available / 2D;
        if (value == (horizontal ? "right" : "bottom")) return available;
        if (value == (horizontal ? "left" : "top")) return 0D;
        if (value.EndsWith("%", StringComparison.Ordinal)
            && double.TryParse(value.Substring(0, value.Length - 1), NumberStyles.Float, CultureInfo.InvariantCulture, out double percentage)) {
            return Math.Max(0D, Math.Min(available, available * percentage / 100D));
        }

        return HtmlRenderCssValues.TryLength(value, available, fontSize, _options.DefaultFontSize, out double length)
            ? Math.Max(0D, Math.Min(available, length))
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
