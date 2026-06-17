using System.Diagnostics;
using OfficeIMO.PowerPoint;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Markup.PowerPoint;

public sealed partial class OfficeMarkupPowerPointExporter {
    private static void ApplyBackground(PowerPointSlide slide, string? background, PowerPointDesignTheme theme, OfficeMarkupPowerPointExportOptions options, SlideCanvasMetrics metrics) {
        var spec = ParseBackground(background, theme, options);
        if (!string.IsNullOrWhiteSpace(spec.GradientStartColor) && !string.IsNullOrWhiteSpace(spec.GradientEndColor)) {
            slide.SetBackgroundGradient(
                spec.GradientStartColor!,
                spec.GradientEndColor!,
                spec.GradientAngleDegrees ?? 135d);
        } else if (!string.IsNullOrWhiteSpace(spec.Color)) {
            slide.BackgroundColor = spec.Color;
        }

        if (!string.IsNullOrWhiteSpace(spec.ImagePath)) {
            slide.SetBackgroundImage(spec.ImagePath!);
        }

        if (!string.IsNullOrWhiteSpace(spec.OverlayColor)) {
            var overlay = slide.AddShapeInches(
                A.ShapeTypeValues.Rectangle,
                0,
                0,
                metrics.Width,
                metrics.Height,
                "OfficeIMO Markup Background Overlay");
            overlay.FillColor = spec.OverlayColor;
            overlay.FillTransparency = spec.OverlayTransparency;
            overlay.OutlineColor = spec.OverlayColor;
            overlay.OutlineWidthPoints = 0;
        }
    }

    private static string? ParseBackgroundColor(string? background, PowerPointDesignTheme? theme) {
        var spec = ParseBackground(background, theme, null);
        return spec.Color ?? spec.GradientStartColor;
    }

    private static OfficeMarkupBackgroundSpec ParseBackground(string? background, PowerPointDesignTheme? theme, OfficeMarkupPowerPointExportOptions? options) {
        if (string.IsNullOrWhiteSpace(background)) {
            return new OfficeMarkupBackgroundSpec();
        }

        var value = background!.Trim();
        var solid = TryExtractFunctionArgument(value, "solid");
        var gradient = TryExtractFunctionArgument(value, "gradient");
        var image = TryExtractFunctionArgument(value, "image");
        var angle = TryExtractNamedValue(value, "angle");
        var overlay = TryExtractNamedValue(value, "overlay");

        string? resolvedImage = null;
        if (!string.IsNullOrWhiteSpace(image)) {
            var candidate = image!.Trim().Trim('"', '\'');
            if (TryResolveFilePath(options, candidate, out var resolved) && File.Exists(resolved)) {
                resolvedImage = resolved;
            }
        }

        TryParseGradient(gradient, theme, out var gradientStartColor, out var gradientEndColor);
        var spec = new OfficeMarkupBackgroundSpec {
            Color = !string.IsNullOrWhiteSpace(solid) ? ResolveThemeColor(solid, theme) : ResolveThemeColor(value, theme),
            GradientStartColor = gradientStartColor,
            GradientEndColor = gradientEndColor,
            GradientAngleDegrees = TryParseGradientAngle(angle, out var gradientAngleDegrees) ? gradientAngleDegrees : null,
            ImagePath = resolvedImage
        };

        if (TryParseOverlay(overlay, out var overlayColor, out var overlayTransparency)) {
            spec.OverlayColor = overlayColor;
            spec.OverlayTransparency = overlayTransparency;
        }

        return spec;
    }

    private static string? TryExtractFunctionArgument(string value, string functionName) {
        var prefix = functionName + "(";
        var start = value.IndexOf(prefix, StringComparison.OrdinalIgnoreCase);
        if (start < 0) {
            return null;
        }

        start += prefix.Length;
        var end = value.IndexOf(')', start);
        if (end < 0) {
            return null;
        }

        return value.Substring(start, end - start).Trim();
    }

    private static string? TryExtractNamedValue(string value, string attributeName) {
        var start = value.IndexOf(attributeName + "=", StringComparison.OrdinalIgnoreCase);
        if (start < 0) {
            return null;
        }

        start += attributeName.Length + 1;
        if (start >= value.Length) {
            return null;
        }

        var remaining = value.Substring(start);
        if (remaining.StartsWith("rgba(", StringComparison.OrdinalIgnoreCase)) {
            var end = value.IndexOf(')', start);
            return end >= start ? value.Substring(start, end - start + 1).Trim() : null;
        }

        var nextSpace = value.IndexOf(' ', start);
        return (nextSpace >= 0 ? value.Substring(start, nextSpace - start) : value.Substring(start)).Trim();
    }

    private static bool TryParseOverlay(string? value, out string color, out int transparency) {
        color = string.Empty;
        transparency = 0;
        if (string.IsNullOrWhiteSpace(value)) {
            return false;
        }

        var normalized = value!.Trim();
        if (normalized.StartsWith("rgba(", StringComparison.OrdinalIgnoreCase) && normalized.EndsWith(")", StringComparison.Ordinal)) {
            var parts = normalized.Substring(5, normalized.Length - 6).Split(',');
            if (parts.Length == 4
                && int.TryParse(parts[0].Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out var red)
                && int.TryParse(parts[1].Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out var green)
                && int.TryParse(parts[2].Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out var blue)
                && double.TryParse(parts[3].Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out var alpha)) {
                red = Math.Max(0, Math.Min(255, red));
                green = Math.Max(0, Math.Min(255, green));
                blue = Math.Max(0, Math.Min(255, blue));
                alpha = Math.Max(0, Math.Min(1, alpha));
                color = $"{red:X2}{green:X2}{blue:X2}";
                transparency = (int)Math.Round((1 - alpha) * 100);
                return true;
            }
        }

        var hex = ToPowerPointColor(normalized);
        if (!string.IsNullOrWhiteSpace(hex)) {
            color = hex!;
            transparency = 0;
            return true;
        }

        return false;
    }

    private static bool TryParseGradient(
        string? value,
        PowerPointDesignTheme? theme,
        out string? startColor,
        out string? endColor) {
        startColor = null;
        endColor = null;
        if (string.IsNullOrWhiteSpace(value)) {
            return false;
        }

        var parts = value!.Split(',')
            .Select(part => ResolveThemeColor(part.Trim(), theme))
            .Where(color => !string.IsNullOrWhiteSpace(color))
            .Cast<string>()
            .ToList();
        if (parts.Count < 2) {
            return false;
        }

        startColor = parts[0];
        endColor = parts[1];
        return true;
    }

    private static bool TryParseGradientAngle(string? value, out double angleDegrees) {
        angleDegrees = 0;
        if (string.IsNullOrWhiteSpace(value)) {
            return false;
        }

        var normalized = value!.Trim();
        if (normalized.EndsWith("deg", StringComparison.OrdinalIgnoreCase)) {
            normalized = normalized.Substring(0, normalized.Length - 3).Trim();
        }

        return double.TryParse(normalized, NumberStyles.Float, CultureInfo.InvariantCulture, out angleDegrees);
    }

    private static string? ResolveThemeColor(string? value, PowerPointDesignTheme? theme) {
        if (string.IsNullOrWhiteSpace(value)) {
            return null;
        }

        var hex = ToPowerPointColor(value);
        if (!string.IsNullOrWhiteSpace(hex)) {
            return hex;
        }

        if (theme == null) {
            return null;
        }

        switch (Normalize(value!)) {
            case "primary":
                return theme.AccentDarkColor;
            case "accent":
            case "accent1":
            case "brand":
                return theme.AccentColor;
            case "accentdark":
                return theme.AccentDarkColor;
            case "accentlight":
                return theme.AccentLightColor;
            case "accent2":
            case "secondary":
                return theme.Accent2Color;
            case "accent3":
            case "tertiary":
                return theme.Accent3Color;
            case "warning":
                return theme.WarningColor;
            case "background":
            case "background1":
            case "bg1":
                return theme.BackgroundColor;
            case "surface":
            case "background2":
            case "bg2":
                return theme.SurfaceColor;
            case "panel":
                return theme.PanelColor;
            case "panelborder":
            case "border":
                return theme.PanelBorderColor;
            case "text":
            case "text1":
            case "foreground":
                return theme.PrimaryTextColor;
            case "text2":
            case "secondarytext":
                return theme.SecondaryTextColor;
            case "muted":
            case "mutedtext":
                return theme.MutedTextColor;
            case "white":
                return "FFFFFF";
            case "black":
                return "000000";
            default:
                return null;
        }
    }

    private sealed class OfficeMarkupBackgroundSpec {
        public string? Color { get; set; }
        public string? GradientStartColor { get; set; }
        public string? GradientEndColor { get; set; }
        public double? GradientAngleDegrees { get; set; }
        public string? ImagePath { get; set; }
        public string? OverlayColor { get; set; }
        public int? OverlayTransparency { get; set; }
    }
}
