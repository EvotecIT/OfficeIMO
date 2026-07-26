using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.Word {
    internal static partial class WordDocumentImageRenderer {
        private static bool TryResolveShadingFillColor(
            Shading? shading,
            A.ColorScheme? colorScheme,
            out OfficeColor fillColor) {
            fillColor = OfficeColor.White;
            if (shading == null || shading.Val?.Value == ShadingPatternValues.Nil) {
                return false;
            }

            string pattern = GetWordAttribute(shading, "val") ?? "clear";
            bool hasBackground = TryResolveShadingBackground(shading, colorScheme, out OfficeColor background);
            if (pattern.Equals("clear", StringComparison.OrdinalIgnoreCase)) {
                if (hasBackground) {
                    fillColor = background;
                }
                return hasBackground;
            }

            OfficeColor foreground = ResolveShadingForeground(shading, colorScheme);
            if (pattern.Equals("solid", StringComparison.OrdinalIgnoreCase)) {
                fillColor = foreground;
                return true;
            }

            background = hasBackground ? background : OfficeColor.White;
            fillColor = BlendShadingPattern(
                foreground,
                background,
                ResolveShadingPatternForegroundRatio(pattern));
            return true;
        }

        private static OfficeColor ResolveShadingForeground(
            Shading shading,
            A.ColorScheme? colorScheme) {
            string? resolvedThemeForeground = ResolveThemeColor(
                GetWordAttribute(shading, "themeColor"),
                GetWordAttribute(shading, "themeTint"),
                GetWordAttribute(shading, "themeShade"),
                colorScheme);
            if (TryParseOfficeColor(resolvedThemeForeground, out OfficeColor themeForeground)) {
                return themeForeground;
            }
            if (TryParseOfficeColor(shading.Color?.Value, out OfficeColor foreground)) {
                return foreground;
            }

            return OfficeColor.Black;
        }

        private static bool TryResolveShadingBackground(
            Shading shading,
            A.ColorScheme? colorScheme,
            out OfficeColor background) {
            string? resolvedThemeBackground = ResolveThemeColor(
                GetWordAttribute(shading, "themeFill"),
                GetWordAttribute(shading, "themeFillTint"),
                GetWordAttribute(shading, "themeFillShade"),
                colorScheme);
            if (TryParseOfficeColor(resolvedThemeBackground, out background)) {
                return true;
            }

            return TryParseOfficeColor(shading.Fill?.Value, out background);
        }

        private static double ResolveShadingPatternForegroundRatio(string pattern) {
            if (pattern.StartsWith("pct", StringComparison.OrdinalIgnoreCase) &&
                int.TryParse(pattern.Substring(3), out int percentage)) {
                return percentage switch {
                    12 => 0.125d,
                    37 => 0.375d,
                    62 => 0.625d,
                    87 => 0.875d,
                    _ => Math.Max(0d, Math.Min(1d, percentage / 100d))
                };
            }

            return pattern.StartsWith("thin", StringComparison.OrdinalIgnoreCase)
                ? 0.25d
                : 0.5d;
        }

        private static OfficeColor BlendShadingPattern(
            OfficeColor foreground,
            OfficeColor background,
            double foregroundRatio) {
            double backgroundRatio = 1d - foregroundRatio;
            return OfficeColor.FromRgb(
                (byte)Math.Round(foreground.R * foregroundRatio + background.R * backgroundRatio),
                (byte)Math.Round(foreground.G * foregroundRatio + background.G * backgroundRatio),
                (byte)Math.Round(foreground.B * foregroundRatio + background.B * backgroundRatio));
        }
    }
}
