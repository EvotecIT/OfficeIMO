using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.Excel.Utilities {
    internal static class ExcelThemeColorResolver {
        private static readonly string?[] IndexedColors = {
            "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF",
            "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF",
            "800000", "008000", "000080", "808000", "800080", "008080", "C0C0C0", "808080",
            "9999FF", "993366", "FFFFCC", "CCFFFF", "660066", "FF8080", "0066CC", "CCCCFF",
            "000080", "FF00FF", "FFFF00", "00FFFF", "800080", "800000", "008080", "0000FF",
            "00CCFF", "CCFFFF", "CCFFCC", "FFFF99", "99CCFF", "FF99CC", "CC99FF", "FFCC99",
            "3366FF", "33CCCC", "99CC00", "FFCC00", "FF9900", "FF6600", "666699", "969696",
            "003366", "339966", "003300", "333300", "993300", "993366", "333399", "333333"
        };

        internal static string? Resolve(OpenXmlElement? colorElement, WorkbookPart? workbookPart) {
            return colorElement is ColorType color ? Resolve(color, workbookPart) : null;
        }

        internal static string? Resolve(ColorType? color, WorkbookPart? workbookPart) {
            if (color == null) {
                return null;
            }

            string alpha = "FF";
            string? rgb = null;
            string? directArgb = NormalizeArgb(color.Rgb?.Value);
            if (directArgb != null) {
                alpha = directArgb.Substring(0, 2);
                rgb = directArgb.Substring(2);
            }

            if (rgb == null && color.Theme?.Value is uint themeIndex) {
                rgb = ResolveThemeRgb(workbookPart, themeIndex);
            }

            if (rgb == null && color.Indexed?.Value is uint indexed && indexed < IndexedColors.Length) {
                rgb = IndexedColors[indexed];
            }

            if (rgb == null || !TryParseRgb(rgb, out int red, out int green, out int blue)) {
                return null;
            }

            if (color.Tint?.Value is double tint && Math.Abs(tint) > double.Epsilon) {
                red = ApplyTint(red, tint);
                green = ApplyTint(green, tint);
                blue = ApplyTint(blue, tint);
            }

            return alpha + red.ToString("X2") + green.ToString("X2") + blue.ToString("X2");
        }

        private static string? ResolveThemeRgb(WorkbookPart? workbookPart, uint themeIndex) {
            A.ColorScheme? scheme = workbookPart?
                .GetPartsOfType<ThemePart>()
                .FirstOrDefault()?
                .Theme?
                .ThemeElements?
                .ColorScheme;
            if (scheme == null) {
                return null;
            }

            OpenXmlCompositeElement? color = themeIndex switch {
                0 => scheme.GetFirstChild<A.Light1Color>(),
                1 => scheme.GetFirstChild<A.Dark1Color>(),
                2 => scheme.GetFirstChild<A.Light2Color>(),
                3 => scheme.GetFirstChild<A.Dark2Color>(),
                4 => scheme.GetFirstChild<A.Accent1Color>(),
                5 => scheme.GetFirstChild<A.Accent2Color>(),
                6 => scheme.GetFirstChild<A.Accent3Color>(),
                7 => scheme.GetFirstChild<A.Accent4Color>(),
                8 => scheme.GetFirstChild<A.Accent5Color>(),
                9 => scheme.GetFirstChild<A.Accent6Color>(),
                10 => scheme.GetFirstChild<A.Hyperlink>(),
                11 => scheme.GetFirstChild<A.FollowedHyperlinkColor>(),
                _ => null
            };

            return NormalizeRgb(color?.GetFirstChild<A.RgbColorModelHex>()?.Val?.Value)
                ?? NormalizeRgb(color?.GetFirstChild<A.SystemColor>()?.LastColor?.Value);
        }

        private static string? NormalizeRgb(string? value) {
            if (string.IsNullOrWhiteSpace(value)) {
                return null;
            }

            string hex = value!.Trim().TrimStart('#');
            if (hex.Length == 8) {
                hex = hex.Substring(2);
            } else if (hex.Length != 6) {
                return null;
            }

            for (int i = 0; i < hex.Length; i++) {
                char ch = hex[i];
                bool isHex = (ch >= '0' && ch <= '9') ||
                    (ch >= 'a' && ch <= 'f') ||
                    (ch >= 'A' && ch <= 'F');
                if (!isHex) {
                    return null;
                }
            }

            return hex.ToUpperInvariant();
        }

        private static string? NormalizeArgb(string? value) {
            if (string.IsNullOrWhiteSpace(value)) {
                return null;
            }

            string hex = value!.Trim().TrimStart('#');
            if (hex.Length == 6) {
                hex = "FF" + hex;
            } else if (hex.Length != 8) {
                return null;
            }

            for (int i = 0; i < hex.Length; i++) {
                char ch = hex[i];
                bool isHex = (ch >= '0' && ch <= '9') ||
                    (ch >= 'a' && ch <= 'f') ||
                    (ch >= 'A' && ch <= 'F');
                if (!isHex) {
                    return null;
                }
            }

            return hex.ToUpperInvariant();
        }

        private static bool TryParseRgb(string rgb, out int red, out int green, out int blue) {
            red = green = blue = 0;
            if (rgb.Length != 6) {
                return false;
            }

            red = Convert.ToInt32(rgb.Substring(0, 2), 16);
            green = Convert.ToInt32(rgb.Substring(2, 2), 16);
            blue = Convert.ToInt32(rgb.Substring(4, 2), 16);
            return true;
        }

        private static int ApplyTint(int channel, double tint) {
            double value = tint < 0D
                ? channel * (1D + tint)
                : channel + (255D - channel) * tint;
            return Math.Max(0, Math.Min(255, (int)Math.Round(value, MidpointRounding.AwayFromZero)));
        }
    }
}
