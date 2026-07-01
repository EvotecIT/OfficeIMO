using A = DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using System.Globalization;

namespace OfficeIMO.Excel.LegacyXls.Write {
    internal sealed class LegacyXlsThemeColorResolver {
        private readonly Dictionary<uint, string> _colors;

        private LegacyXlsThemeColorResolver(Dictionary<uint, string> colors) {
            _colors = colors;
        }

        internal static LegacyXlsThemeColorResolver Create(ExcelDocument document) {
            if (document == null) throw new ArgumentNullException(nameof(document));

            var colors = new Dictionary<uint, string>();
            WorkbookPart? workbookPart = document.WorkbookPartRoot;
            A.ColorScheme? scheme = workbookPart
                ?.GetPartsOfType<ThemePart>()
                .FirstOrDefault()
                ?.Theme
                ?.ThemeElements
                ?.ColorScheme;

            if (scheme != null) {
                AddThemeColor(colors, 0U, scheme.GetFirstChild<A.Light1Color>());
                AddThemeColor(colors, 1U, scheme.GetFirstChild<A.Dark1Color>());
                AddThemeColor(colors, 2U, scheme.GetFirstChild<A.Light2Color>());
                AddThemeColor(colors, 3U, scheme.GetFirstChild<A.Dark2Color>());
                AddThemeColor(colors, 4U, scheme.GetFirstChild<A.Accent1Color>());
                AddThemeColor(colors, 5U, scheme.GetFirstChild<A.Accent2Color>());
                AddThemeColor(colors, 6U, scheme.GetFirstChild<A.Accent3Color>());
                AddThemeColor(colors, 7U, scheme.GetFirstChild<A.Accent4Color>());
                AddThemeColor(colors, 8U, scheme.GetFirstChild<A.Accent5Color>());
                AddThemeColor(colors, 9U, scheme.GetFirstChild<A.Accent6Color>());
                AddThemeColor(colors, 10U, scheme.GetFirstChild<A.Hyperlink>());
                AddThemeColor(colors, 11U, scheme.GetFirstChild<A.FollowedHyperlinkColor>());
            }

            if (colors.Count == 0) {
                AddDefaultThemeColors(colors);
            }

            return new LegacyXlsThemeColorResolver(colors);
        }

        internal bool TryResolve(uint themeIndex, double? tint, out string? argb) {
            if (!_colors.TryGetValue(themeIndex, out argb)) {
                argb = null;
                return false;
            }

            if (tint.HasValue) {
                argb = ApplyTint(argb, tint.Value);
            }

            return true;
        }

        private static void AddDefaultThemeColors(Dictionary<uint, string> colors) {
            colors[0U] = "FFFFFFFF";
            colors[1U] = "FF000000";
            colors[2U] = "FFEEECE1";
            colors[3U] = "FF1F497D";
            colors[4U] = "FF4F81BD";
            colors[5U] = "FFC0504D";
            colors[6U] = "FF9BBB59";
            colors[7U] = "FF8064A2";
            colors[8U] = "FF4BACC6";
            colors[9U] = "FFF79646";
            colors[10U] = "FF0000FF";
            colors[11U] = "FF800080";
        }

        private static void AddThemeColor(Dictionary<uint, string> colors, uint index, OpenXmlCompositeElement? colorElement) {
            string? rgb = colorElement?.GetFirstChild<A.RgbColorModelHex>()?.Val?.Value
                ?? colorElement?.GetFirstChild<A.SystemColor>()?.LastColor?.Value;
            if (string.IsNullOrWhiteSpace(rgb)) {
                return;
            }

            string? argb = NormalizeArgb(rgb!);
            if (argb != null) {
                colors[index] = argb;
            }
        }

        private static string? NormalizeArgb(string value) {
            string hex = value.Trim().TrimStart('#');
            if (hex.Length == 6) {
                hex = "FF" + hex;
            }

            if (hex.Length != 8 || hex.Any(ch => !Uri.IsHexDigit(ch))) {
                return null;
            }

            return hex.ToUpperInvariant();
        }

        private static string ApplyTint(string argb, double tint) {
            byte alpha = Convert.ToByte(argb.Substring(0, 2), 16);
            byte red = Convert.ToByte(argb.Substring(2, 2), 16);
            byte green = Convert.ToByte(argb.Substring(4, 2), 16);
            byte blue = Convert.ToByte(argb.Substring(6, 2), 16);
            return alpha.ToString("X2", CultureInfo.InvariantCulture)
                + ApplyTintChannel(red, tint).ToString("X2", CultureInfo.InvariantCulture)
                + ApplyTintChannel(green, tint).ToString("X2", CultureInfo.InvariantCulture)
                + ApplyTintChannel(blue, tint).ToString("X2", CultureInfo.InvariantCulture);
        }

        private static byte ApplyTintChannel(byte channel, double tint) {
            double value = tint < 0D
                ? channel * (1D + tint)
                : channel * (1D - tint) + (255D * tint);
            return (byte)Math.Max(0D, Math.Min(255D, Math.Round(value)));
        }
    }
}
