using OfficeIMO.Drawing;
using SixLabors.Fonts;

namespace OfficeIMO.Excel {
    internal sealed class ExcelTextMeasurer {
        private readonly SixLabors.Fonts.Font _fallbackFont;
        private readonly OfficeFontInfo _fallbackFontInfo;

        private ExcelTextMeasurer(SixLabors.Fonts.Font fallbackFont, OfficeFontInfo fallbackFontInfo) {
            _fallbackFont = fallbackFont;
            _fallbackFontInfo = fallbackFontInfo;
            DefaultStyle = CreateStyle(fallbackFont, dpi: null);
        }

        internal Style DefaultStyle { get; }

        internal float DefaultFontSize => _fallbackFont.Size;

        internal OfficeFontInfo FallbackFontInfo => _fallbackFontInfo;

        internal static ExcelTextMeasurer Create(OfficeFontInfo? workbookDefaultFontInfo) {
            var fallbackInfo = workbookDefaultFontInfo ?? OfficeFontInfo.Default;
            var fallbackFont = ResolveDefaultFont(workbookDefaultFontInfo);
            return new ExcelTextMeasurer(fallbackFont, fallbackInfo);
        }

        internal Style CreateDefaultStyle(float dpi)
            => CreateStyle(_fallbackFont, dpi);

        internal Style CreateStyle(OfficeFontInfo fontInfo) {
            var font = CreateFontFromInfo(fontInfo, _fallbackFont);
            return CreateStyle(font, dpi: null);
        }

        internal Style CreateStyle(OfficeFontInfo fontInfo, float dpi) {
            var font = CreateFontFromInfo(fontInfo, _fallbackFont);
            return CreateStyle(font, (float?)dpi);
        }

        internal float MeasureWidthOrDefault(string text, Style style, float fallback) {
            try {
                float measured = TextMeasurer.MeasureSize(text, style.Options).Width;
                return measured > 0.0001f ? measured : fallback;
            } catch {
                return fallback;
            }
        }

        internal float MeasureHeightOrDefault(string text, Style style, float fallback) {
            try {
                float measured = TextMeasurer.MeasureSize(text, style.Options).Height;
                return measured > 0.0001f ? measured : fallback;
            } catch {
                return fallback;
            }
        }

        private static Style CreateStyle(SixLabors.Fonts.Font font, float? dpi) {
            var options = new TextOptions(font);
            if (dpi != null) {
                options.Dpi = dpi.Value;
            }

            float mdw = MeasureWidthOrDefault("0", options, fallback: 0);
            return new Style(options, mdw);
        }

        private static SixLabors.Fonts.Font ResolveDefaultFont(OfficeFontInfo? workbookDefaultFontInfo) {
            if (workbookDefaultFontInfo != null) {
                var font = TryCreateSystemFont(workbookDefaultFontInfo.Value);
                if (font != null && IsFontUsable(font)) {
                    return font;
                }
            }

            string[] preferred = { "Calibri", "Arial", "Liberation Sans", "DejaVu Sans", "Times New Roman" };

            foreach (var name in preferred) {
                try {
                    var font = SystemFonts.CreateFont(name, (float)OfficeFontInfo.Default.Size);
                    if (IsFontUsable(font)) {
                        return font;
                    }
                } catch (FontFamilyNotFoundException) {
                    // Try next option.
                }
            }

            foreach (var family in SystemFonts.Collection.Families) {
                try {
                    var font = family.CreateFont(11);
                    if (IsFontUsable(font)) {
                        return font;
                    }
                } catch {
                    // Skip fonts that cannot be loaded or measured.
                }
            }

            return SystemFonts.Collection.Families.First().CreateFont(11);
        }

        private static bool IsFontUsable(SixLabors.Fonts.Font font) {
            try {
                TextMeasurer.MeasureSize("0", new TextOptions(font));
                return true;
            } catch {
                return false;
            }
        }

        private static SixLabors.Fonts.Font CreateFontFromInfo(OfficeFontInfo fontInfo, SixLabors.Fonts.Font fallbackFont) {
            var style = ToSixLaborsFontStyle(fontInfo.Style);
            var size = (float)fontInfo.Size;
            try {
                if (!string.IsNullOrWhiteSpace(fontInfo.FamilyName)) {
                    return SystemFonts.CreateFont(fontInfo.FamilyName, size, style);
                }

                return fallbackFont.Family.CreateFont(size, style);
            } catch (FontFamilyNotFoundException) {
                return fallbackFont.Family.CreateFont(size, style);
            }
        }

        private static SixLabors.Fonts.Font? TryCreateSystemFont(OfficeFontInfo fontInfo) {
            if (string.IsNullOrWhiteSpace(fontInfo.FamilyName)) {
                return null;
            }

            try {
                return SystemFonts.CreateFont(fontInfo.FamilyName, (float)fontInfo.Size, ToSixLaborsFontStyle(fontInfo.Style));
            } catch (FontFamilyNotFoundException) {
                return null;
            }
        }

        private static FontStyle ToSixLaborsFontStyle(OfficeFontStyle style) {
            bool bold = (style & OfficeFontStyle.Bold) == OfficeFontStyle.Bold;
            bool italic = (style & OfficeFontStyle.Italic) == OfficeFontStyle.Italic;
            return bold && italic ? FontStyle.BoldItalic : bold ? FontStyle.Bold : italic ? FontStyle.Italic : FontStyle.Regular;
        }

        internal readonly struct Style {
            internal Style(TextOptions options, float maximumDigitWidth) {
                Options = options;
                MaximumDigitWidth = maximumDigitWidth;
            }

            internal TextOptions Options { get; }

            internal float MaximumDigitWidth { get; }
        }

        private static float MeasureWidthOrDefault(string text, TextOptions options, float fallback) {
            try {
                float measured = TextMeasurer.MeasureSize(text, options).Width;
                return measured > 0.0001f ? measured : fallback;
            } catch {
                return fallback;
            }
        }
    }
}
