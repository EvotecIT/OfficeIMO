using OfficeIMO.Drawing;

namespace OfficeIMO.Excel {
    internal sealed class ExcelTextMeasurer {
        private const float DefaultDpi = 96f;
        private const float PointsPerInch = 72f;
        private const float DefaultDigitEmWidth = 0.62f;

        private readonly OfficeFontInfo _fallbackFontInfo;

        private ExcelTextMeasurer(OfficeFontInfo fallbackFontInfo) {
            _fallbackFontInfo = NormalizeFontInfo(fallbackFontInfo);
            DefaultStyle = CreateStyle(_fallbackFontInfo, DefaultDpi);
        }

        internal Style DefaultStyle { get; }

        internal float DefaultFontSize => (float)_fallbackFontInfo.Size;

        internal OfficeFontInfo FallbackFontInfo => _fallbackFontInfo;

        internal static ExcelTextMeasurer Create(OfficeFontInfo? workbookDefaultFontInfo) =>
            new ExcelTextMeasurer(workbookDefaultFontInfo ?? OfficeFontInfo.Default);

        internal Style CreateDefaultStyle(float dpi)
            => CreateStyle(_fallbackFontInfo, dpi);

        internal Style CreateStyle(OfficeFontInfo fontInfo)
            => CreateStyle(fontInfo, DefaultDpi);

        internal Style CreateStyle(OfficeFontInfo fontInfo, float dpi) {
            var normalized = NormalizeFontInfo(fontInfo);
            var style = new Style(normalized, NormalizeDpi(dpi));
            return style.MaximumDigitWidth > 0.0001f
                ? style
                : new Style(OfficeFontInfo.Default, DefaultDpi);
        }

        internal float MeasureWidthOrDefault(string text, Style style, float fallback) {
            if (string.IsNullOrEmpty(text)) {
                return fallback;
            }

            float measured = MeasureTextWidth(text, style);
            return measured > 0.0001f ? measured : fallback;
        }

        internal float MeasureHeightOrDefault(string text, Style style, float fallback) {
            if (string.IsNullOrEmpty(text)) {
                return fallback;
            }

            float measured = style.FontSizePixels * GetLineHeightFactor(style.FontInfo);
            return measured > 0.0001f ? measured : fallback;
        }

        private static OfficeFontInfo NormalizeFontInfo(OfficeFontInfo fontInfo) {
            string familyName = string.IsNullOrWhiteSpace(fontInfo.FamilyName)
                ? OfficeFontInfo.Default.FamilyName
                : fontInfo.FamilyName;

            double size = fontInfo.Size > 0.1 ? fontInfo.Size : OfficeFontInfo.Default.Size;
            return new OfficeFontInfo(familyName, size, fontInfo.Style);
        }

        private static float NormalizeDpi(float dpi) =>
            dpi > 0.0001f ? dpi : DefaultDpi;

        private static float MeasureTextWidth(string text, Style style) {
            float width = 0;
            for (int i = 0; i < text.Length; i++) {
                char value = text[i];
                if (value == '\t') {
                    width += style.SpaceWidth * 4;
                    continue;
                }

                if (char.IsControl(value)) {
                    continue;
                }

                width += style.FontSizePixels * GetCharacterWidthFactor(value, style.FontInfo);
            }

            return width;
        }

        private static float GetCharacterWidthFactor(char value, OfficeFontInfo fontInfo) {
            float factor;
            if (value == ' ') {
                factor = 0.34f;
            } else if (value >= '0' && value <= '9') {
                factor = DefaultDigitEmWidth;
            } else if (value >= 'A' && value <= 'Z') {
                factor = IsNarrowUppercase(value) ? 0.38f : 0.68f;
            } else if (value >= 'a' && value <= 'z') {
                factor = IsNarrowLowercase(value) ? 0.28f : value == 'm' || value == 'w' ? 0.82f : 0.55f;
            } else if (IsCjkOrWide(value)) {
                factor = 1.0f;
            } else {
                factor = value switch {
                    '.' or ',' or ':' or ';' or '\'' or '"' or '`' => 0.28f,
                    '!' or '|' => 0.30f,
                    '-' or '_' => 0.40f,
                    '(' or ')' or '[' or ']' or '{' or '}' => 0.36f,
                    '/' or '\\' => 0.42f,
                    '+' or '=' or '<' or '>' => 0.58f,
                    '@' => 0.92f,
                    '#' or '$' or '&' => 0.72f,
                    '%' => 0.90f,
                    _ => 0.62f
                };
            }

            return factor * GetFontFamilyWidthFactor(fontInfo) * GetStyleWidthFactor(fontInfo);
        }

        private static bool IsNarrowUppercase(char value) =>
            value == 'I' || value == 'J';

        private static bool IsNarrowLowercase(char value) =>
            value == 'i' || value == 'j' || value == 'l' || value == 't' || value == 'f' || value == 'r';

        private static bool IsCjkOrWide(char value) =>
            (value >= '\u1100' && value <= '\u11ff')
            || (value >= '\u2e80' && value <= '\u9fff')
            || (value >= '\uf900' && value <= '\ufaff')
            || (value >= '\uff00' && value <= '\uffef');

        private static float GetFontFamilyWidthFactor(OfficeFontInfo fontInfo) {
            string name = fontInfo.FamilyName ?? string.Empty;
            if (Contains(name, "Courier")
                || Contains(name, "Consolas")
                || Contains(name, "Mono")) {
                return 1.12f;
            }

            if (Contains(name, "Times")
                || Contains(name, "Serif")
                || Contains(name, "Garamond")) {
                return 0.96f;
            }

            if (Contains(name, "Aptos")
                || Contains(name, "Calibri")
                || Contains(name, "Arial")
                || Contains(name, "Helvetica")
                || Contains(name, "Liberation Sans")
                || Contains(name, "DejaVu Sans")) {
                return 1.0f;
            }

            return 1.02f;
        }

        private static bool Contains(string value, string text) =>
            value.IndexOf(text, StringComparison.OrdinalIgnoreCase) >= 0;

        private static float GetStyleWidthFactor(OfficeFontInfo fontInfo) {
            float factor = 1.0f;
            if (fontInfo.IsBold) {
                factor *= 1.06f;
            }

            if (fontInfo.IsItalic) {
                factor *= 1.02f;
            }

            return factor;
        }

        private static float GetLineHeightFactor(OfficeFontInfo fontInfo) {
            float factor = 1.30f;
            if (fontInfo.IsBold) {
                factor += 0.03f;
            }

            if (fontInfo.IsUnderline) {
                factor += 0.06f;
            }

            return factor;
        }

        internal readonly struct Style {
            internal Style(OfficeFontInfo fontInfo, float dpi) {
                FontInfo = fontInfo;
                Dpi = dpi;
                FontSizePixels = (float)fontInfo.Size * dpi / PointsPerInch;
                SpaceWidth = FontSizePixels * 0.34f * GetFontFamilyWidthFactor(fontInfo) * GetStyleWidthFactor(fontInfo);
                MaximumDigitWidth = FontSizePixels * DefaultDigitEmWidth * GetFontFamilyWidthFactor(fontInfo) * GetStyleWidthFactor(fontInfo);
            }

            internal OfficeFontInfo FontInfo { get; }

            internal float Dpi { get; }

            internal float FontSizePixels { get; }

            internal float SpaceWidth { get; }

            internal float MaximumDigitWidth { get; }
        }
    }
}
