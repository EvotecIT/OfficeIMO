using System;
using System.Globalization;

namespace OfficeIMO.Drawing;

/// <summary>
/// Deterministic, zero-dependency text measurer for Office layout estimates.
/// </summary>
/// <remarks>
/// This measurer intentionally does not call operating-system font APIs. It provides stable
/// Office-oriented estimates for layout decisions such as autofit and wrapping.
/// </remarks>
public sealed class OfficeTextMeasurer {
    /// <summary>Default measurement DPI.</summary>
    public const double DefaultDpi = 96D;

    internal const double PointsPerInch = 72D;
    internal const double DefaultDigitEmWidth = 0.62D;

    private readonly OfficeFontInfo _fallbackFontInfo;

    private OfficeTextMeasurer(OfficeFontInfo fallbackFontInfo) {
        _fallbackFontInfo = NormalizeFontInfo(fallbackFontInfo);
        DefaultStyle = CreateStyle(_fallbackFontInfo, DefaultDpi);
    }

    /// <summary>Default style derived from the fallback font.</summary>
    public OfficeTextMeasurementStyle DefaultStyle { get; }

    /// <summary>Default font size in points.</summary>
    public double DefaultFontSize => _fallbackFontInfo.Size;

    /// <summary>Fallback font descriptor used when no explicit font is available.</summary>
    public OfficeFontInfo FallbackFontInfo => _fallbackFontInfo;

    /// <summary>Creates a deterministic text measurer.</summary>
    public static OfficeTextMeasurer Create(OfficeFontInfo? fallbackFontInfo = null) =>
        new OfficeTextMeasurer(fallbackFontInfo ?? OfficeFontInfo.Default);

    /// <summary>Creates a style from the measurer fallback font and a DPI value.</summary>
    public OfficeTextMeasurementStyle CreateDefaultStyle(double dpi) =>
        CreateStyle(_fallbackFontInfo, dpi);

    /// <summary>Creates a style from a font descriptor using the default DPI.</summary>
    public OfficeTextMeasurementStyle CreateStyle(OfficeFontInfo fontInfo) =>
        CreateStyle(fontInfo, DefaultDpi);

    /// <summary>Creates a style from a font descriptor and DPI value.</summary>
    public OfficeTextMeasurementStyle CreateStyle(OfficeFontInfo fontInfo, double dpi) {
        var style = new OfficeTextMeasurementStyle(fontInfo, dpi);
        return style.MaximumDigitWidthPixels > 0.0001D
            ? style
            : new OfficeTextMeasurementStyle(OfficeFontInfo.Default, DefaultDpi);
    }

    /// <summary>Measures text width and line metrics for the supplied style.</summary>
    public OfficeTextMetrics Measure(string? text, OfficeTextMeasurementStyle style) {
        double width = MeasureWidth(text, style);
        double lineHeight = MeasureLineHeight(style);
        return new OfficeTextMetrics(width, lineHeight, style.SpaceWidthPixels, style.MaximumDigitWidthPixels);
    }

    /// <summary>Measures text width in pixels for the supplied style.</summary>
    public double MeasureWidth(string? text, OfficeTextMeasurementStyle style) {
        if (string.IsNullOrEmpty(text)) {
            return 0D;
        }

        double width = 0D;
        foreach (string element in OfficeTextElements.Enumerate(text)) {
            if (element == "\t") {
                width += style.SpaceWidthPixels * 4D;
                continue;
            }

            width += style.FontSizePixels * GetTextElementWidthFactor(element, style.FontInfo);
        }

        return width;
    }

    /// <summary>Measures text width or returns the provided fallback when text cannot be measured.</summary>
    public double MeasureWidthOrDefault(string? text, OfficeTextMeasurementStyle style, double fallback) {
        double measured = MeasureWidth(text, style);
        return measured > 0.0001D ? measured : fallback;
    }

    /// <summary>Measures a single-line text height in pixels for the supplied style.</summary>
    public double MeasureLineHeight(OfficeTextMeasurementStyle style) =>
        style.FontSizePixels * GetLineHeightFactor(style.FontInfo);

    /// <summary>Measures line height or returns the provided fallback when text is empty.</summary>
    public double MeasureLineHeightOrDefault(string? text, OfficeTextMeasurementStyle style, double fallback) {
        if (string.IsNullOrEmpty(text)) {
            return fallback;
        }

        double measured = MeasureLineHeight(style);
        return measured > 0.0001D ? measured : fallback;
    }

    internal static OfficeFontInfo NormalizeFontInfo(OfficeFontInfo fontInfo) {
        string familyName = string.IsNullOrWhiteSpace(fontInfo.FamilyName)
            ? OfficeFontInfo.Default.FamilyName
            : fontInfo.FamilyName;

        double size = fontInfo.Size > 0.1D ? fontInfo.Size : OfficeFontInfo.Default.Size;
        return new OfficeFontInfo(familyName, size, fontInfo.Style);
    }

    internal static double NormalizeDpi(double dpi) =>
        dpi > 0.0001D ? dpi : DefaultDpi;

    internal static double GetCharacterWidthFactor(char value, OfficeFontInfo fontInfo) {
        double factor;
        if (value == ' ' || value == '\u00a0') {
            factor = 0.34D;
        } else if (value >= '0' && value <= '9') {
            factor = DefaultDigitEmWidth;
        } else if (value >= 'A' && value <= 'Z') {
            factor = IsNarrowUppercase(value) ? 0.38D : 0.68D;
        } else if (value >= 'a' && value <= 'z') {
            factor = IsNarrowLowercase(value) ? 0.28D : value == 'm' || value == 'w' ? 0.82D : 0.55D;
        } else if (IsCjkOrWide(value)) {
            factor = 1.0D;
        } else {
            factor = value switch {
                '.' or ',' or ':' or ';' or '\'' or '"' or '`' => 0.28D,
                '!' or '|' => 0.30D,
                '-' or '_' => 0.40D,
                '(' or ')' or '[' or ']' or '{' or '}' => 0.36D,
                '/' or '\\' => 0.42D,
                '+' or '=' or '<' or '>' => 0.58D,
                '@' => 0.92D,
                '#' or '$' or '&' => 0.72D,
                '%' => 0.90D,
                _ => 0.62D
            };
        }

        return factor * GetFontFamilyWidthFactor(fontInfo) * GetStyleWidthFactor(fontInfo);
    }

    private static double GetTextElementWidthFactor(string element, OfficeFontInfo fontInfo) {
        if (!TryGetBaseScalar(element, out int scalar)) return 0D;
        double factor = scalar <= char.MaxValue
            ? GetCharacterWidthFactor((char)scalar, fontInfo)
            : (IsCjkOrWide(scalar) || IsEmojiLike(scalar) ? 1D : 0.62D) * GetFontFamilyWidthFactor(fontInfo) * GetStyleWidthFactor(fontInfo);
        if (element.IndexOf('\u200d') >= 0 || element.IndexOf('\ufe0f') >= 0) {
            factor = 1D * GetFontFamilyWidthFactor(fontInfo) * GetStyleWidthFactor(fontInfo);
        }

        return factor;
    }

    private static bool TryGetBaseScalar(string element, out int scalar) {
        for (int index = 0; index < element.Length;) {
            bool surrogatePair = char.IsHighSurrogate(element[index])
                && index + 1 < element.Length
                && char.IsLowSurrogate(element[index + 1]);
            scalar = surrogatePair ? char.ConvertToUtf32(element[index], element[index + 1]) : element[index];
            UnicodeCategory category = surrogatePair
                ? CharUnicodeInfo.GetUnicodeCategory(element, index)
                : char.IsSurrogate(element[index]) ? UnicodeCategory.OtherNotAssigned : CharUnicodeInfo.GetUnicodeCategory(element, index);
            if (category != UnicodeCategory.Control
                && category != UnicodeCategory.Format
                && category != UnicodeCategory.NonSpacingMark
                && category != UnicodeCategory.SpacingCombiningMark
                && category != UnicodeCategory.EnclosingMark
                && scalar != 0x00AD
                && scalar != 0x200B) return true;
            index += surrogatePair ? 2 : 1;
        }

        scalar = 0;
        return false;
    }

    internal static double GetFontFamilyWidthFactor(OfficeFontInfo fontInfo) {
        string name = fontInfo.FamilyName ?? string.Empty;
        if (Contains(name, "Courier")
            || Contains(name, "Consolas")
            || Contains(name, "Mono")) {
            return 1.12D;
        }

        if (Contains(name, "Times")
            || Contains(name, "Serif")
            || Contains(name, "Garamond")) {
            return 0.96D;
        }

        if (Contains(name, "Aptos")
            || Contains(name, "Calibri")
            || Contains(name, "Arial")
            || Contains(name, "Helvetica")
            || Contains(name, "Liberation Sans")
            || Contains(name, "DejaVu Sans")) {
            return 1.0D;
        }

        return 1.02D;
    }

    internal static double GetStyleWidthFactor(OfficeFontInfo fontInfo) {
        double factor = 1.0D;
        if (fontInfo.IsBold) {
            factor *= 1.06D;
        }

        if (fontInfo.IsItalic) {
            factor *= 1.02D;
        }

        return factor;
    }

    internal static double GetLineHeightFactor(OfficeFontInfo fontInfo) {
        double factor = 1.30D;
        if (fontInfo.IsBold) {
            factor += 0.03D;
        }

        if (fontInfo.IsUnderline) {
            factor += 0.06D;
        }

        return factor;
    }

    private static bool IsNarrowUppercase(char value) =>
        value == 'I' || value == 'J';

    private static bool IsNarrowLowercase(char value) =>
        value == 'i' || value == 'j' || value == 'l' || value == 't' || value == 'f' || value == 'r';

    private static bool IsCjkOrWide(int value) =>
        (value >= '\u1100' && value <= '\u11ff')
        || (value >= '\u2e80' && value <= '\u9fff')
        || (value >= '\uf900' && value <= '\ufaff')
        || (value >= '\uff00' && value <= '\uffef')
        || (value >= 0x20000 && value <= 0x3FFFD);

    private static bool IsEmojiLike(int value) =>
        (value >= 0x1F000 && value <= 0x1FAFF)
        || (value >= 0x2600 && value <= 0x27BF)
        || (value >= 0x1F1E6 && value <= 0x1F1FF);

    private static bool Contains(string value, string text) =>
        value.IndexOf(text, StringComparison.OrdinalIgnoreCase) >= 0;
}
