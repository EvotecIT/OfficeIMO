// Generated from Build/Visual/themes.json by generate_theme_builtins.py. Do not edit by hand.
using OfficeIMO.Drawing;

namespace OfficeIMO;

public sealed partial class OfficeTheme {
    /// <summary>Gets a clean sans-serif theme with a cobalt accent and airy, tinted-header tables (the default).</summary>
    public static OfficeTheme Modern { get; } = CreateModern();
    /// <summary>Gets a traditional serif document theme with navy headings and an ochre title rule.</summary>
    public static OfficeTheme Classic { get; } = CreateClassic();
    /// <summary>Gets a restrained theme without fills whose fonts have metric-compatible open substitutes, so layout matches across office suites.</summary>
    public static OfficeTheme Minimal { get; } = CreateMinimal();
    /// <summary>Gets a formal theme with regular-weight serif headings, a gold rule and generous margins.</summary>
    public static OfficeTheme Boardroom { get; } = CreateBoardroom();
    /// <summary>Gets a screen theme with light text on a dark background; Word and Excel keep a light page.</summary>
    public static OfficeTheme Dark { get; } = CreateDark();

    private static OfficeTheme CreateModern() {
        var palette = new OfficeThemePalette(C("FFFFFF"), C("1B2433"), C("5B6575"), C("3F64D7"), C("DC7536"), C("058E80"), C("A52758"), C("3F2F6B"), C("255319"), C("08684A"), C("865500"), C("BF3B2C"));
        var fonts = new OfficeThemeFonts(F("Aptos", "Calibri", "Carlito", "Helvetica Neue", "Arial", "Liberation Sans", "sans-serif"), F("Aptos", "Calibri", "Carlito", "Helvetica Neue", "Arial", "Liberation Sans", "sans-serif"));
        var typeScale = new OfficeThemeTypeScale(P(28), P(20), P(15), P(12.5), P(11), P(11), P(10), P(11), P(9));
        var table = new OfficeThemeTableStyle(C("EAF0FC"), OfficeColor.Text, true, C("F6F8FB"), C("DCE2EA"), P(0.5), P(5),
            OfficeThemeTableBorderMode.Horizontal, new OfficeThemeRule(OfficeColor.Accent1, P(1)), true, OfficeAlignment.Right);
        var chart = new OfficeThemeChartStyle(new[] { OfficeColor.Accent1, OfficeColor.Accent2, OfficeColor.Accent3, OfficeColor.Accent4, OfficeColor.Accent5, OfficeColor.Accent6 },
            F("Aptos", "Calibri", "Carlito", "Helvetica Neue", "Arial", "Liberation Sans", "sans-serif"), F("Aptos", "Calibri", "Carlito", "Helvetica Neue", "Arial", "Liberation Sans", "sans-serif"),
            P(9), P(12), C("E6EAF0"), false, true,
            C("C9D0DA"), OfficeColor.TextMuted, P(0.5), P(2), 60, OfficeChartMarkerShape.None);
        var spacing = new OfficeThemeSpacing(P(0), P(8), P(18), P(36), P(48), P(36), P(48));
        var text = new OfficeThemeTextStyle(1.2, OfficeColor.Text, OfficeColor.Accent1, OfficeColor.Text, OfficeColor.Accent1, true, OfficeThemeRule.None,
            S(0, 6), S(18, 6), S(14, 4), S(12, 3));
        var surfaces = new OfficeThemeSurfaces(C("F6F8FB"), C("DCE2EA"), C("ECF0FB"), C("E1EDE9"), C("F0EBE0"), C("F7E7E6"));
        var page = new OfficeThemePage(P(64), P(64), P(56), P(64));
        var footer = new OfficeThemeFooter(P(8.5), OfficeColor.TextMuted, OfficeThemeRule.None);
        return new OfficeTheme("Modern", palette, fonts, typeScale, table, chart, spacing, text, surfaces, page, footer);
    }

    private static OfficeTheme CreateClassic() {
        var palette = new OfficeThemePalette(C("FFFFFF"), C("1F1F1F"), C("5C5850"), C("1F3A5F"), C("BB8842"), C("8A1D3E"), C("03655B"), C("3A85AD"), C("A3553A"), C("08684A"), C("865500"), C("BF3B2C"));
        var fonts = new OfficeThemeFonts(F("Cambria", "Caladea", "Georgia", "Liberation Serif", "Times New Roman", "serif"), F("Cambria", "Caladea", "Georgia", "Liberation Serif", "Times New Roman", "serif"));
        var typeScale = new OfficeThemeTypeScale(P(30), P(20), P(15.5), P(13), P(11.5), P(11), P(10.5), P(11), P(9));
        var table = new OfficeThemeTableStyle(OfficeColor.Accent1, OfficeColor.Background, true, C("F7F4EE"), C("D9D1C3"), P(0.5), P(5),
            OfficeThemeTableBorderMode.Horizontal, OfficeThemeRule.None, true, OfficeAlignment.Right);
        var chart = new OfficeThemeChartStyle(new[] { OfficeColor.Accent1, OfficeColor.Accent2, OfficeColor.Accent3, OfficeColor.Accent4, OfficeColor.Accent5, OfficeColor.Accent6 },
            F("Calibri", "Carlito", "Aptos", "Helvetica Neue", "Arial", "Liberation Sans", "sans-serif"), F("Calibri", "Carlito", "Aptos", "Helvetica Neue", "Arial", "Liberation Sans", "sans-serif"),
            P(9), P(12), C("E7E2D8"), false, true,
            C("BDB4A4"), OfficeColor.TextMuted, P(0.5), P(1.75), 70, OfficeChartMarkerShape.Circle);
        var spacing = new OfficeThemeSpacing(P(0), P(10), P(20), P(40), P(54), P(40), P(54));
        var text = new OfficeThemeTextStyle(1.25, OfficeColor.Accent1, C("8A6427"), OfficeColor.Accent1, OfficeColor.Accent1, true, new OfficeThemeRule(OfficeColor.Accent2, P(1.5)),
            S(0, 4), S(20, 8), S(16, 6), S(12, 4));
        var surfaces = new OfficeThemeSurfaces(C("F7F4EE"), C("D9D1C3"), C("E9EBEF"), C("E1EDE9"), C("F0EBE0"), C("F7E7E6"));
        var page = new OfficeThemePage(P(72), P(72), P(64), P(72));
        var footer = new OfficeThemeFooter(P(9), OfficeColor.TextMuted, new OfficeThemeRule(C("D9D1C3"), P(0.5)));
        return new OfficeTheme("Classic", palette, fonts, typeScale, table, chart, spacing, text, surfaces, page, footer);
    }

    private static OfficeTheme CreateMinimal() {
        var palette = new OfficeThemePalette(C("FFFFFF"), C("1A1A1A"), C("616161"), C("353535"), C("BE3E2B"), C("0287B8"), C("B28C54"), C("00655C"), C("5A3E6D"), C("08684A"), C("865500"), C("BF3B2C"));
        var fonts = new OfficeThemeFonts(F("Arial", "Liberation Sans", "Arimo", "Helvetica Neue", "Helvetica", "sans-serif"), F("Arial", "Liberation Sans", "Arimo", "Helvetica Neue", "Helvetica", "sans-serif"));
        var typeScale = new OfficeThemeTypeScale(P(24), P(16), P(12.5), P(10.5), P(10), P(10), P(9), P(10), P(8.5));
        var table = new OfficeThemeTableStyle(OfficeColor.Background, OfficeColor.Text, false, C("F4F4F4"), C("D6D6D6"), P(0.5), P(4),
            OfficeThemeTableBorderMode.Horizontal, new OfficeThemeRule(OfficeColor.Text, P(1)), true, OfficeAlignment.Right);
        var chart = new OfficeThemeChartStyle(new[] { OfficeColor.Accent1, OfficeColor.Accent2, OfficeColor.Accent3, OfficeColor.Accent4, OfficeColor.Accent5, OfficeColor.Accent6 },
            F("Arial", "Liberation Sans", "Arimo", "Helvetica Neue", "Helvetica", "sans-serif"), F("Arial", "Liberation Sans", "Arimo", "Helvetica Neue", "Helvetica", "sans-serif"),
            P(8), P(10.5), C("E8E8E8"), false, true,
            C("1A1A1A"), OfficeColor.TextMuted, P(0.5), P(1.5), 80, OfficeChartMarkerShape.None);
        var spacing = new OfficeThemeSpacing(P(0), P(6), P(14), P(32), P(40), P(32), P(40));
        var text = new OfficeThemeTextStyle(1.25, OfficeColor.Text, OfficeColor.TextMuted, OfficeColor.Text, OfficeColor.Text, true, OfficeThemeRule.None,
            S(0, 4), S(16, 4), S(12, 3), S(10, 2));
        var surfaces = new OfficeThemeSurfaces(C("F4F4F4"), C("D6D6D6"), C("EBEBEB"), C("E1EDE9"), C("F0EBE0"), C("F7E7E6"));
        var page = new OfficeThemePage(P(56), P(56), P(48), P(56));
        var footer = new OfficeThemeFooter(P(8), OfficeColor.TextMuted, OfficeThemeRule.None);
        return new OfficeTheme("Minimal", palette, fonts, typeScale, table, chart, spacing, text, surfaces, page, footer);
    }

    private static OfficeTheme CreateBoardroom() {
        var palette = new OfficeThemePalette(C("FFFFFF"), C("16202E"), C("586274"), C("14335C"), C("B0883E"), C("5181B9"), C("882139"), C("00655C"), C("766386"), C("08684A"), C("865500"), C("BF3B2C"));
        var fonts = new OfficeThemeFonts(F("Georgia", "Gelasio", "Cambria", "Caladea", "Liberation Serif", "serif"), F("Aptos", "Calibri", "Carlito", "Helvetica Neue", "Arial", "Liberation Sans", "sans-serif"));
        var typeScale = new OfficeThemeTypeScale(P(32), P(22), P(16), P(13), P(12), P(11), P(10.5), P(11), P(9));
        var table = new OfficeThemeTableStyle(OfficeColor.Accent1, OfficeColor.Background, true, C("F3F5F8"), C("D3D9E2"), P(0.5), P(6),
            OfficeThemeTableBorderMode.Horizontal, OfficeThemeRule.None, true, OfficeAlignment.Right);
        var chart = new OfficeThemeChartStyle(new[] { OfficeColor.Accent1, OfficeColor.Accent2, OfficeColor.Accent3, OfficeColor.Accent4, OfficeColor.Accent5, OfficeColor.Accent6 },
            F("Aptos", "Calibri", "Carlito", "Helvetica Neue", "Arial", "Liberation Sans", "sans-serif"), F("Aptos", "Calibri", "Carlito", "Helvetica Neue", "Arial", "Liberation Sans", "sans-serif"),
            P(9), P(12), C("E4E8EE"), true, true,
            C("C3CAD5"), OfficeColor.TextMuted, P(0.5), P(2.25), 60, OfficeChartMarkerShape.None);
        var spacing = new OfficeThemeSpacing(P(0), P(10), P(20), P(44), P(60), P(40), P(60));
        var text = new OfficeThemeTextStyle(1.25, OfficeColor.Accent1, C("876630"), OfficeColor.Accent1, OfficeColor.Accent1, false, new OfficeThemeRule(OfficeColor.Accent2, P(2)),
            S(0, 6), S(22, 8), S(16, 6), S(12, 4));
        var surfaces = new OfficeThemeSurfaces(C("F3F5F8"), C("D3D9E2"), C("E8EBEF"), C("E1EDE9"), C("F0EBE0"), C("F7E7E6"));
        var page = new OfficeThemePage(P(72), P(72), P(64), P(72));
        var footer = new OfficeThemeFooter(P(8.5), OfficeColor.TextMuted, new OfficeThemeRule(C("D3D9E2"), P(0.5)));
        return new OfficeTheme("Boardroom", palette, fonts, typeScale, table, chart, spacing, text, surfaces, page, footer);
    }

    private static OfficeTheme CreateDark() {
        var palette = new OfficeThemePalette(C("0F1620"), C("E7ECF2"), C("9AA6B5"), C("68A1F9"), C("F2B349"), C("90E7D6"), C("C84F7B"), C("675DAD"), C("6E974B"), C("3FA66E"), C("F2C14E"), C("FF8F8F"));
        var fonts = new OfficeThemeFonts(F("Aptos", "Calibri", "Carlito", "Helvetica Neue", "Arial", "Liberation Sans", "sans-serif"), F("Aptos", "Calibri", "Carlito", "Helvetica Neue", "Arial", "Liberation Sans", "sans-serif"));
        var typeScale = new OfficeThemeTypeScale(P(28), P(20), P(15), P(12.5), P(11), P(11), P(10), P(11), P(9));
        var table = new OfficeThemeTableStyle(C("1E2A39"), OfficeColor.Text, true, C("151F2B"), C("2A3747"), P(0.5), P(5),
            OfficeThemeTableBorderMode.Horizontal, new OfficeThemeRule(OfficeColor.Accent1, P(1)), true, OfficeAlignment.Right);
        var chart = new OfficeThemeChartStyle(new[] { OfficeColor.Accent1, OfficeColor.Accent2, OfficeColor.Accent3, OfficeColor.Accent4, OfficeColor.Accent5, OfficeColor.Accent6 },
            F("Aptos", "Calibri", "Carlito", "Helvetica Neue", "Arial", "Liberation Sans", "sans-serif"), F("Aptos", "Calibri", "Carlito", "Helvetica Neue", "Arial", "Liberation Sans", "sans-serif"),
            P(9), P(12), C("243142"), false, true,
            C("3A4859"), OfficeColor.TextMuted, P(0.5), P(2), 60, OfficeChartMarkerShape.None);
        var spacing = new OfficeThemeSpacing(P(0), P(8), P(18), P(36), P(48), P(36), P(48));
        var text = new OfficeThemeTextStyle(1.25, OfficeColor.Text, OfficeColor.Accent1, OfficeColor.Text, OfficeColor.Accent1, true, OfficeThemeRule.None,
            S(0, 6), S(18, 6), S(14, 4), S(12, 3));
        var surfaces = new OfficeThemeSurfaces(C("151F2B"), C("2A3747"), C("182436"), C("18302E"), C("383528"), C("3A2C34"));
        var page = new OfficeThemePage(P(64), P(64), P(56), P(64));
        var footer = new OfficeThemeFooter(P(8.5), OfficeColor.TextMuted, OfficeThemeRule.None);
        return new OfficeTheme("Dark", palette, fonts, typeScale, table, chart, spacing, text, surfaces, page, footer);
    }

    private static OfficeColor C(string value) => OfficeColor.ParseHex(value);
    private static OfficeLength P(double value) => OfficeLength.Points(value);
    private static OfficeThemeFont F(string family, params string[] fallbacks) => new OfficeThemeFont(family, fallbacks);
    private static OfficeThemeParagraphSpacing S(double before, double after) => new OfficeThemeParagraphSpacing(P(before), P(after));
}
