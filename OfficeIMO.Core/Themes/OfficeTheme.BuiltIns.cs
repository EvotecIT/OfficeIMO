using OfficeIMO.Drawing;

namespace OfficeIMO;

// All built-in visual values live here so changes stay consistent across formats.
public sealed partial class OfficeTheme {
    /// <summary>Gets a clean sans-serif theme with blue accents.</summary>
    public static OfficeTheme Modern { get; } = Create("Modern", "FFFFFF", "182230", "667085", "2563EB", "Aptos", "Aptos", 11);
    /// <summary>Gets a traditional document theme with serif headings.</summary>
    public static OfficeTheme Classic { get; } = Create("Classic", "FFFFFF", "222222", "666666", "365F91", "Cambria", "Calibri", 11);
    /// <summary>Gets a restrained theme with compact typography.</summary>
    public static OfficeTheme Minimal { get; } = Create("Minimal", "FFFFFF", "202020", "737373", "525252", "Arial", "Arial", 10);
    /// <summary>Gets a formal theme with navy headings and generous spacing.</summary>
    public static OfficeTheme Boardroom { get; } = Create("Boardroom", "FFFFFF", "172B4D", "5E6C84", "17365D", "Aptos Display", "Aptos", 12);
    /// <summary>Gets a dark theme with light text and bright accents.</summary>
    public static OfficeTheme Dark { get; } = Create("Dark", "111827", "F3F4F6", "9CA3AF", "60A5FA", "Aptos", "Aptos", 11);

    private static OfficeTheme Create(string name, string background, string text, string muted, string accent, string heading, string body, double size) {
        bool dark = name == "Dark";
        var palette = new OfficeThemePalette(C(background), C(text), C(muted), C(accent), C("14B8A6"), C("8B5CF6"),
            C("F59E0B"), C("EC4899"), C("06B6D4"), C(dark ? "4ADE80" : "16803C"), C(dark ? "FBBF24" : "B45309"), C(dark ? "F87171" : "DC2626"));
        var fallbacks = new[] { "Arial", "Liberation Sans", "DejaVu Sans", "sans-serif" };
        var fonts = new OfficeThemeFonts(new OfficeThemeFont(heading, fallbacks), new OfficeThemeFont(body, fallbacks));
        var scale = new OfficeThemeTypeScale(P(32), P(26), P(22), P(18), P(16), P(14), P(12), P(size), P(9));
        var table = new OfficeThemeTableStyle(OfficeColor.Accent1, OfficeColor.White, true,
            C(dark ? "1F2937" : "F3F6FA"), C(dark ? "374151" : "D0D5DD"), P(0.5), P(6));
        var charts = new OfficeThemeChartStyle(new[] { OfficeColor.Accent1, OfficeColor.Accent2, OfficeColor.Accent3,
            OfficeColor.Accent4, OfficeColor.Accent5, OfficeColor.Accent6 }, fonts.Body, fonts.Heading, P(10), P(14),
            C(dark ? "374151" : "E5E7EB"), false, true);
        var spacing = new OfficeThemeSpacing(P(0), P(8), P(18), P(28), P(36), P(28), P(36));
        return new OfficeTheme(name, palette, fonts, scale, table, charts, spacing);
    }
    private static OfficeColor C(string value) => OfficeColor.ParseHex(value);
    private static OfficeLength P(double value) => OfficeLength.Points(value);
}
