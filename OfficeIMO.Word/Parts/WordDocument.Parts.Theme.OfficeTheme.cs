using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.Word;

public partial class WordDocument {
    /// <summary>Projects the selected shared palette and fonts into this document's native theme part.</summary>
    private static void ApplyCreationThemePart(ThemePart part, OfficeTheme theme) {
        A.Theme native = (A.Theme)part.Theme!.CloneNode(true);
        native.Name = theme.Name;
        A.ThemeElements elements = native.ThemeElements!;
        OfficeThemePalette p = theme.Palette;
        string Hex(OfficeColor value) => value.Resolve(theme).ToRgbHex();
        A.RgbColorModelHex Rgb(OfficeColor value) => new A.RgbColorModelHex { Val = Hex(value) };
        elements.ColorScheme = new A.ColorScheme(
            new A.Dark1Color(Rgb(p.Text)), new A.Light1Color(Rgb(p.Background)),
            new A.Dark2Color(Rgb(p.TextMuted)), new A.Light2Color(Rgb(theme.Surfaces.Surface)),
            new A.Accent1Color(Rgb(p.Accent1)), new A.Accent2Color(Rgb(p.Accent2)),
            new A.Accent3Color(Rgb(p.Accent3)), new A.Accent4Color(Rgb(p.Accent4)),
            new A.Accent5Color(Rgb(p.Accent5)), new A.Accent6Color(Rgb(p.Accent6)),
            new A.Hyperlink(Rgb(theme.TextStyle.LinkColor)), new A.FollowedHyperlinkColor(Rgb(p.Accent2))) { Name = theme.Name };
        A.FontScheme fonts = elements.FontScheme!;
        fonts.Name = theme.Name;
        SetCreationFont(fonts.MajorFont!, theme.Fonts.Heading.Family);
        SetCreationFont(fonts.MinorFont!, theme.Fonts.Body.Family);
        part.Theme = native;
    }

    private static void SetCreationFont(OpenXmlCompositeElement fonts, string family) {
        A.LatinFont latin = fonts.GetFirstChild<A.LatinFont>() ?? fonts.PrependChild(new A.LatinFont());
        latin.Typeface = family;
        latin.Panose = null;
        // Empty script defaults defer to the selected family; retain explicit supplemental script fonts.
        if (fonts.GetFirstChild<A.EastAsianFont>() is { } eastAsian) eastAsian.Typeface = family;
        if (fonts.GetFirstChild<A.ComplexScriptFont>() is { } complex) complex.Typeface = family;
    }
}
