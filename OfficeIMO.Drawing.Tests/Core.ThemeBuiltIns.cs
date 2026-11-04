using System.Globalization;
using System.Reflection;
using System.Text.Json;
using OfficeIMO;
using Xunit;

namespace OfficeIMO.Drawing.Tests;

public class OfficeThemeBuiltInsTests {
    private static JsonElement[] LoadThemes() {
        using JsonDocument document = JsonDocument.Parse(File.ReadAllText(Path.Combine(AppContext.BaseDirectory, "TestAssets", "themes.json")));
        return document.RootElement.GetProperty("themes").EnumerateArray().Select(t => t.Clone()).ToArray();
    }

    public static IEnumerable<object[]> ThemeNames() => LoadThemes().Select(t => new object[] { t.GetProperty("name").GetString()! });

    private static OfficeTheme BuiltIn(string name) =>
        (OfficeTheme)typeof(OfficeTheme).GetProperty(name, BindingFlags.Public | BindingFlags.Static)!.GetValue(null)!;

    [Fact]
    public void EveryDesignThemeIsABuiltInAndTheDefaultMatches() {
        JsonElement[] themes = LoadThemes();
        Assert.Equal(new[] { "Modern", "Classic", "Minimal", "Boardroom", "Dark" }, themes.Select(t => t.GetProperty("name").GetString()));
        string defaultName = themes.Single(t => t.TryGetProperty("default", out JsonElement d) && d.GetBoolean()).GetProperty("name").GetString()!;
        Assert.Same(BuiltIn(defaultName), OfficeTheme.Default);
    }

    [Theory]
    [MemberData(nameof(ThemeNames))]
    public void BuiltInCarriesEveryDesignValue(string name) {
        JsonElement t = LoadThemes().Single(x => x.GetProperty("name").GetString() == name);
        OfficeTheme theme = BuiltIn(name);
        Assert.Equal(name, theme.Name);

        JsonElement p = t.GetProperty("palette");
        OfficeThemePalette palette = theme.Palette;
        Color(p, "background", palette.Background); Color(p, "text", palette.Text); Color(p, "textMuted", palette.TextMuted);
        Color(p, "accent1", palette.Accent1); Color(p, "accent2", palette.Accent2); Color(p, "accent3", palette.Accent3);
        Color(p, "accent4", palette.Accent4); Color(p, "accent5", palette.Accent5); Color(p, "accent6", palette.Accent6);
        Color(p, "success", palette.Success); Color(p, "warning", palette.Warning); Color(p, "danger", palette.Danger);

        JsonElement f = t.GetProperty("fonts");
        Font(f.GetProperty("heading"), theme.Fonts.Heading);
        Font(f.GetProperty("body"), theme.Fonts.Body);

        JsonElement s = t.GetProperty("typeScale");
        OfficeThemeTypeScale scale = theme.TypeScale;
        Points(s, "title", scale.Title); Points(s, "heading1", scale.Heading1); Points(s, "heading2", scale.Heading2);
        Points(s, "heading3", scale.Heading3); Points(s, "heading4", scale.Heading4); Points(s, "heading5", scale.Heading5);
        Points(s, "heading6", scale.Heading6); Points(s, "body", scale.Body); Points(s, "caption", scale.Caption);

        JsonElement tb = t.GetProperty("tableStyle");
        JsonElement ext = t.GetProperty("extensions");
        JsonElement tx = ext.GetProperty("table");
        OfficeThemeTableStyle table = theme.TableStyle;
        Color(tb, "headerFill", table.HeaderFill); Color(tb, "headerText", table.HeaderText);
        Assert.Equal(tb.GetProperty("banding").GetBoolean(), table.Banding);
        Color(tb, "bandFill", table.BandFill); Color(tb, "borderColor", table.BorderColor);
        Points(tb, "borderWidth", table.BorderWidth); Points(tb, "cellPadding", table.CellPadding);
        Assert.Equal(tx.GetProperty("borderMode").GetString(), table.BorderMode.ToString(), ignoreCase: true);
        Rule(tx.GetProperty("headerRule"), table.HeaderRule);
        Assert.Equal(tx.GetProperty("headerBold").GetBoolean(), table.HeaderBold);
        Assert.Equal(tx.GetProperty("numbersAlign").GetString(), table.NumberAlignment.ToString(), ignoreCase: true);

        JsonElement ch = t.GetProperty("chartStyle");
        JsonElement cx = ext.GetProperty("chart");
        OfficeThemeChartStyle chart = theme.ChartStyle;
        Assert.Equal(ch.GetProperty("seriesPalette").EnumerateArray().Select(c => c.GetString()), chart.SeriesPalette.Select(c => c.ToString()));
        Font(ch.GetProperty("axisFont"), chart.AxisFont);
        Font(ch.GetProperty("titleFont"), chart.TitleFont);
        Points(ch, "axisFontSize", chart.AxisFontSize); Points(ch, "titleFontSize", chart.TitleFontSize);
        Color(ch, "gridlineColor", chart.GridlineColor);
        Assert.Equal(ch.GetProperty("showDataLabels").GetBoolean(), chart.ShowDataLabels);
        Assert.Equal(ch.GetProperty("showLegend").GetBoolean(), chart.ShowLegend);
        Color(cx, "axisLine", chart.AxisLineColor); Color(cx, "axisText", chart.AxisTextColor);
        Points(cx, "gridlineWidth", chart.GridlineWidth); Points(cx, "lineWidth", chart.LineWidth);
        Assert.Equal(cx.GetProperty("barGapPercent").GetInt32(), chart.BarGapPercent);
        Assert.Equal(cx.GetProperty("markers").GetString(), chart.Markers.ToString(), ignoreCase: true);

        JsonElement sp = t.GetProperty("spacing");
        OfficeThemeSpacing spacing = theme.Spacing;
        Points(sp, "paragraphBefore", spacing.ParagraphBefore); Points(sp, "paragraphAfter", spacing.ParagraphAfter);
        Points(sp, "listIndent", spacing.ListIndent); Points(sp, "slideMarginTop", spacing.SlideMarginTop);
        Points(sp, "slideMarginRight", spacing.SlideMarginRight); Points(sp, "slideMarginBottom", spacing.SlideMarginBottom);
        Points(sp, "slideMarginLeft", spacing.SlideMarginLeft);

        OfficeThemeTextStyle text = theme.TextStyle;
        Assert.Equal(ext.GetProperty("lineSpacing").GetDouble(), text.LineSpacing, 6);
        Color(ext, "titleColor", text.TitleColor); Color(ext, "subtitleColor", text.SubtitleColor);
        Color(ext, "headingColor", text.HeadingColor); Color(ext, "link", text.LinkColor);
        Assert.Equal(ext.GetProperty("headingBold").GetBoolean(), text.HeadingBold);
        Rule(ext.TryGetProperty("titleRule", out JsonElement titleRule) ? titleRule : default, text.TitleRule);
        JsonElement hs = ext.GetProperty("headingSpacing");
        string[] levels = { "title", "heading1", "heading2", "heading3" };
        for (int level = 0; level < levels.Length; level++) {
            JsonElement pair = hs.GetProperty(levels[level]);
            Assert.Equal(pair[0].GetDouble(), text.GetHeadingSpacing(level).Before.ToPoints(), 6);
            Assert.Equal(pair[1].GetDouble(), text.GetHeadingSpacing(level).After.ToPoints(), 6);
        }
        Assert.Same(text.Heading3Spacing, text.GetHeadingSpacing(6));

        JsonElement de = t.GetProperty("derived");
        OfficeThemeSurfaces surfaces = theme.Surfaces;
        Color(ext, "surface", surfaces.Surface); Color(ext, "rule", surfaces.Rule);
        Color(de, "accent1Tint", surfaces.Accent1Tint); Color(de, "successFill", surfaces.SuccessFill);
        Color(de, "warningFill", surfaces.WarningFill); Color(de, "dangerFill", surfaces.DangerFill);

        JsonElement pg = ext.GetProperty("page");
        Points(pg, "marginTop", theme.Page.MarginTop); Points(pg, "marginRight", theme.Page.MarginRight);
        Points(pg, "marginBottom", theme.Page.MarginBottom); Points(pg, "marginLeft", theme.Page.MarginLeft);

        JsonElement ft = ext.GetProperty("footer");
        Points(ft, "size", theme.Footer.FontSize);
        Color(ft, "color", theme.Footer.Color);
        JsonElement footerRule = ft.GetProperty("rule");
        if (footerRule.ValueKind == JsonValueKind.Null) Assert.False(theme.Footer.Rule.IsVisible);
        else Assert.Equal(footerRule.GetString(), theme.Footer.Rule.Color.ToString().TrimStart('#'));
    }

    [Theory]
    [MemberData(nameof(ThemeNames))]
    public void StatusAndAccentTintsFollowTheDerivationRule(string name) {
        OfficeTheme theme = BuiltIn(name);
        OfficeThemeSurfaces derived = OfficeThemeSurfaces.FromPalette(theme.Palette, theme.Surfaces.Surface, theme.Surfaces.Rule);
        Close(theme.Surfaces.Accent1Tint, derived.Accent1Tint);
        Close(theme.Surfaces.SuccessFill, derived.SuccessFill);
        Close(theme.Surfaces.WarningFill, derived.WarningFill);
        Close(theme.Surfaces.DangerFill, derived.DangerFill);
        Assert.Equal(name == "Dark", theme.IsDark);
    }

    [Fact]
    public void SevenComponentConstructorDerivesTheRemainingComponents() {
        OfficeTheme modern = OfficeTheme.Modern;
        var palette = modern.Palette.With(accent1: OfficeColor.ParseHex("000080"));
        var custom = new OfficeTheme("Company", palette, modern.Fonts, modern.TypeScale, modern.TableStyle, modern.ChartStyle, modern.Spacing);
        Assert.Same(modern.TextStyle, custom.TextStyle);
        Assert.Same(modern.Page, custom.Page);
        Assert.Same(modern.Footer, custom.Footer);
        Assert.Equal(modern.TableStyle.BandFill, custom.Surfaces.Surface);
        Assert.Equal(modern.TableStyle.BorderColor, custom.Surfaces.Rule);
        Assert.Equal(OfficeColor.ParseHex("E6E6F2"), custom.Surfaces.Accent1Tint);
        Assert.Equal(modern.Surfaces.SuccessFill, custom.Surfaces.SuccessFill);

        var legacyTable = new OfficeThemeTableStyle(OfficeColor.Accent1, OfficeColor.Background, false, OfficeColor.White, OfficeColor.Gray, OfficeLength.Points(1), OfficeLength.Points(4));
        Assert.Equal(OfficeThemeTableBorderMode.Horizontal, legacyTable.BorderMode);
        Assert.False(legacyTable.HeaderRule.IsVisible);
        Assert.True(legacyTable.HeaderBold);
        Assert.Equal(OfficeAlignment.Right, legacyTable.NumberAlignment);
        var withRule = legacyTable.With(headerRule: new OfficeThemeRule(OfficeColor.Text, OfficeLength.Points(1)), borderMode: OfficeThemeTableBorderMode.Grid);
        Assert.True(withRule.HeaderRule.IsVisible);
        Assert.Equal(OfficeThemeTableBorderMode.Grid, withRule.BorderMode);
        Assert.Equal(legacyTable.CellPadding, withRule.CellPadding);
        Assert.False(withRule.With(headerRule: OfficeThemeRule.None).HeaderRule.IsVisible);

        OfficeThemeChartStyle chart = modern.ChartStyle.With(lineWidth: OfficeLength.Points(3), markers: OfficeChartMarkerShape.Circle);
        Assert.Equal(3, chart.LineWidth.ToPoints());
        Assert.Equal(OfficeChartMarkerShape.Circle, chart.Markers);
        Assert.Equal(modern.ChartStyle.BarGapPercent, chart.BarGapPercent);
        Assert.Throws<ArgumentOutOfRangeException>(() => modern.ChartStyle.With(barGapPercent: -1));
        Assert.Throws<ArgumentOutOfRangeException>(() => modern.TextStyle.With(lineSpacing: 0));
        Assert.Throws<ArgumentOutOfRangeException>(() => modern.TextStyle.GetHeadingSpacing(7));

        OfficeTheme replaced = modern.WithPage(modern.Page.With(marginLeft: OfficeLength.Points(90)))
            .WithFooter(modern.Footer.With(fontSize: OfficeLength.Points(9)))
            .WithTextStyle(modern.TextStyle.With(headingBold: false))
            .WithSurfaces(modern.Surfaces.With(surface: OfficeColor.White));
        Assert.Equal(90, replaced.Page.MarginLeft.ToPoints());
        Assert.Equal(64, modern.Page.MarginLeft.ToPoints());
        Assert.Equal(9, replaced.Footer.FontSize.ToPoints());
        Assert.False(replaced.TextStyle.HeadingBold);
        Assert.Equal(OfficeColor.White, replaced.Surfaces.Surface);
        Assert.Same(modern.Palette, replaced.Palette);
    }

    private static void Color(JsonElement parent, string property, OfficeColor actual) {
        string expected = parent.GetProperty(property).GetString()!;
        Assert.Equal(expected, actual.IsThemeRole ? actual.ToString() : actual.ToString().TrimStart('#'));
    }

    private static void Points(JsonElement parent, string property, OfficeLength actual) =>
        Assert.Equal(parent.GetProperty(property).GetDouble(), actual.ToPoints(), 6);

    private static void Font(JsonElement expected, OfficeThemeFont actual) {
        Assert.Equal(expected.GetProperty("family").GetString(), actual.Family);
        Assert.Equal(expected.GetProperty("fallbacks").EnumerateArray().Select(x => x.GetString()), actual.Fallbacks);
    }

    private static void Rule(JsonElement expected, OfficeThemeRule actual) {
        if (expected.ValueKind is JsonValueKind.Undefined or JsonValueKind.Null) {
            Assert.False(actual.IsVisible);
            return;
        }
        Color(expected, "color", actual.Color);
        Points(expected, "width", actual.Width);
    }

    private static void Close(OfficeColor expected, OfficeColor actual) {
        Assert.InRange(Math.Abs(expected.R - actual.R), 0, 1);
        Assert.InRange(Math.Abs(expected.G - actual.G), 0, 1);
        Assert.InRange(Math.Abs(expected.B - actual.B), 0, 1);
    }
}
