using OfficeIMO;
using Xunit;

namespace OfficeIMO.Drawing.Tests;

public class OfficeThemeTests {
    [Theory]
    [InlineData("Accent1", "accent1")]
    [InlineData("Accent2", "accent2")]
    [InlineData("Accent3", "accent3")]
    [InlineData("Accent4", "accent4")]
    [InlineData("Accent5", "accent5")]
    [InlineData("Accent6", "accent6")]
    [InlineData("Text", "dk1")]
    [InlineData("TextMuted", "dk2")]
    [InlineData("Background", "lt1")]
    [InlineData("Success", null)]
    [InlineData("Warning", null)]
    [InlineData("Danger", null)]
    public void RolesKeepTheirIdentityAndExposeNativeSlots(string name, string? slot) {
        OfficeColor color = OfficeColor.Parse(" " + name.ToLowerInvariant() + " ");
        Assert.True(color.IsThemeRole);
        Assert.Equal(slot, color.ThemeSlot);
        Assert.Equal(name, color.ToString());
        Assert.Equal(color, OfficeColor.Parse(color.ToString()));
        Assert.False(color.Resolve(OfficeTheme.Modern).IsThemeRole);
        Assert.NotEqual(color, color.Resolve(OfficeTheme.Modern));
        Assert.Equal(color.GetHashCode(), OfficeColor.Parse(name).GetHashCode());
        Assert.Throws<InvalidOperationException>(() => color.ToRgbHex());
        Assert.Throws<InvalidOperationException>(() => { _ = color.R; });
    }

    [Fact]
    public void ResolutionFollowsCustomizedPaletteAndNeverMutatesATheme() {
        OfficeTheme theme = OfficeTheme.Default;
        var palette = theme.Palette.With(accent1: OfficeColor.Red, success: OfficeColor.Lime, textMuted: OfficeColor.Gray);
        var custom = theme.WithPalette(palette).WithName("Company");
        Assert.Equal(OfficeColor.Red, OfficeColor.Accent1.Resolve(custom));
        Assert.Equal(OfficeColor.Lime, OfficeColor.Success.Resolve(custom));
        Assert.Equal(OfficeColor.Gray, OfficeColor.TextMuted.Resolve(custom));
        Assert.Equal(OfficeColor.Blue, OfficeColor.Blue.Resolve(custom));
        Assert.NotEqual(OfficeColor.Red, theme.Palette.Accent1);
        Assert.Same(theme.Fonts, custom.Fonts);
        Assert.Equal("Company", custom.Name);
        Assert.Throws<ArgumentException>(() => palette.With(accent1: OfficeColor.Accent2));
        Assert.Throws<ArgumentNullException>(() => OfficeColor.Accent1.Resolve(null!));
    }

    [Fact]
    public void ThemeComponentsSnapshotListsAndCopyOnlyReplacedValues() {
        var names = new[] { "Arial", "Liberation Sans" };
        var font = new OfficeThemeFont("Aptos", names);
        names[0] = "Changed";
        Assert.Equal("Arial", font.Fallbacks[0]);
        var colors = new[] { OfficeColor.Accent1, OfficeColor.Accent2 };
        OfficeTheme original = OfficeTheme.Modern;
        var charts = original.ChartStyle.With(seriesPalette: colors, showDataLabels: true);
        colors[0] = OfficeColor.Red;
        Assert.Equal(OfficeColor.Accent1, charts.SeriesPalette[0]);
        OfficeLength originalAfter = original.Spacing.ParagraphAfter;
        var copy = original.WithFonts(original.Fonts.With(body: font))
            .WithTypeScale(original.TypeScale.With(body: OfficeLength.Points(13)))
            .WithTableStyle(original.TableStyle.With(banding: false, cellPadding: OfficeLength.Points(8)))
            .WithChartStyle(charts).WithSpacing(original.Spacing.With(paragraphAfter: OfficeLength.Points(10)));
        Assert.Equal(13, copy.TypeScale.Body.ToPoints());
        Assert.False(copy.TableStyle.Banding);
        Assert.True(copy.ChartStyle.ShowDataLabels);
        Assert.False(original.ChartStyle.ShowDataLabels);
        Assert.True(original.TableStyle.Banding);
        Assert.Equal(originalAfter, original.Spacing.ParagraphAfter);
        Assert.Equal(10, copy.Spacing.ParagraphAfter.ToPoints());
        Assert.Same(original.Palette, copy.Palette);
        Assert.Throws<ArgumentOutOfRangeException>(() => original.TypeScale.With(body: OfficeLength.Points(-1)));
        Assert.Throws<ArgumentException>(() => charts.With(seriesPalette: Array.Empty<OfficeColor>()));
        Assert.Throws<ArgumentException>(() => font.With(family: " "));
    }

    [Fact]
    public void EveryBuiltInHasResolvableAuthoringTokens() {
        Assert.Same(OfficeTheme.Modern, OfficeTheme.Default);
        foreach (var theme in new[] { OfficeTheme.Modern, OfficeTheme.Classic, OfficeTheme.Minimal, OfficeTheme.Boardroom, OfficeTheme.Dark }) {
            Assert.True(theme.TypeScale.Body > OfficeLength.Zero);
            Assert.Equal(6, theme.ChartStyle.SeriesPalette.Count);
            Assert.NotEmpty(theme.Fonts.Body.Fallbacks);
            Assert.False(theme.TableStyle.HeaderFill.Resolve(theme).IsThemeRole);
            foreach (OfficeColorRole role in Enum.GetValues(typeof(OfficeColorRole))) {
                if (role == OfficeColorRole.None) continue;
                Assert.False(OfficeColor.Parse(role.ToString()).Resolve(theme).IsThemeRole);
            }
        }
        Assert.NotEqual(OfficeColor.Text.Resolve(OfficeTheme.Modern), OfficeColor.Text.Resolve(OfficeTheme.Dark));
    }

    [Theory]
    [InlineData("#12345678", "#12345678")]
    [InlineData("#ABCDEF", "#ABCDEF")]
    [InlineData("RebeccaPurple", "#663399")]
    public void ExistingLiteralParsingRemainsUnchanged(string text, string expected) {
        OfficeColor literal = OfficeColor.Parse(text);
        Assert.False(literal.IsThemeRole);
        Assert.Null(literal.ThemeSlot);
        Assert.Equal(expected, literal.ToString());
        Assert.Equal(literal, OfficeColor.Parse(literal.ToCssColor()));
        Assert.False(OfficeColor.TryParse("1", out _));
    }
}
