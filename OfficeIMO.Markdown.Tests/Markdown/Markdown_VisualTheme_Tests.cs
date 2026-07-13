using OfficeIMO.Drawing;
using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite {
    public class Markdown_VisualTheme_Tests {
        [Theory]
        [InlineData("#abc", "AABBCC")]
        [InlineData("aabbcc", "AABBCC")]
        [InlineData("#AABBCCDD", "AABBCC")]
        [InlineData("CornflowerBlue", "6495ED")]
        [InlineData("DarkGrey", "A9A9A9")]
        public void OfficeColor_Parses_MarkdownTheme_Hex_And_Named_Colors(string value, string expectedRgbHex) {
            OfficeColor color = OfficeColor.Parse(value);

            Assert.Equal(expectedRgbHex, color.ToRgbHex());
        }

        [Fact]
        public void MarkdownVisualTheme_Customizes_Colors_And_Table_Surface() {
            MarkdownVisualTheme theme = MarkdownVisualTheme.Report()
                .WithColorScheme(MarkdownColorSchemeKind.Emerald)
                .WithColors(accent: "SeaGreen", heading: "#064e3b");

            theme.Table.BorderWidth = 1.2;
            theme.Table.CellPaddingX = 11;
            theme.Table.UseRowStripes = false;

            Assert.Equal("2E8B57", theme.Palette.Accent.ToRgbHex());
            Assert.Equal("064E3B", theme.Palette.Heading.ToRgbHex());
            Assert.Equal(1.2, theme.Table.BorderWidth);
            Assert.Equal(11, theme.Table.CellPaddingX);
            Assert.False(theme.Table.UseRowStripes);
        }

        [Fact]
        public void MarkdownVisualTheme_Default_Uses_WordLike_Document_Profile() {
            MarkdownVisualTheme theme = MarkdownVisualTheme.Default();

            Assert.Equal(OfficeVisualThemeKind.WordLike, theme.Kind);
            Assert.Equal(HtmlStyle.Word, theme.HtmlStyle);
            Assert.Equal("1F2937", theme.Palette.Text.ToRgbHex());
            Assert.Equal("111827", theme.Palette.Heading.ToRgbHex());
        }

        [Fact]
        public void MarkdownVisualTheme_Presets_List_Stable_Theme_Choices() {
            Assert.Collection(
                MarkdownVisualTheme.Presets,
                preset => Assert.Equal(OfficeVisualThemeKind.Plain, preset.Kind),
                preset => Assert.Equal(OfficeVisualThemeKind.WordLike, preset.Kind),
                preset => Assert.Equal(OfficeVisualThemeKind.TechnicalDocument, preset.Kind),
                preset => Assert.Equal(OfficeVisualThemeKind.GitHubLike, preset.Kind),
                preset => Assert.Equal(OfficeVisualThemeKind.Compact, preset.Kind),
                preset => Assert.Equal(OfficeVisualThemeKind.Report, preset.Kind));

            Assert.Contains(MarkdownColorSchemeKind.Emerald, MarkdownVisualTheme.ColorSchemes);
            Assert.Contains(MarkdownColorSchemeKind.Slate, MarkdownVisualTheme.ColorSchemes);
        }

        [Fact]
        public void MarkdownVisualTheme_Creates_Named_Preset_With_Color_Scheme() {
            Assert.True(MarkdownVisualTheme.TryCreate("business report", MarkdownColorSchemeKind.Emerald, out MarkdownVisualTheme? theme));

            Assert.NotNull(theme);
            Assert.Equal(OfficeVisualThemeKind.Report, theme!.Kind);
            Assert.Equal("059669", theme.Palette.Accent.ToRgbHex());
            Assert.Equal("064E3B", theme.Palette.Heading.ToRgbHex());
            Assert.Equal(HtmlStyle.Word, theme.HtmlStyle);
        }
    }
}
