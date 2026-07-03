using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite {
    public class Markdown_VisualTheme_Tests {
        [Theory]
        [InlineData("#abc", "aabbcc")]
        [InlineData("aabbcc", "aabbcc")]
        [InlineData("#AABBCCDD", "aabbcc")]
        [InlineData("CornflowerBlue", "6495ed")]
        [InlineData("DarkGrey", "a9a9a9")]
        public void MarkdownColor_Parses_Hex_And_Named_Colors(string value, string expectedRgbHex) {
            MarkdownColor color = MarkdownColor.Parse(value);

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

            Assert.Equal("2e8b57", theme.Palette.Accent.ToRgbHex());
            Assert.Equal("064e3b", theme.Palette.Heading.ToRgbHex());
            Assert.Equal(1.2, theme.Table.BorderWidth);
            Assert.Equal(11, theme.Table.CellPaddingX);
            Assert.False(theme.Table.UseRowStripes);
        }

        [Fact]
        public void MarkdownVisualTheme_Default_Uses_WordLike_Document_Profile() {
            MarkdownVisualTheme theme = MarkdownVisualTheme.Default();

            Assert.Equal(MarkdownVisualThemeKind.WordLike, theme.Kind);
            Assert.Equal(HtmlStyle.Word, theme.HtmlStyle);
            Assert.Equal("1f2937", theme.Palette.Text.ToRgbHex());
            Assert.Equal("111827", theme.Palette.Heading.ToRgbHex());
        }

        [Fact]
        public void MarkdownVisualTheme_Presets_List_Stable_Theme_Choices() {
            Assert.Collection(
                MarkdownVisualTheme.Presets,
                preset => Assert.Equal(MarkdownVisualThemeKind.Plain, preset.Kind),
                preset => Assert.Equal(MarkdownVisualThemeKind.WordLike, preset.Kind),
                preset => Assert.Equal(MarkdownVisualThemeKind.TechnicalDocument, preset.Kind),
                preset => Assert.Equal(MarkdownVisualThemeKind.GitHubLike, preset.Kind),
                preset => Assert.Equal(MarkdownVisualThemeKind.Compact, preset.Kind),
                preset => Assert.Equal(MarkdownVisualThemeKind.Report, preset.Kind));

            Assert.Contains(MarkdownColorSchemeKind.Emerald, MarkdownVisualTheme.ColorSchemes);
            Assert.Contains(MarkdownColorSchemeKind.Slate, MarkdownVisualTheme.ColorSchemes);
        }

        [Fact]
        public void MarkdownVisualTheme_Creates_Named_Preset_With_Color_Scheme() {
            Assert.True(MarkdownVisualTheme.TryCreate("business report", MarkdownColorSchemeKind.Emerald, out MarkdownVisualTheme? theme));

            Assert.NotNull(theme);
            Assert.Equal(MarkdownVisualThemeKind.Report, theme!.Kind);
            Assert.Equal("059669", theme.Palette.Accent.ToRgbHex());
            Assert.Equal("064e3b", theme.Palette.Heading.ToRgbHex());
            Assert.Equal(HtmlStyle.Word, theme.HtmlStyle);
        }
    }
}
