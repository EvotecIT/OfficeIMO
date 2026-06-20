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
    }
}
