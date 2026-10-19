using OfficeIMO.Drawing;
using OfficeIMO.Html;
using AngleSharp.Html.Parser;
using Xunit;

namespace OfficeIMO.Tests {
    public sealed partial class HtmlRenderingTests {
        [Theory]
        [InlineData("<em>X</em>", OfficeFontSlant.Italic, 400, 0D)]
        [InlineData("<i>X</i>", OfficeFontSlant.Italic, 400, 0D)]
        [InlineData("<cite>X</cite>", OfficeFontSlant.Italic, 400, 0D)]
        [InlineData("<dfn>X</dfn>", OfficeFontSlant.Italic, 400, 0D)]
        [InlineData("<var>X</var>", OfficeFontSlant.Italic, 400, 0D)]
        [InlineData("<address>X</address>", OfficeFontSlant.Italic, 400, 0D)]
        [InlineData("<em><span>X</span></em>", OfficeFontSlant.Italic, 400, 0D)]
        [InlineData("<strong><em>X</em></strong>", OfficeFontSlant.Italic, 700, 0D)]
        [InlineData("<em><strong>X</strong></em>", OfficeFontSlant.Italic, 700, 0D)]
        [InlineData("<em style='font-style:normal'>X</em>", OfficeFontSlant.Normal, 400, 0D)]
        [InlineData("<em style='font-style:initial'>X</em>", OfficeFontSlant.Normal, 400, 0D)]
        [InlineData("<em style='font-style:inherit'>X</em>", OfficeFontSlant.Normal, 400, 0D)]
        [InlineData("<em style='font-style:unset'>X</em>", OfficeFontSlant.Normal, 400, 0D)]
        [InlineData("<em style='font-style:revert'>X</em>", OfficeFontSlant.Italic, 400, 0D)]
        [InlineData("<em style='font:16px Arial'>X</em>", OfficeFontSlant.Normal, 400, 0D)]
        [InlineData("<em style='font:inherit'>X</em>", OfficeFontSlant.Normal, 400, 0D)]
        [InlineData("<style>@layer base,over;@layer base{em{font-style:normal}}@layer over{em{font-style:revert-layer}}</style><em>X</em>", OfficeFontSlant.Normal, 400, 0D)]
        [InlineData("<style>@layer over{em{font-style:revert-layer}}</style><em>X</em>", OfficeFontSlant.Italic, 400, 0D)]
        [InlineData("<em><span style='font-style:inherit'>X</span></em>", OfficeFontSlant.Italic, 400, 0D)]
        [InlineData("<em><span style='font-style:unset'>X</span></em>", OfficeFontSlant.Italic, 400, 0D)]
        [InlineData("<em><span style='font-style:initial'>X</span></em>", OfficeFontSlant.Normal, 400, 0D)]
        [InlineData("<span style='font-style:oblique 20deg'><em>X</em></span>", OfficeFontSlant.Italic, 400, 0D)]
        [InlineData("<span style='font-style:oblique 20deg'><em style='font-style:inherit'>X</em></span>", OfficeFontSlant.Oblique, 400, 20D)]
        [InlineData("<span style='font-style:italic'><span style='font-style:normal'>X</span></span>", OfficeFontSlant.Normal, 400, 0D)]
        public void HtmlText_SemanticItalicDefaultsRespectAuthoredStyleAndInheritance(
            string markup, OfficeFontSlant expectedSlant, int expectedWeight, double expectedAngle) {
            HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(
                "<body style='margin:0;font:16px/24px Arial'>" + markup + "</body>",
                new HtmlRenderOptions {
                    ViewportWidth = 200D,
                    ViewportHeight = 80D,
                    Margins = HtmlRenderMargins.All(0D),
                    AllowSystemFontFallback = false
                });
            HtmlRenderText text = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderText>());

            Assert.Equal("X", text.Text);
            Assert.Equal(expectedSlant, text.FontDescriptor.Slant);
            Assert.Equal(expectedWeight, text.FontDescriptor.Weight);
            Assert.Equal(expectedAngle, text.FontDescriptor.ObliqueAngleDegrees, 6);
        }

        [Theory]
        [InlineData("<strong>X</strong>", 300, 400)]
        [InlineData("<strong>X</strong>", 900, 900)]
        [InlineData("<b>X</b>", 300, 400)]
        [InlineData("<b>X</b>", 900, 900)]
        [InlineData("<strong style='font-weight:inherit'>X</strong>", 300, 300)]
        [InlineData("<b style='font-weight:unset'>X</b>", 900, 900)]
        [InlineData("<strong style='font-weight:initial'>X</strong>", 900, 400)]
        [InlineData("<strong><span style='font-weight:inherit'>X</span></strong>", 300, 400)]
        [InlineData("<strong><span style='font-weight:unset'>X</span></strong>", 900, 900)]
        [InlineData("<strong><span style='font-weight:initial'>X</span></strong>", 900, 400)]
        [InlineData("<strong><strong>X</strong></strong>", 400, 900)]
        [InlineData("<h1 style='font-weight:initial'>X</h1>", 900, 400)]
        [InlineData("<table><tr><th style='font-weight:initial'>X</th></tr></table>", 900, 400)]
        public void HtmlText_SemanticBoldDefaultsRespectRelativeWeightAndInheritance(
            string markup, int inheritedWeight, int expectedWeight) {
            HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(
                "<body style='margin:0;font:16px/24px Arial;font-weight:" + inheritedWeight + "'>" + markup + "</body>",
                new HtmlRenderOptions {
                    ViewportWidth = 200D,
                    ViewportHeight = 80D,
                    Margins = HtmlRenderMargins.All(0D),
                    AllowSystemFontFallback = false
                });
            HtmlRenderText text = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderText>());

            Assert.Equal("X", text.Text);
            Assert.Equal(expectedWeight, text.FontDescriptor.Weight);
        }

        [Fact]
        public void HtmlComputedFontDefaultsReachExplicitInheritanceAndRelativeWeightConsumers() {
            var document = new HtmlParser().ParseDocument("<body style='font:16px Arial;font-weight:300'>"
                + "<em id='emphasis'><span id='inherited' style='font-style:inherit'>X</span></em>"
                + "<strong id='strong'><span id='relative' style='font-weight:bolder'>X</span></strong></body>");
            var styles = HtmlComputedStyleEngine.Compute(document);

            Assert.Equal("italic", styles[document.QuerySelector("#emphasis")!].GetValue("font-style"));
            Assert.Equal("italic", styles[document.QuerySelector("#inherited")!].GetValue("font-style"));
            Assert.Equal("400", styles[document.QuerySelector("#strong")!].GetValue("font-weight"));
            Assert.Equal("700", styles[document.QuerySelector("#relative")!].GetValue("font-weight"));
        }
    }
}
