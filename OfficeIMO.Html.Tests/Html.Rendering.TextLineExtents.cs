using System.Globalization;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests {

    public sealed partial class HtmlRenderingTests {
        [Theory]
        [InlineData(10D, 0D, 36D, 23D, false)]
        [InlineData(10D, 0D, 36D, 23D, true)]
        [InlineData(10D, 10D, 38D, 29D, false)]
        [InlineData(10D, 10D, 38D, 29D, true)]
        [InlineData(10D, -10D, 46D, 23D, false)]
        [InlineData(10D, -10D, 46D, 23D, true)]
        [InlineData(20D, 10D, 42D, 33D, false)]
        [InlineData(20D, 10D, 42D, 33D, true)]
        public void HtmlTextLine_FaceLeadingAndDisplacementAdvanceFollowingContentTogether(
            double size, double shift, double expectedHeight, double expectedBaseline, bool floating) {
            // The supplied face has hhea ascent1950/descent550 at2048 units/em.
            // Paired native witnesses have a20px/32px strut A23/D9 and
            // a10px/32px run A19/D13 after CSS font/leading alignment. These
            // extents include vertical-align displacement and agree beside floats.
            string floatBox = floating ? "<i style='float:left;width:20px;height:200px'></i>" : "";
            string html = "<style>body{margin:0;font:20px/32px Pinned}p{margin:0}</style>"
                + "<div style='width:120px'>" + floatBox + "BASE <span style='font-size:"
                + size.ToString(CultureInfo.InvariantCulture) + "px;vertical-align:"
                + shift.ToString(CultureInfo.InvariantCulture) + "px'>X</span><br>TAIL</div><p>FOLLOWING</p>";
            HtmlConversionDocument document = HtmlConversionDocument.Parse(html);
            var options = new HtmlRenderOptions { ViewportWidth = 400D, Margins = HtmlRenderMargins.All(0D) };
            options.Fonts.Add("Pinned", File.ReadAllBytes(Path.Combine(RepositoryTestPaths.Find(),
                "OfficeIMO.TestAssets", "Fonts", "OfficeIMOBaselineSans-Regular.ttf")));
            HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(document, options);
            HtmlRenderText[] text = EnumerateRenderVisuals(rendered.Pages[0].Scene)
                .OfType<HtmlRenderText>().ToArray();
            HtmlRenderText tail = Assert.Single(text, item => item.Text == "TAIL");
            Assert.Equal(expectedHeight, tail.LayoutY, 6);
            HtmlRenderText following = Assert.Single(text, item => item.Text == "FOLLOWING");
            Assert.True(following.LayoutY >= tail.LayoutY + 32D - 0.000001D);

            double baseline = PaintBaseline(Assert.Single(text, item => item.Text.Trim() == "BASE"));
            Assert.Equal(expectedBaseline, baseline, 3);
            Assert.Equal(baseline - shift, PaintBaseline(Assert.Single(text, item => item.Text == "X")), 3);
            AssertPdfBackedSvgRetainsBaselineInk(rendered.Pages[0], document.ToSvg(options));
        }

        [Fact]
        public void HtmlScriptText_ZeroLineHeightDoesNotAddItsPaintDisplacementToFlow() {
            const string html = "<style>body{margin:0;font:20px/32px Helvetica}p{margin:0}"
                + "sub,sup{line-height:0}</style><p>Water H<sub>2</sub>O and m<sup>2</sup></p><p>FOLLOWING</p>";
            HtmlRenderText[] text = EnumerateRenderVisuals(HtmlRenderTestDriver.Render(html,
                new HtmlRenderOptions { Margins = HtmlRenderMargins.All(0D) }).Pages[0].Scene)
                .OfType<HtmlRenderText>().ToArray();
            Assert.Equal(32D, Assert.Single(text, item => item.Text == "FOLLOWING").LayoutY, 6);
        }
    }
}
