using System.Globalization;
using System.Xml.Linq;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests {

    public sealed partial class HtmlRenderingTests {
        [Theory]
        [InlineData(10D, 0D, 35.41796875D, 20D, false)]
        [InlineData(10D, 0D, 35.41796875D, 20D, true)]
        [InlineData(10D, 10D, 38.58203125D, 26.58203125D, false)]
        [InlineData(10D, 10D, 38.58203125D, 26.58203125D, true)]
        [InlineData(10D, -10D, 45.41796875D, 20D, false)]
        [InlineData(10D, -10D, 45.41796875D, 20D, true)]
        [InlineData(20D, 10D, 42D, 30D, false)]
        [InlineData(20D, 10D, 42D, 30D, true)]
        public void HtmlTextLine_FaceLeadingAndDisplacementAdvanceFollowingContentTogether(
            double size, double shift, double expectedHeight, double expectedBaseline, bool floating) {
            // The supplied face has hhea ascent1950/descent550 at2048 units/em.
            // A20px/32px strut has A22.8359375/D9.1640625; a10px/32px
            // run has A19.41796875/D12.58203125. These authored extents include
            // half-leading and the explicit vertical-align shift, rather than em ascent.
            string floatBox = floating ? "<i style='float:left;width:20px;height:200px'></i>" : "";
            string html = "<style>body{margin:0;font:20px/32px Pinned}p{margin:0}</style>"
                + "<div style='width:120px'>" + floatBox + "BASE <span style='font-size:"
                + size.ToString(CultureInfo.InvariantCulture) + "px;vertical-align:"
                + shift.ToString(CultureInfo.InvariantCulture) + "px'>X</span><br>TAIL</div><p>FOLLOWING</p>";
            HtmlConversionDocument document = HtmlConversionDocument.Parse(html);
            var options = new HtmlRenderOptions { ViewportWidth = 400D, Margins = HtmlRenderMargins.All(0D) };
            options.Fonts.Add("Pinned", File.ReadAllBytes(Path.Combine(RepositoryTestPaths.Find(),
                "OfficeIMO.TestAssets", "Fonts", "OfficeIMOBaselineSans-Regular.ttf")));
            HtmlRenderText[] text = EnumerateRenderVisuals(HtmlRenderTestDriver.Render(document, options).Pages[0].Scene)
                .OfType<HtmlRenderText>().ToArray();
            HtmlRenderText tail = Assert.Single(text, item => item.Text == "TAIL");
            Assert.Equal(expectedHeight, tail.LayoutY, 6);
            HtmlRenderText following = Assert.Single(text, item => item.Text == "FOLLOWING");
            Assert.True(following.LayoutY >= tail.LayoutY + 32D - 0.000001D);

            XElement[] svg = XDocument.Parse(document.ToSvg(options)).Descendants()
                .Where(node => node.Name.LocalName == "text").ToArray();
            double baseline = Y(Assert.Single(svg, item => item.Value.Trim() == "BASE"));
            Assert.Equal(expectedBaseline, baseline, 3);
            Assert.Equal(baseline - shift, Y(Assert.Single(svg, item => item.Value == "X")), 3);

            static double Y(XElement element) => double.Parse(element.Attribute("y")!.Value, CultureInfo.InvariantCulture);
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
