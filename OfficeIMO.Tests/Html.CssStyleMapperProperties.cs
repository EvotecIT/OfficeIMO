using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word.Html;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Html {
        [Fact]
        public void CssStyleMapper_ParsesMarginsAndDecorations() {
            var css = "margin:10pt 20pt;text-decoration:underline line-through;background-color:#112233;line-height:150%;white-space:pre-wrap";

            var properties = CssStyleMapper.ParseStyles(css);

            Assert.Equal(200, properties.MarginTop);
            Assert.Equal(200, properties.MarginBottom);
            Assert.Equal(400, properties.MarginLeft);
            Assert.Equal(400, properties.MarginRight);
            Assert.True(properties.Underline);
            Assert.True(properties.Strike);
            Assert.Equal("112233", properties.BackgroundColor);
            Assert.Equal(360, properties.LineHeight);
            Assert.Equal(LineSpacingRuleValues.Auto, properties.LineHeightRule);
            Assert.Equal(WhiteSpaceMode.PreWrap, properties.WhiteSpace);
        }

        [Fact]
        public void CssStyleMapper_ParsesIndividualMargins() {
            var css = "margin-left:5pt;margin-right:7pt;margin-top:9pt;margin-bottom:11pt";

            var properties = CssStyleMapper.ParseStyles(css);

            Assert.Equal(100, properties.MarginLeft);
            Assert.Equal(140, properties.MarginRight);
            Assert.Equal(180, properties.MarginTop);
            Assert.Equal(220, properties.MarginBottom);
        }
    }
}
