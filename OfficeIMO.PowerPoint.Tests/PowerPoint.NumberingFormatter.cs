using OfficeIMO.PowerPoint;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.Tests {
    public partial class PowerPointTests {
        [Fact]
        public void NumberingFormatter_PreservesScriptAndSymbolSchemes() {
            Assert.Equal("๓.", PowerPointNumberingFormatter.FormatMarker(3,
                PowerPointNumberingScheme.ThaiNumberPeriod));
            Assert.Equal("३)", PowerPointNumberingFormatter.FormatMarker(3,
                PowerPointNumberingScheme
                    .HindiNumberParenthesisRight));
            Assert.Equal("３．", PowerPointNumberingFormatter.FormatMarker(3,
                PowerPointNumberingScheme.ArabicDoubleBytePeriod));
            Assert.Equal("三.", PowerPointNumberingFormatter.FormatMarker(3,
                PowerPointNumberingScheme
                    .EastAsianSimplifiedChinesePeriod));
            Assert.Equal("❸", PowerPointNumberingFormatter.FormatMarker(3,
                PowerPointNumberingScheme
                    .CircleNumberWingdingsBlackPlain));
            Assert.Equal("③", PowerPointNumberingFormatter.FormatMarker(3,
                PowerPointNumberingScheme
                    .CircleNumberDoubleBytePlain));
            Assert.Equal("ג-", PowerPointNumberingFormatter.FormatMarker(3,
                PowerPointNumberingScheme.Hebrew2Minus));
        }
    }
}
