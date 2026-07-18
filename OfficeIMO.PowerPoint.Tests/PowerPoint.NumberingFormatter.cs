using OfficeIMO.PowerPoint;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.Tests {
    public partial class PowerPointTests {
        [Fact]
        public void NumberingFormatter_PreservesScriptAndSymbolSchemes() {
            Assert.Equal("๓.", PowerPointNumberingFormatter.FormatMarker(3,
                A.TextAutoNumberSchemeValues.ThaiNumberPeriod));
            Assert.Equal("३)", PowerPointNumberingFormatter.FormatMarker(3,
                A.TextAutoNumberSchemeValues
                    .HindiNumberParenthesisRight));
            Assert.Equal("３．", PowerPointNumberingFormatter.FormatMarker(3,
                A.TextAutoNumberSchemeValues.ArabicDoubleBytePeriod));
            Assert.Equal("三.", PowerPointNumberingFormatter.FormatMarker(3,
                A.TextAutoNumberSchemeValues
                    .EastAsianSimplifiedChinesePeriod));
            Assert.Equal("❸", PowerPointNumberingFormatter.FormatMarker(3,
                A.TextAutoNumberSchemeValues
                    .CircleNumberWingdingsBlackPlain));
            Assert.Equal("③", PowerPointNumberingFormatter.FormatMarker(3,
                A.TextAutoNumberSchemeValues
                    .CircleNumberDoubleBytePlain));
            Assert.Equal("ג-", PowerPointNumberingFormatter.FormatMarker(3,
                A.TextAutoNumberSchemeValues.Hebrew2Minus));
        }
    }
}
