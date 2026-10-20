using OfficeIMO.Drawing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Theory]
        [InlineData("form='prefix'")]
        [InlineData("lspace='.5em'")]
        [InlineData("rspace='0'")]
        public void NativeWordMathRejectsAuthoredOperatorPropertiesItCannotPreserve(string attributes) {
            OfficeMathExpression expression = OfficeMathMarkup.FromMathMl(
                "<math><mrow><mi>x</mi><mo " + attributes + ">+</mo><mi>y</mi></mrow></math>");
            NotSupportedException error = Assert.Throws<NotSupportedException>(() => WordMathMarkup.ToOmml(expression));
            Assert.Contains("operator form and spacing", error.Message, StringComparison.Ordinal);
            Assert.Contains("MathML", error.Message, StringComparison.Ordinal);
        }
    }
}
