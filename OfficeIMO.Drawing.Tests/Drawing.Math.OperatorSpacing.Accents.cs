using System.Xml.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingMathOperatorAccentMetadataTests {
    [Theory]
    [InlineData("accent", "", "^")]
    [InlineData("overbar", "", "¯")]
    [InlineData("underbar", "", "_")]
    [InlineData("accent", "mrow", "^")]
    [InlineData("overbar", "mstyle", "¯")]
    public void AccentOperatorPresentationSurvivesCanonicalizationAndTextTransformation(
        string kind, string wrapper, string symbol) {
        const string attributes = "form='POSTFIX' lspace='.5em' rspace='25%'";
        string source = AccentSource(kind, wrapper, symbol, attributes);
        OfficeMathExpression expression = OfficeMathMarkup.FromMathMl(source);
        OfficeMathExpression unqualified = OfficeMathMarkup.FromMathMl(AccentSource(kind, wrapper, symbol, string.Empty));

        Assert.NotEqual(unqualified, expression);
        Assert.Equal(unqualified.ToPlainText(), expression.ToPlainText());
        AssertPortableOperator(expression);
        OfficeMathExpression transformed = expression.TransformTextCase(OfficeTextCase.Uppercase);
        Assert.Equal(expression.ToPlainText().Replace("x", "X"), transformed.ToPlainText());
        AssertPortableOperator(transformed);
    }

    [Fact]
    public void ComplexAccentOperandWithMultipleMetadataOwnersFailsExplicitly() {
        const string source = "<math><mover accent='true'><mi>x</mi><mrow>"
            + "<mo lspace='.5em'>^</mo><mo rspace='.25em'>~</mo>"
            + "</mrow></mover></math>";

        OfficeMathParseException error = Assert.Throws<OfficeMathParseException>(() => OfficeMathMarkup.FromMathMl(source));

        Assert.Equal("DRAWING_MATH_OPERATOR_PRESENTATION", error.Code);
    }

    [Fact]
    public void ComplexAccentOperandWithoutMetadataKeepsItsLegacyPortableProjection() {
        OfficeMathExpression expression = OfficeMathMarkup.FromMathMl(
            "<math><mover accent='true'><mi>x</mi><mrow><mo>^</mo><mo>~</mo></mrow></mover></math>");

        Assert.Equal("x^~", expression.ToPlainText());
        Assert.Equal(expression, OfficeMathMarkup.FromMathMl(OfficeMathMarkup.ToMathMl(expression)));
        Assert.Equal("X^~", expression.TransformTextCase(OfficeTextCase.Uppercase).ToPlainText());
    }

    private static void AssertPortableOperator(OfficeMathExpression expression) {
        string serialized = OfficeMathMarkup.ToMathMl(expression);
        XElement op = Assert.Single(XElement.Parse(serialized).Descendants(), element => element.Name.LocalName == "mo");
        Assert.Equal("postfix", (string?)op.Attribute("form"));
        Assert.Equal(".5em", (string?)op.Attribute("lspace"));
        Assert.Equal("25%", (string?)op.Attribute("rspace"));
        OfficeMathExpression reopened = OfficeMathMarkup.FromMathMl(serialized);
        Assert.Equal(expression, reopened);
        Assert.Equal(expression.GetHashCode(), reopened.GetHashCode());
    }

    private static string AccentSource(string kind, string wrapper, string symbol, string attributes) {
        string op = "<mo " + attributes + ">" + symbol + "</mo>";
        if (wrapper.Length != 0) op = "<" + wrapper + ">" + op + "</" + wrapper + ">";
        string element = kind == "underbar" ? "munder" : "mover";
        string accent = kind == "underbar" ? "accentunder" : "accent";
        return "<math><" + element + " " + accent + "='true'><mi>x</mi>" + op + "</" + element + "></math>";
    }
}
