using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public class DrawingMathTests {
    [Fact]
    public void MathFactoryBuildsReusableStructuredExpression() {
        OfficeMathExpression expression = OfficeMath.Row(
            OfficeMath.Fraction(
                OfficeMath.Row(OfficeMath.Superscript(OfficeMath.Identifier("x"), OfficeMath.Number("2")), OfficeMath.Operator("+"), OfficeMath.Number("1")),
                OfficeMath.Identifier("y")),
            OfficeMath.Operator("="),
            OfficeMath.Radical(OfficeMath.Identifier("z")));

        Assert.Equal(OfficeMathKind.Row, expression.Kind);
        Assert.Contains("(x^(2)+1)/(y)", expression.ToPlainText());
        Assert.Contains("sqrt(z)", expression.ToPlainText());
    }

    [Fact]
    public void MathMlRoundTripsCoreStructures() {
        OfficeMathExpression source = OfficeMath.Matrix(2, 2,
            OfficeMath.Identifier("a"), OfficeMath.Identifier("b"),
            OfficeMath.Identifier("c"), OfficeMath.Fraction(OfficeMath.Number("1"), OfficeMath.Number("2")));

        string mathMl = OfficeMathMarkup.ToMathMl(source, display: true);
        OfficeMathExpression parsed = OfficeMathMarkup.FromMathMl(mathMl);

        Assert.Contains("<math", mathMl);
        Assert.Contains("display=\"block\"", mathMl);
        Assert.Contains("<mfrac>", mathMl);
        Assert.Equal("[a, b; c, (1)/(2)]", parsed.ToPlainText());
    }

    [Fact]
    public void MathMlSerializerOutputRoundTripsFunctionsAndNaryOperators() {
        OfficeMathExpression[] expressions = {
            OfficeMath.Function("sin", OfficeMath.Superscript(OfficeMath.Identifier("x"), OfficeMath.Number("2"))),
            OfficeMath.Nary("∑", OfficeMath.Identifier("x"),
                OfficeMath.Row(OfficeMath.Identifier("i"), OfficeMath.Operator("="), OfficeMath.Number("0")),
                OfficeMath.Identifier("n")),
            OfficeMath.Nary("∫", OfficeMath.Identifier("f"))
        };

        Assert.All(expressions, expression =>
            Assert.Equal(expression, OfficeMathMarkup.FromMathMl(OfficeMathMarkup.ToMathMl(expression))));
    }

    [Fact]
    public void MathMlSerializerOutputPreservesUnderbars() {
        OfficeMathExpression expression = OfficeMath.Underbar(OfficeMath.Identifier("x"));

        Assert.Equal(expression, OfficeMathMarkup.FromMathMl(OfficeMathMarkup.ToMathMl(expression)));
    }

    [Theory]
    [InlineData(@"\frac{x^2+1}{y}", "(x^(2)+1)/(y)")]
    [InlineData(@"\sqrt[3]{x}", "root[3](x)")]
    [InlineData(@"\left(x+1\right)", "(x+1)")]
    public void LatexParserProducesStructuredExpressions(string latex, string plainText) {
        OfficeMathExpression expression = OfficeMathMarkup.FromLatex(latex);

        Assert.Equal(plainText, expression.ToPlainText());
        Assert.False(string.IsNullOrWhiteSpace(OfficeMathMarkup.ToMathMl(expression)));
    }

    [Fact]
    public void LatexSerializerSupportsMatricesLimitsAndDecorations() {
        OfficeMathExpression expression = OfficeMath.Row(
            OfficeMath.Nary("∑", OfficeMath.Identifier("x"), OfficeMath.Row(OfficeMath.Identifier("i"), OfficeMath.Operator("="), OfficeMath.Number("0")), OfficeMath.Identifier("n")),
            OfficeMath.Box(OfficeMath.Overbar(OfficeMath.Identifier("v"))),
            OfficeMath.Matrix(1, 2, OfficeMath.Number("1"), OfficeMath.Number("2")));

        string latex = OfficeMathMarkup.ToLatex(expression);

        Assert.Contains(@"\sum_{i=0}^{n} {x}", latex);
        Assert.Contains(@"\boxed{\overline{v}}", latex);
        Assert.Contains(@"\begin{bmatrix}1&2\end{bmatrix}", latex);
    }

    [Fact]
    public void LatexSerializerOutputRoundTripsFunctionsMatricesAndEquationArrays() {
        OfficeMathExpression[] expressions = {
            OfficeMath.Function("sin", OfficeMath.Superscript(OfficeMath.Identifier("x"), OfficeMath.Number("2"))),
            OfficeMath.Function("custom", OfficeMath.Identifier("z")),
            OfficeMath.Matrix(2, 2, OfficeMath.Number("1"), OfficeMath.Number("2"), OfficeMath.Number("3"), OfficeMath.Number("4")),
            OfficeMath.EquationArray(2, 2, OfficeMath.Identifier("a"), OfficeMath.Identifier("b"), OfficeMath.Identifier("c"), OfficeMath.Identifier("d"))
        };

        Assert.All(expressions, expression =>
            Assert.Equal(expression, OfficeMathMarkup.FromLatex(OfficeMathMarkup.ToLatex(expression))));
    }

    [Fact]
    public void LatexParserAllowsWhitespaceBeforeScripts() {
        OfficeMathExpression expression = OfficeMathMarkup.FromLatex("x ^{2} + y _ i");

        Assert.Equal("x^(2)+y_(i)", expression.ToPlainText());
        Assert.Equal(OfficeMathKind.Superscript, expression.Children[0].Kind);
        Assert.Equal(OfficeMathKind.Subscript, expression.Children[2].Kind);
    }

    [Fact]
    public void LatexSerializerEscapesTextWithoutTurningItIntoScripts() {
        OfficeMathExpression expression = OfficeMath.Text(@"literal x_y^z\{#%&}");
        string latex = OfficeMathMarkup.ToLatex(expression);

        Assert.Contains(@"\text{literal x\_y\^z\backslash \{\#\%\&\}}", latex, StringComparison.Ordinal);
        Assert.Equal(expression, OfficeMathMarkup.FromLatex(latex));
    }

    [Fact]
    public void PortableMarkupRoundTripsAdvancedSharedMathStructures() {
        OfficeMathExpression[] expressions = {
            OfficeMath.LeftSubSuperscript(OfficeMath.Identifier("T"), OfficeMath.Identifier("i"), OfficeMath.Identifier("j")),
            OfficeMath.LowerLimit(OfficeMath.Identifier("lim"), OfficeMath.Identifier("x")),
            OfficeMath.UpperLimit(OfficeMath.Identifier("max"), OfficeMath.Identifier("n")),
            OfficeMath.SlashedFraction(OfficeMath.Identifier("a"), OfficeMath.Identifier("b")),
            OfficeMath.Stack(OfficeMath.Identifier("a"), OfficeMath.Identifier("b")),
            OfficeMath.StretchStack(OfficeMath.Identifier("x"), OfficeMath.Identifier("y")),
            OfficeMath.DelimiterList("[", "]", ";", OfficeMath.Identifier("a"), OfficeMath.Identifier("b"))
        };

        Assert.All(expressions, expression => {
            Assert.Equal(expression, OfficeMathMarkup.FromMathMl(OfficeMathMarkup.ToMathMl(expression)));
            Assert.Equal(expression, OfficeMathMarkup.FromLatex(OfficeMathMarkup.ToLatex(expression)));
            OfficeMathLayoutMetrics metrics = OfficeMathRenderer.Measure(expression);
            Assert.True(metrics.Width > 0D);
            Assert.True(metrics.Height > 0D);
        });
    }

    [Fact]
    public void PortableMathParsersRejectExcessiveNestingWithStableCode() {
        string latex = new string('{', 20) + "x" + new string('}', 20);
        OfficeMathParseException latexError = Assert.Throws<OfficeMathParseException>(() => OfficeMathMarkup.FromLatex(latex, 8));
        string mathMl = "<math xmlns=\"http://www.w3.org/1998/Math/MathML\">" +
            string.Concat(Enumerable.Repeat("<mrow>", 20)) + "<mi>x</mi>" +
            string.Concat(Enumerable.Repeat("</mrow>", 20)) + "</math>";
        OfficeMathParseException mathMlError = Assert.Throws<OfficeMathParseException>(() => OfficeMathMarkup.FromMathMl(mathMl, 8));

        Assert.Equal("DRAWING_MATH_DEPTH", latexError.Code);
        Assert.Equal("DRAWING_MATH_DEPTH", mathMlError.Code);
    }

    [Fact]
    public void LatexRoundTripsNestedDelimiterPairs() {
        OfficeMathExpression expression = OfficeMath.Delimited(
            OfficeMath.Row(
                OfficeMath.Identifier("x"),
                OfficeMath.Operator("+"),
                OfficeMath.Delimited(OfficeMath.Identifier("y"), "[", "]")),
            "(", ")");

        Assert.Equal(expression, OfficeMathMarkup.FromLatex(OfficeMathMarkup.ToLatex(expression)));
    }

    [Fact]
    public void MathRendererBuildsTightlySizedVectorScene() {
        OfficeMathExpression expression = OfficeMath.Row(
            OfficeMath.Fraction(OfficeMath.Number("1"), OfficeMath.Number("2")),
            OfficeMath.Operator("+"),
            OfficeMath.Radical(OfficeMath.Superscript(OfficeMath.Identifier("x"), OfficeMath.Number("2"))));

        OfficeMathLayoutMetrics measured = OfficeMathRenderer.Measure(expression);
        OfficeDrawing drawing = OfficeMathRenderer.Render(expression, new OfficeMathRenderOptions { Padding = 4, BackgroundColor = OfficeColor.White });

        Assert.True(measured.Width > 20);
        Assert.True(measured.Height > 20);
        Assert.Equal(measured.Width + 8, drawing.Width, 6);
        Assert.Equal(measured.Height + 8, drawing.Height, 6);
        Assert.Contains(drawing.Elements, element => element is OfficeDrawingText);
        Assert.Contains(drawing.Elements, element => element is OfficeDrawingShape);
    }

    [Fact]
    public void MathRendererRetainsPhantomSpaceWithoutPaintingItsContent() {
        OfficeMathExpression visible = OfficeMath.Identifier("wide");
        OfficeMathExpression phantom = OfficeMath.Phantom(visible);

        OfficeMathLayoutMetrics visibleMetrics = OfficeMathRenderer.Measure(visible);
        OfficeMathLayoutMetrics phantomMetrics = OfficeMathRenderer.Measure(phantom);
        OfficeDrawing drawing = OfficeMathRenderer.Render(phantom);

        Assert.Equal(visibleMetrics.Width, phantomMetrics.Width, 6);
        Assert.Equal(visibleMetrics.Height, phantomMetrics.Height, 6);
        Assert.Empty(drawing.Elements);
    }

    [Fact]
    public void MatrixRowsReserveAlignedAscentAndDescent() {
        OfficeMathExpression expression = OfficeMath.EquationArray(
            1,
            2,
            OfficeMath.UpperLimit(OfficeMath.Identifier("max"), OfficeMath.Identifier("n")),
            OfficeMath.LowerLimit(OfficeMath.Identifier("lim"), OfficeMath.Identifier("x")));

        OfficeDrawing drawing = OfficeMathRenderer.Render(expression, new OfficeMathRenderOptions { Padding = 0 });

        Assert.All(drawing.Elements.OfType<OfficeDrawingText>(), text => {
            Assert.True(text.Y >= -0.000001D);
            Assert.True(text.Y + text.Height <= drawing.Height + 0.000001D);
        });
    }
}
