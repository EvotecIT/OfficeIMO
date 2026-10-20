using OfficeIMO.Drawing;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingMathOperatorSpacingTests {
    [Fact]
    public void AuthoredOperatorSidesAndFormSurvivePortableMathAndTextTransformation() {
        var expression = Parse("<mi>x</mi><mo form='PREFIX' lspace='.5em' rspace='25%'>+</mo><mi>y</mi>");
        var transformed = expression.TransformTextCase(OfficeTextCase.Uppercase);
        string markup = OfficeMathMarkup.ToMathMl(transformed);
        var reopened = OfficeMathMarkup.FromMathMl(markup);
        Assert.Equal(transformed, reopened);
        Assert.Equal(transformed.GetHashCode(), reopened.GetHashCode());
        Assert.Equal("X+Y", reopened.ToPlainText());
        Assert.Contains("form=\"prefix\"", markup, StringComparison.Ordinal);
        Assert.Contains("lspace=\".5em\"", markup, StringComparison.Ordinal);
        Assert.Contains("rspace=\"25%\"", markup, StringComparison.Ordinal);
        Assert.NotEqual(expression, Parse("<mi>x</mi><mo form='prefix' lspace='0' rspace='25%'>+</mo><mi>y</mi>"));
    }

    [Theory]
    [InlineData("<mo>+</mo><mi>x</mi>", 20D)]
    [InlineData("<mi>x</mi><mo>+</mo><mi>y</mi>", 30D + 80D / 9D)]
    [InlineData("<mi>x</mi><mo form='prefix'>+</mo><mi>y</mi>", 30D)]
    [InlineData("<mi>x</mi><mo form='invalid'>+</mo><mi>y</mi>", 30D + 80D / 9D)]
    [InlineData("<mo form='infix'>+</mo><mi>x</mi>", 20D + 80D / 9D)]
    [InlineData("<mo>+</mo>", 10D + 80D / 9D)]
    [InlineData("<mtext>y</mtext><mo>+</mo><mi>x</mi>", 30D)]
    [InlineData("<mphantom><mi>y</mi></mphantom><mo>+</mo><mi>x</mi>", 30D + 80D / 9D)]
    [InlineData("<mphantom><mtext>y</mtext></mphantom><mo>+</mo><mi>x</mi>", 30D)]
    [InlineData("<mo>=</mo><mi>x</mi>", 20D + 100D / 9D)]
    [InlineData("<mi>x</mi><mo>,</mo><mi>y</mi>", 30D + 10D / 3D)]
    [InlineData("<mi>x</mi><mo>×</mo><mi>y</mi>", 30D + 20D / 3D)]
    public void OrdinaryRowsUseAuthoredFormOrNonSpaceLikePositionAndDictionaryFallback(string row, double width) {
        Assert.Equal(width, OfficeMathRenderer.Measure(Parse(row), Options()).Width, 6);
    }

    [Fact]
    public void AuthoredSpacingMovesOnlyItsOwnSideOfTheOperator() {
        OfficeDrawing drawing = OfficeMathRenderer.Render(Parse("<mi>x</mi><mo lspace='0' rspace='.75em'>=</mo><mi>y</mi>"), Options());
        Assert.Equal(45D, drawing.Width, 6);
        Assert.Equal(10D, Assert.Single(drawing.Elements.OfType<OfficeDrawingText>(), t => t.Text == "=").X, 6);
        Assert.Equal(35D, Assert.Single(drawing.Elements.OfType<OfficeDrawingText>(), t => t.Text == "y").X, 6);
        Assert.Equal("x=y", string.Concat(drawing.Elements.OfType<OfficeDrawingText>().Select(t => t.Text)));
    }

    [Fact]
    public void PercentageSpacingUsesEachDictionarySideInsteadOfTheEmOrContainer() {
        OfficeDrawing drawing = OfficeMathRenderer.Render(Parse("<mi>x</mi><mo lspace='50%' rspace='200%'>=</mo><mi>y</mi>"), Options());
        Assert.Equal(30D + 125D / 9D, drawing.Width, 6);
        Assert.Equal(10D + 25D / 9D, Assert.Single(drawing.Elements.OfType<OfficeDrawingText>(), t => t.Text == "=").X, 6);
        OfficeDrawing punctuation = OfficeMathRenderer.Render(Parse("<mi>x</mi><mo lspace='200%' rspace='200%'>,</mo><mi>y</mi>"), Options());
        Assert.Equal(10D, Assert.Single(punctuation.Elements.OfType<OfficeDrawingText>(), t => t.Text == ",").X, 6);
        Assert.Equal(30D + 20D / 3D, punctuation.Width, 6);
    }

    [Theory]
    [InlineData(72D, 33D)]
    [InlineData(96D, 44D)]
    public void StandaloneCssPixelsConvertToTheDeclaredDrawingDensity(double dpi, double width) {
        var options = Options(); options.Dpi = dpi;
        Assert.Equal(width, OfficeMathRenderer.Measure(Parse("<mi>x</mi><mo lspace='4px' rspace='0'>=</mo><mi>y</mi>"), options).Width, 6);
    }

    [Theory]
    [InlineData("-1em")]
    [InlineData("1 em")]
    [InlineData("1.em")]
    [InlineData("2ex")]
    [InlineData("calc(1em + 1px)")]
    public void UnsupportedLengthGrammarRetainsTheDictionaryFallback(string spacing) {
        var expression = Parse("<mi>x</mi><mo lspace='" + spacing + "'>=</mo><mi>y</mi>");
        Assert.Equal(30D + 100D / 9D, OfficeMathRenderer.Measure(expression, Options()).Width, 6);
    }

    [Theory]
    [InlineData("<mo lspace='1e308px'>+</mo><mo lspace='1e308px'>+</mo>")]
    [InlineData("<mtable><mtr><mtd><mrow><mi>x</mi><mo lspace='1e308px'>+</mo><mi>y</mi></mrow></mtd><mtd><mrow><mi>x</mi><mo lspace='1e308px'>+</mo><mi>y</mi></mrow></mtd></mtr></mtable>")]
    public void FiniteAuthoredSidesCannotOverflowTheMeasuredContainer(string markup) {
        var expression = Parse(markup);
        var options = Options(); options.Dpi = 96D;
        double measuredWidth = double.NaN;
        Exception? error = Record.Exception(() => { measuredWidth = OfficeMathRenderer.Measure(expression, options).Width; });
        Assert.True(error is ArgumentOutOfRangeException,
            $"Expected finite-bounds rejection; observed width {measuredWidth}, exception {error?.GetType().Name ?? "none"}.");
    }

    [Fact]
    public void NaryImportDoesNotEraseAuthoredOperatorMetadata() {
        var expression = Parse("<mo lspace='.5em' rspace='0'>∑</mo><mi>x</mi>");
        Assert.Equal(OfficeMathKind.Operator, expression.Children[0].Kind);
        Assert.Equal("∑x", expression.ToPlainText());
        Assert.Equal(expression, OfficeMathMarkup.FromMathMl(OfficeMathMarkup.ToMathMl(expression)));
        var limited = Parse("<mover><mo lspace='.5em'>∑</mo><mi>y</mi></mover><mi>x</mi>");
        Assert.Equal(limited, OfficeMathMarkup.FromMathMl(OfficeMathMarkup.ToMathMl(limited)));
    }

    [Fact]
    public void DesignedFenceStretchingRetainsOutsideSpacingAndOnceOnlyLogicalGlyphs() {
        const string body = "<mfrac><mi>x</mi><mi>y</mi></mfrac><mo>)</mo>";
        var options = DrawingMathGlyphConstructionTests.Options();
        OfficeDrawing control = OfficeMathRenderer.Render(Parse("<mo>(</mo>" + body), options);
        OfficeDrawing spaced = OfficeMathRenderer.Render(Parse("<mo lspace='.5em' rspace='.25em'>(</mo>" + body), options);
        var before = control.Elements.OfType<OfficeDrawingShape>().Where(s => s.Shape.Kind == OfficeShapeKind.Path).ToArray();
        var after = spaced.Elements.OfType<OfficeDrawingShape>().Where(s => s.Shape.Kind == OfficeShapeKind.Path).ToArray();
        Assert.Equal(2, before.Length); Assert.Equal(2, after.Length);
        Assert.Equal(control.Width + 15D, spaced.Width, 6);
        Assert.Equal(control.Height, spaced.Height, 6);
        Assert.Equal(before[0].X + 10D, after[0].X, 6);
        Assert.Equal(before[1].X + 15D, after[1].X, 6);
        Assert.Single(spaced.Elements.OfType<OfficeDrawingText>(), t => t.Text == "(");
        Assert.Single(spaced.Elements.OfType<OfficeDrawingText>(), t => t.Text == ")");
    }

    private static OfficeMathExpression Parse(string row) => OfficeMathMarkup.FromMathMl("<math><mrow>" + row + "</mrow></math>");

    private static OfficeMathRenderOptions Options() {
        var options = new OfficeMathRenderOptions { Font = new OfficeFontInfo("Spacing Math", 20D), Padding = 0D };
        options.Fonts.Add("Spacing Math", ManagedTextShapingTestAssets.CreateFont('x', 'y', '+', '-', '−', '=', '×', ',', '∑'));
        return options;
    }
}
