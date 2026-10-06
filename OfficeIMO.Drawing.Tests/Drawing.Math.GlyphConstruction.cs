using OfficeIMO.Drawing;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingMathGlyphConstructionTests {
    [Fact]
    public void DisplayOperatorSelectsDesignedVariantAtOriginalEmSize() {
        var options = Options();
        var drawing = OfficeMathRenderer.Render(OfficeMath.Operator("∑"), options);
        var outline = Assert.Single(drawing.Elements.OfType<OfficeDrawingShape>());
        Assert.Equal(10D, outline.Shape.Width, 6);
        Assert.Equal(44D, outline.Shape.Height, 6); // first height >= 1800 is 2200 units at 20 px.
        var logical = Assert.Single(drawing.Elements.OfType<OfficeDrawingText>());
        Assert.Equal("∑", logical.Text); Assert.Equal(20D, logical.Font.Size);
        options.DisplayStyle = false;
        Assert.Empty(OfficeMathRenderer.Render(OfficeMath.Operator("∑"), options).Elements.OfType<OfficeDrawingShape>());
    }

    [Fact]
    public void TallFencesAssembleWithBoundedConnectorsWithoutWideningTheStroke() {
        var options = Options();
        var content = OfficeMath.Fraction(OfficeMath.Fraction(OfficeMath.Identifier("x"), OfficeMath.Identifier("y")),
            OfficeMath.Fraction(OfficeMath.Identifier("y"), OfficeMath.Number("2")));
        var drawing = OfficeMathRenderer.Render(OfficeMath.Delimited(content), options);
        var fences = drawing.Elements.OfType<OfficeDrawingShape>().Where(s => s.Shape.Kind == OfficeShapeKind.Path).ToArray();
        Assert.Equal(2, fences.Length);
        Assert.All(fences, f => { Assert.Equal(10D, f.Shape.Width, 6); Assert.True(f.Shape.Height > 44D); });
        Assert.Equal(1, drawing.Elements.OfType<OfficeDrawingText>().Count(t => t.Text == "("));
        Assert.Equal(1, drawing.Elements.OfType<OfficeDrawingText>().Count(t => t.Text == ")"));
        Assert.All(drawing.Elements.OfType<OfficeDrawingShape>(), s => {
            Assert.InRange(s.X, 0D, drawing.Width); Assert.InRange(s.Y, 0D, drawing.Height);
            if (s.Shape.Kind == OfficeShapeKind.Path) Assert.True(s.Y + s.Shape.Height <= drawing.Height + .000001D);
        });
    }

    [Fact]
    public void RootDegreeUsesTwoDepthsAndFontRadicalRule() {
        var drawing = OfficeMathRenderer.Render(OfficeMath.Radical(OfficeMath.Identifier("x"), OfficeMath.Number("2")), Options());
        var degree = Assert.Single(drawing.Elements.OfType<OfficeDrawingText>(), t => t.Text == "2");
        Assert.Equal(11D, degree.Font.Size, 6);
        var rule = Assert.Single(drawing.Elements.OfType<OfficeDrawingShape>(), s => s.Shape.Kind == OfficeShapeKind.Line);
        Assert.Equal(1.36D, rule.Shape.StrokeWidth, 6);
        var radical = Assert.Single(drawing.Elements.OfType<OfficeDrawingShape>(), s => s.Shape.Kind == OfficeShapeKind.Path);
        Assert.Equal(radical.Y + rule.Shape.StrokeWidth / 2D, rule.Y, 6);
        Assert.All(drawing.Elements.OfType<OfficeDrawingText>(), t => Assert.InRange(t.Y, 0D, drawing.Height - t.Height + .000001D));
    }

    [Fact]
    public void InvalidConstructionOffsetRetainsConstantsAndFallbackText() {
        var options = Options(math => { math[8] = 255; math[9] = 255; });
        var font = Assert.Single(options.Fonts.Faces).Program;
        Assert.NotNull(((IOfficeMathFontProgram)font).MathConstants);
        Assert.Null(((IOfficeMathGlyphProgram)font).MathGlyphData);
        var drawing = OfficeMathRenderer.Render(OfficeMath.Operator("∑"), options);
        Assert.Empty(drawing.Elements.OfType<OfficeDrawingShape>());
        Assert.Equal("∑", Assert.Single(drawing.Elements.OfType<OfficeDrawingText>()).Text);
    }

    [Fact]
    public void NativeMathMlFencesStretchAndAuthoredOverridesSurviveSerializationAndTransformation() {
        const string source = "<math><mrow><mo>(</mo><mfrac><mtext>x</mtext><mtext>y</mtext></mfrac><mo stretchy='false'>)</mo></mrow></math>";
        var expression = OfficeMathMarkup.FromMathMl(source);
        var roundTrip = OfficeMathMarkup.FromMathMl(OfficeMathMarkup.ToMathMl(expression));
        Assert.Equal(expression, roundTrip);
        var drawing = OfficeMathRenderer.Render(expression, Options());
        Assert.Single(drawing.Elements.OfType<OfficeDrawingShape>(), s => s.Shape.Kind == OfficeShapeKind.Path);
        Assert.Equal(20D, Assert.Single(drawing.Elements.OfType<OfficeDrawingText>(), t=>t.Text == ")").Font.Size);
        var disabled = OfficeMathMarkup.FromMathMl("<math><mrow><mo largeop='false'>∑</mo><mtext>x</mtext></mrow></math>");
        var transformed = disabled.TransformTextCase(OfficeTextCase.Uppercase);
        Assert.Contains("largeop=\"false\"", OfficeMathMarkup.ToMathMl(transformed));
        Assert.Empty(OfficeMathRenderer.Render(transformed, Options()).Elements.OfType<OfficeDrawingShape>());
    }

    [Theory]
    [InlineData(20D)] [InlineData(40D)] [InlineData(53.33333333333333D)]
    public void TightNestedRootContainsEveryPaintFrameAtCssAndPointSizes(double size) {
        var options = Options(); options.Font = options.Font.WithSize(size);
        var fraction = OfficeMath.Fraction(OfficeMath.Identifier("x"), OfficeMath.Identifier("y"));
        var expression = OfficeMath.Radical(OfficeMath.Fraction(fraction, OfficeMath.Fraction(fraction, fraction)));
        var metrics = OfficeMathRenderer.Measure(expression, options);
        var drawing = OfficeMathRenderer.Render(expression, options);
        Assert.Equal(metrics.Width, drawing.Width, 6); Assert.Equal(metrics.Height, drawing.Height, 6);
        Assert.All(drawing.Elements.OfType<OfficeDrawingText>(), t=> {
            Assert.True(t.X >= 0D && t.X + t.Width <= drawing.Width);
            Assert.True(t.Y >= 0D && t.Y + t.Height <= drawing.Height);
        });
    }

    [Fact]
    public void DegreeStartsAtDepthTwoBeforeNestedSuperscriptScaling() {
        var degree = OfficeMath.Superscript(OfficeMath.Number("2"), OfficeMath.Identifier("i"));
        var drawing = OfficeMathRenderer.Render(OfficeMath.Radical(OfficeMath.Identifier("x"), degree), Options());
        Assert.Equal(7.81D, Assert.Single(drawing.Elements.OfType<OfficeDrawingText>(), t=>t.Text == "i").Font.Size, 6);
    }

    [Theory]
    [InlineData(0)] [InlineData(1)] [InlineData(2)]
    public void MalformedGlyphListsAndKernOffsetsDoNotEscapeTheMathTable(int damage) {
        var options = Options(math=> {
            int variants = (math[8] << 8) | math[9];
            int construction = variants + ((math[variants + 10] << 8) | math[variants + 11]);
            if (damage == 0) { math[construction + 4] = 255; math[construction + 5] = 255; } // glyph outside maxp.
            if (damage == 1) { math[construction + 2] = 255; math[construction + 3] = 255; } // bounded records.
            if (damage == 2) {
                int info = (math[6] << 8) | math[7];
                math[info + 6] = 255; math[info + 7] = 255; // kern offset outside owning table.
            }
        });
        var font = Assert.Single(options.Fonts.Faces).Program;
        Assert.Null(((IOfficeMathGlyphProgram)font).MathGlyphData);
        Assert.NotNull(((IOfficeMathFontProgram)font).MathConstants);
        Assert.NotEmpty(OfficeMathRenderer.Render(OfficeMath.Identifier("x"), options).Elements);
    }

    internal static OfficeMathRenderOptions Options(Action<byte[]>? edit = null) {
        var options = new OfficeMathRenderOptions { Font = new OfficeFontInfo("Fixture Math", 20D), Padding = 0D };
        options.Fonts.Add("Fixture Math", ManagedTextShapingTestAssets.CreateMathConstructionFont(edit)); return options;
    }
}
