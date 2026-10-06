using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingMathGlyphPlacementTests {
    [Fact]
    public void ScriptOriginsUseGlyphAdvancesBeforeReservingOverhangRoom() {
        var options = DrawingMathGlyphConstructionTests.Options();
        options.Fonts = new OfficeFontFaceCollection();
        options.Fonts.Add("Fixture Math", OfficeIMO.TestAssets.ManagedTextShapingTestAssets.CreateMathConstructionFont(overhangs: true));
        var drawing = OfficeMathRenderer.Render(OfficeMath.SubSuperscript(OfficeMath.Identifier("x"),
            OfficeMath.Identifier("i"), OfficeMath.Number("2")), options);
        var texts = drawing.Elements.OfType<OfficeDrawingText>().ToDictionary(t => t.Text);
        // x origin 2 + advance 10 + italic 3 - kern 1; sub origin has 4.2 px of bearing room.
        Assert.Equal(texts["x"].X + 14D, texts["2"].X, 6);
        Assert.Equal(texts["x"].X + 5.8D, texts["i"].X, 6);
        Assert.All(texts.Values, t => Assert.True(t.X >= 0D && t.X + t.Width <= drawing.Width));
    }

    [Fact]
    public void AccentStretchOptOutSurvivesRoundTripAndKeepsNaturalGlyph() {
        const string markup = "<math><mover accent='true'><mrow><mtext>x</mtext><mtext>x</mtext><mtext>x</mtext></mrow><mo stretchy='false'>^</mo></mover></math>";
        var expression = OfficeMathMarkup.FromMathMl(markup);
        var transformed = expression.TransformTextCase(OfficeTextCase.None);
        var serialized = OfficeMathMarkup.ToMathMl(transformed);
        Assert.Contains("stretchy=\"false\"", serialized);
        Assert.Equal(expression, OfficeMathMarkup.FromMathMl(serialized));
        var drawing = OfficeMathRenderer.Render(transformed, DrawingMathGlyphConstructionTests.Options());
        Assert.Empty(drawing.Elements.OfType<OfficeDrawingShape>());
        Assert.Equal(10D, Assert.Single(drawing.Elements.OfType<OfficeDrawingText>(), t => t.Text == "^").Width, 6);
    }

    [Fact]
    public void LargeOperatorLimitsKeepTheSymbolCenterWhenTheUpperLimitIsWider() {
        var options = DrawingMathGlyphConstructionTests.Options(); options.DisplayStyle = false;
        var upper = OfficeMath.Row(OfficeMath.Identifier("x"), OfficeMath.Identifier("x"), OfficeMath.Identifier("x"));
        var drawing = OfficeMathRenderer.Render(OfficeMath.Nary("∑", OfficeMath.Identifier("y"),
            OfficeMath.Identifier("i"), upper), options);
        var symbol = Assert.Single(drawing.Elements.OfType<OfficeDrawingText>(), t => t.Text == "∑");
        var lower = Assert.Single(drawing.Elements.OfType<OfficeDrawingText>(), t => t.Text == "i");
        var upperTexts = drawing.Elements.OfType<OfficeDrawingText>().Where(t => t.Text == "x").ToArray();
        double center = symbol.X + symbol.Width / 2D;
        Assert.Equal(center + 2D, (upperTexts[0].X + upperTexts[2].X + upperTexts[2].Width) / 2D, 6);
        Assert.Equal(center - 2D, lower.X + lower.Width / 2D, 6);
    }

    [Fact]
    public void WideAccentUsesHorizontalConnectorsWithOneLogicalTokenAndTightInk() {
        var content = OfficeMath.Row(Enumerable.Range(0, 6).Select(_ => OfficeMath.Identifier("x")).ToArray());
        var drawing = OfficeMathRenderer.Render(OfficeMath.Accent(content, "^"), DrawingMathGlyphConstructionTests.Options());
        var shape = Assert.Single(drawing.Elements.OfType<OfficeDrawingShape>());
        Assert.Equal(60D, shape.Shape.Width, 6);
        Assert.Equal(4D, shape.Shape.Height, 6);
        Assert.Equal(6, drawing.Elements.OfType<OfficeDrawingText>().Count(t => t.Text == "x"));
        Assert.Single(drawing.Elements.OfType<OfficeDrawingText>(), t => t.Text == "^");
        Assert.All(drawing.Elements.OfType<OfficeDrawingText>(), t => Assert.True(t.Y + t.Height <= drawing.Height));
    }

    [Fact]
    public void RightScriptsUseItalicCorrectionAndTheMinimumOfTwoCornerHeights() {
        var options = DrawingMathGlyphConstructionTests.Options();
        var drawing = OfficeMathRenderer.Render(OfficeMath.SubSuperscript(
            OfficeMath.Identifier("x"), OfficeMath.Identifier("i"), OfficeMath.Number("2")), options);
        var texts = drawing.Elements.OfType<OfficeDrawingText>().ToDictionary(t => t.Text);
        // 150-unit italic correction adds 3 px; minimum top-right kern is -50 units (-1 px).
        Assert.Equal(texts["x"].X + texts["x"].Width + 2D, texts["2"].X, 6);
        Assert.Equal(texts["x"].X + texts["x"].Width - 2D, texts["i"].X, 6); // bottom-right -100.
        Assert.True(drawing.Width >= texts["2"].X + texts["2"].Width);
    }

    [Fact]
    public void AccentAttachmentUsesBothGlyphAnchorsWithoutChangingTheBaseSize() {
        var drawing = OfficeMathRenderer.Render(OfficeMath.Accent(OfficeMath.Identifier("x"), "^"),
            DrawingMathGlyphConstructionTests.Options());
        var basis = Assert.Single(drawing.Elements.OfType<OfficeDrawingText>(), t => t.Text == "x");
        var accent = Assert.Single(drawing.Elements.OfType<OfficeDrawingText>(), t => t.Text == "^");
        Assert.Equal(20D, accent.Font.Size, 6);
        Assert.Equal(basis.X + 7D, accent.X + 2D, 6); // 350 and 100 design units.
        Assert.Equal(4D, accent.Height, 6); // Ink at 500..700; the baseline gap is not part of the accent.
        Assert.Equal(accent.Y + accent.Height, basis.Y, 6);
        Assert.InRange(accent.X, 0D, drawing.Width - accent.Width + .000001D);
    }

    [Fact]
    public void RootAndAccentPreserveDisplayStyleInsideTheirCrampedBase() {
        var options = DrawingMathGlyphConstructionTests.Options();
        var fraction = OfficeMath.Fraction(OfficeMath.Identifier("x"), OfficeMath.Identifier("y"));
        foreach (var expression in new[] { OfficeMath.Radical(fraction), OfficeMath.Accent(fraction, "^") }) {
            var drawing = OfficeMathRenderer.Render(expression, options);
            Assert.All(drawing.Elements.OfType<OfficeDrawingText>().Where(t => t.Text == "x" || t.Text == "y"),
                t => Assert.Equal(20D, t.Font.Size, 6));
        }
    }

    [Fact]
    public void DisabledMetricsKeepCallerSpacingAndAccentScale() {
        var options = DrawingMathGlyphConstructionTests.Options(); options.UseFontMathMetrics = false; options.ScriptScale = .5D;
        var scripts = OfficeMathRenderer.Render(OfficeMath.Superscript(OfficeMath.Identifier("x"), OfficeMath.Number("2")), options);
        var texts = scripts.Elements.OfType<OfficeDrawingText>().ToDictionary(t=>t.Text);
        Assert.Equal(texts["x"].X + texts["x"].Width, texts["2"].X, 6);
        var accent = OfficeMathRenderer.Render(OfficeMath.Accent(OfficeMath.Identifier("x"), "^"), options);
        Assert.Equal(10D, Assert.Single(accent.Elements.OfType<OfficeDrawingText>(), t=>t.Text == "^").Font.Size, 6);
    }
}
