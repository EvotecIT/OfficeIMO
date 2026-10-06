using OfficeIMO.Drawing;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingMathFontConstantsTests {
    [Fact]
    public void FontProgramExposesSignedDesignMetricsAndPercentages() {
        OfficeMathRenderOptions options = Options();
        var constants = Assert.IsAssignableFrom<IOfficeMathFontProgram>(Assert.Single(options.Fonts.Faces).Program).MathConstants;
        Assert.NotNull(constants);
        Assert.Equal(1000, constants!.UnitsPerEm);
        Assert.Equal(258, constants.GetValue(OfficeMathConstant.AxisHeight));
        Assert.Equal(68, constants.GetValue(OfficeMathConstant.FractionRuleThickness));
        Assert.Equal(55, constants.GetValue(OfficeMathConstant.ScriptScriptPercentScaleDown));
        Assert.Equal(-335, constants.GetValue(OfficeMathConstant.RadicalKernAfterDegree));
        Assert.Throws<ArgumentOutOfRangeException>(() => constants.GetValue((OfficeMathConstant)56));
    }

    [Fact]
    public void ProviderConstantsAreDetachedAndValidateThePublicContract() {
        var original = Assert.IsAssignableFrom<IOfficeMathFontProgram>(Assert.Single(Options().Fonts.Faces).Program).MathConstants!;
        var values = Enum.GetValues(typeof(OfficeMathConstant)).Cast<OfficeMathConstant>().ToDictionary(c => c, original.GetValue);
        var snapshot = new OfficeMathFontConstants(1000, values);
        values[OfficeMathConstant.AxisHeight] = 999;
        Assert.Equal(258, snapshot.GetValue(OfficeMathConstant.AxisHeight));
        values[OfficeMathConstant.ScriptPercentScaleDown] = 0;
        Assert.Throws<ArgumentException>(() => new OfficeMathFontConstants(1000, values));
        values.Remove(OfficeMathConstant.AxisHeight);
        Assert.Throws<ArgumentException>(() => new OfficeMathFontConstants(1000, values));
    }

    [Theory]
    [InlineData(0)] [InlineData(1)] [InlineData(2)] [InlineData(3)]
    public void InvalidMathDataKeepsTheUsableFontAndFallsBack(int damage) {
        var options = Options(table => {
            if (damage == 0) { table[4] = 255; table[5] = 255; } // Offset outside MATH, even though font has more bytes.
            if (damage == 1) { table[4] = 0; table[5] = 12; } // Constants run past their table end.
            if (damage == 2) table[1] = 2; // Unsupported table version.
            if (damage == 3) table[11] = 0; // Invalid script scale.
        });
        Assert.Null(Assert.IsAssignableFrom<IOfficeMathFontProgram>(Assert.Single(options.Fonts.Faces).Program).MathConstants);
        OfficeDrawing fallback = OfficeMathRenderer.Render(OfficeMath.Superscript(OfficeMath.Identifier("x"), OfficeMath.Number("2")), options);
        Assert.Equal(14.2D, Assert.Single(fallback.Elements.OfType<OfficeDrawingText>(), t => t.Text == "2").Font.Size, 6);
    }

    [Theory]
    [InlineData(true, 20D, 12.8D, 12.8D)]
    [InlineData(false, 14D, 11.7D, 11.7D)]
    public void FractionsUseFontAxisBaselineShiftsAndRuleThickness(bool display, double childSize, double rise, double drop) {
        var options = Options(); options.DisplayStyle = display;
        var expression = OfficeMath.Fraction(OfficeMath.Identifier("x"), OfficeMath.Identifier("y"));
        OfficeMathLayoutMetrics metrics = OfficeMathRenderer.Measure(expression, options);
        OfficeDrawing drawing = OfficeMathRenderer.Render(expression, options);
        var numerator = Assert.Single(drawing.Elements.OfType<OfficeDrawingText>(), t => t.Text == "x");
        var denominator = Assert.Single(drawing.Elements.OfType<OfficeDrawingText>(), t => t.Text == "y");
        var rule = Assert.Single(drawing.Elements.OfType<OfficeDrawingShape>());
        Assert.Equal(childSize, numerator.Font.Size, 6);
        Assert.Equal(rise, metrics.Baseline - Baseline(numerator), 6);
        Assert.Equal(drop, Baseline(denominator) - metrics.Baseline, 6);
        Assert.Equal(1.36D, rule.Shape.StrokeWidth, 6);
        Assert.Equal(5.16D, metrics.Baseline - rule.Y, 6);
        Assert.Equal(metrics.Height, drawing.Height, 6);
        Assert.InRange(denominator.Y - (rule.Y + .68D), display ? 3D : 1.36D, 20D);
    }

    [Fact]
    public void ScriptGeometryUsesFontShiftsAndContinuesScalingBeyondSecondDepth() {
        var expression = OfficeMath.SubSuperscript(OfficeMath.Identifier("x"), OfficeMath.Identifier("i"), OfficeMath.Number("2"));
        var drawing = OfficeMathRenderer.Render(expression, Options());
        var basis = Assert.Single(drawing.Elements.OfType<OfficeDrawingText>(), t => t.Text == "x");
        var sup = Assert.Single(drawing.Elements.OfType<OfficeDrawingText>(), t => t.Text == "2");
        var sub = Assert.Single(drawing.Elements.OfType<OfficeDrawingText>(), t => t.Text == "i");
        Assert.Equal(9.4D, Baseline(basis) - Baseline(sup), 6);
        Assert.Equal(4.2D, Baseline(sub) - Baseline(basis), 6);
        Assert.True(sub.Y - (sup.Y + sup.Height) >= 3D);
        var nested = OfficeMath.Superscript(OfficeMath.Identifier("x"), OfficeMath.Superscript(OfficeMath.Identifier("y"),
            OfficeMath.Superscript(OfficeMath.Number("2"), OfficeMath.Number("3"))));
        var sizes = OfficeMathRenderer.Render(nested, Options()).Elements.OfType<OfficeDrawingText>().ToDictionary(t => t.Text, t => t.Font.Size);
        Assert.Equal(14D, sizes["y"], 6);
        Assert.Equal(11D, sizes["2"], 6);
        Assert.Equal(7.81D, sizes["3"], 6);
    }

    [Fact]
    public void PrescriptSpacePrecedesScriptsWithoutSeparatingThemFromTheBase() {
        var drawing = OfficeMathRenderer.Render(OfficeMath.LeftSubSuperscript(
            OfficeMath.Identifier("x"), OfficeMath.Identifier("i"), OfficeMath.Number("2")), Options());
        var basis = Assert.Single(drawing.Elements.OfType<OfficeDrawingText>(), t => t.Text == "x");
        var scripts = drawing.Elements.OfType<OfficeDrawingText>().Where(t => t.Text != "x").ToArray();
        Assert.Equal(.8D, scripts.Min(t => t.X), 6);
        Assert.All(scripts, t => Assert.Equal(basis.X, t.X + t.Width, 6));
        Assert.Equal(basis.X + basis.Width, drawing.Width, 6);
    }

    [Theory]
    [InlineData(false)] [InlineData(true)]
    public void CrowdedRightAndLeftScriptsPreserveFontMinimumGapAndAllInk(bool left) {
        var fraction = OfficeMath.Fraction(OfficeMath.Identifier("y"), OfficeMath.Number("2"));
        var expression = left ? OfficeMath.LeftSubSuperscript(OfficeMath.Identifier("x"), fraction, fraction)
            : OfficeMath.SubSuperscript(OfficeMath.Identifier("x"), fraction, fraction);
        OfficeDrawing drawing = OfficeMathRenderer.Render(expression, Options());
        var ys = drawing.Elements.OfType<OfficeDrawingText>().Where(t => t.Text == "y").OrderBy(t => t.Y).ToArray();
        var twos = drawing.Elements.OfType<OfficeDrawingText>().Where(t => t.Text == "2").OrderBy(t => t.Y).ToArray();
        Assert.Equal(2, ys.Length); Assert.Equal(2, twos.Length);
        Assert.True(ys[1].Y - (twos[0].Y + twos[0].Height) >= 3D - 0.000001D);
        Assert.All(drawing.Elements.OfType<OfficeDrawingText>(), t => {
            Assert.InRange(t.Y, 0D, drawing.Height - t.Height + 0.000001D);
            Assert.InRange(t.X, 0D, drawing.Width - t.Width + 0.000001D);
        });
    }

    [Theory]
    [InlineData(true, 2.04D)] [InlineData(false, 18.18D)]
    public void BarsReserveFontRuleGapAndOutsideInkSpace(bool over, double ruleY) {
        OfficeDrawing drawing = OfficeMathRenderer.Render(over ? OfficeMath.Overbar(OfficeMath.Identifier("x"))
            : OfficeMath.Underbar(OfficeMath.Identifier("x")), Options());
        Assert.Equal(20.22D, drawing.Height, 6);
        var rule = Assert.Single(drawing.Elements.OfType<OfficeDrawingShape>());
        Assert.Equal(1.36D, rule.Shape.StrokeWidth, 6);
        Assert.Equal(ruleY, rule.Y, 6);
    }

    [Fact]
    public void DensityScalesFontGeometryAndDisabledMetricsHonorCallerSettings() {
        var expression = OfficeMath.Fraction(OfficeMath.Identifier("x"), OfficeMath.Identifier("y"));
        var options = Options(); var first = OfficeMathRenderer.Render(expression, options);
        options.Dpi = 144; var doubled = OfficeMathRenderer.Render(expression, options);
        Assert.Equal(first.Width * 2, doubled.Width, 6); Assert.Equal(first.Height * 2, doubled.Height, 6);
        Assert.Equal(2.72D, Assert.Single(doubled.Elements.OfType<OfficeDrawingShape>()).Shape.StrokeWidth, 6);
        options.Dpi = 72; options.UseFontMathMetrics = false; options.ScriptScale = .5D; options.RuleThickness = 3D;
        options.DisplayStyle = false;
        var custom = OfficeMathRenderer.Render(expression, options);
        Assert.All(custom.Elements.OfType<OfficeDrawingText>(), t => Assert.Equal(10D, t.Font.Size, 6));
        Assert.Equal(3D, Assert.Single(custom.Elements.OfType<OfficeDrawingShape>()).Shape.StrokeWidth, 6);
    }

    [Theory]
    [InlineData(true, 6D)] [InlineData(false, 3.6D)]
    public void LimitsRespectFontMinimumGapAndBaselineDistance(bool over, double gap) {
        var expression = over ? OfficeMath.UpperLimit(OfficeMath.Identifier("x"), OfficeMath.Identifier("y"))
            : OfficeMath.LowerLimit(OfficeMath.Identifier("x"), OfficeMath.Identifier("y"));
        var drawing = OfficeMathRenderer.Render(expression, Options());
        var texts = drawing.Elements.OfType<OfficeDrawingText>().OrderBy(t => t.Y).ToArray();
        Assert.Equal(gap, texts[1].Y - texts[0].Y - texts[0].Height, 6);
    }

    private static double Baseline(OfficeDrawingText text) => text.Y + text.Font.Size + text.BaselineOffset;
    private static OfficeMathRenderOptions Options(Action<byte[]>? edit = null) {
        var options = new OfficeMathRenderOptions { Font = new OfficeFontInfo("Fixture Math", 20D), Padding = 0D };
        options.Fonts.Add("Fixture Math", ManagedTextShapingTestAssets.CreateMathFont(edit));
        return options;
    }
}
