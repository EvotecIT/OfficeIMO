using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlMathOperatorDiagnosticRowTests {
    [Theory]
    [InlineData("<semantics><mn>1</mn><annotation-xml encoding='application/mathml+xml'><mrow><mo lspace='.5em'>+</mo></mrow></annotation-xml></semantics>")]
    [InlineData("<semantics><mn>1</mn><annotation-xml encoding='application/mathml+xml'><mrow><mn>1</mn><mo lspace='2ex'>+</mo><mn>2</mn></mrow></annotation-xml></semantics>")]
    [InlineData("<semantics><annotation encoding='text/plain'>ignored</annotation><mn>1</mn><mrow><mn>1</mn><mo lspace='2ex'>+</mo><mn>2</mn></mrow></semantics>")]
    public void IgnoredAnnotationAndUnselectedSemanticsContentDoesNotCreateMathLoss(string content) {
        HtmlRenderDocument rendered = Render(content);
        HtmlRenderDocument control = Render("<mn>1</mn>");

        Assert.Equal(control.Text, rendered.Text);
        Assert.Equal(Formula(control).Drawing.Width, Formula(rendered).Drawing.Width, 6);
        AssertNoMathLoss(control);
        AssertNoMathLoss(rendered);
    }

    [Theory]
    [InlineData("mtd")]
    [InlineData("mphantom")]
    [InlineData("menclose")]
    [InlineData("msqrt")]
    public void ImplicitOrdinaryRowsApplyAuthoredSpacingWithoutFalseLoss(string owner) {
        const string spaced = "<mn>1</mn><mo lspace='.5em' rspace='0'>+</mo><mn>2</mn>";
        const string zero = "<mn>1</mn><mo lspace='0' rspace='0'>+</mo><mn>2</mn>";
        HtmlRenderDocument implicitSpaced = Render(Wrap(owner, spaced));
        HtmlRenderDocument implicitZero = Render(Wrap(owner, zero));
        HtmlRenderDocument ordinarySpaced = Render(Wrap(owner, "<mrow>" + spaced + "</mrow>"));
        HtmlRenderDocument ordinaryZero = Render(Wrap(owner, "<mrow>" + zero + "</mrow>"));
        double implicitDelta = Formula(implicitSpaced).Drawing.Width - Formula(implicitZero).Drawing.Width;
        double ordinaryDelta = Formula(ordinarySpaced).Drawing.Width - Formula(ordinaryZero).Drawing.Width;

        Assert.Equal(10D, implicitDelta, 6);
        Assert.Equal(10D, ordinaryDelta, 6);
        Assert.Equal(Formula(ordinarySpaced).Drawing.Width, Formula(implicitSpaced).Drawing.Width, 6);
        Assert.Equal(ordinarySpaced.Text, implicitSpaced.Text);
        foreach (HtmlRenderDocument result in new[] { implicitSpaced, implicitZero, ordinarySpaced, ordinaryZero })
            AssertNoMathLoss(result);
    }

    private static HtmlRenderDocument Render(string content) {
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Continuous,
            ViewportWidth = 800D,
            ViewportHeight = 600D,
            DefaultFontFamily = "SpacingMath",
            DefaultFontSize = 20D,
            Margins = HtmlRenderMargins.All(0D),
            AllowSystemFontFallback = false
        };
        options.Fonts.Add("SpacingMath", ManagedTextShapingTestAssets.CreateFont(' ', '+', '1', '2', '√'));
        string html = "<!doctype html><html lang='en'><body style='margin:0'><math xmlns='http://www.w3.org/1998/Math/MathML' "
            + "style='font:20px SpacingMath'>" + content + "</math></body></html>";
        return HtmlRenderTestDriver.Render(html, options);
    }

    private static HtmlRenderDrawing Formula(HtmlRenderDocument rendered) =>
        Assert.Single(rendered.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderDrawing>());

    private static void AssertNoMathLoss(HtmlRenderDocument rendered) {
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.MathMlContentUnsupported);
        Assert.False(rendered.HasLoss);
    }

    private static string Wrap(string owner, string content) => owner switch {
        "mtd" => "<mtable><mtr><mtd>" + content + "</mtd></mtr></mtable>",
        "menclose" => "<menclose notation='box'>" + content + "</menclose>",
        _ => "<" + owner + ">" + content + "</" + owner + ">"
    };
}
