using OfficeIMO.Drawing;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData(400, "bolder", 700, "attribute")]
    [InlineData(700, "lighter", 400, "attribute")]
    [InlineData(950, "bolder", 950, "attribute")]
    [InlineData(50, "lighter", 50, "attribute")]
    [InlineData(400, "bolder", 700, "inline")]
    [InlineData(700, "lighter", 400, "inline")]
    [InlineData(400, "bolder", 700, "stylesheet")]
    [InlineData(700, "lighter", 400, "stylesheet")]
    [InlineData(400, "bolder", 700, "variable")]
    [InlineData(700, "lighter", 400, "shorthand")]
    [InlineData(400, "bolder", 700, "variable-whitespace-attribute")]
    [InlineData(700, "lighter", 400, "variable-whitespace-inline")]
    public void SvgRelativeWeightsAreResolvedOnceBeforeNestedTextInheritance(
        int inherited, string relative, int expected, string source) {
        string attribute = source switch {
            "attribute" => $"font-weight='{relative}'",
            "variable-whitespace-attribute" => $"font-weight='var(--empty,) {relative}'",
            _ => ""
        };
        string inline = source switch {
            "inline" => $"font-weight:{relative}",
            "variable" => $"--weight:{relative};font-weight:var(--weight)",
            "shorthand" => $"--font:{relative} 16px Arial;font:var(--font)",
            "variable-whitespace-inline" => $"font-weight:var(--empty,) {relative}",
            _ => ""
        };
        string css = source == "stylesheet" ? $"<style>#weight {{ font-weight:{relative} }}</style>" : "";
        string html = css + $"<svg width='240' height='80' font-weight='{inherited}'>"
            + $"<g id='weight' {attribute} style='{inline}'><text x='2' y='35'>Outer"
            + "<tspan id='child'>Inner<tspan>Deep</tspan></tspan></text></g></svg>";

        var styles = HtmlComputedStyleEngine.Compute(html);
        Assert.Equal(expected.ToString(System.Globalization.CultureInfo.InvariantCulture),
            styles.Single(pair => pair.Key.Id == "child").Value.Properties["font-weight"]);
        OfficeDrawingText[] runs = InlineSvgTextRuns(html);
        Assert.Equal(new[] { "Outer", "Inner", "Deep" }, runs.Select(run => run.Text));
        Assert.All(runs, run => Assert.Equal(expected, run.Font.Face.Weight));
    }

    [Theory]
    [InlineData(false, "", 15D)]
    [InlineData(false, "inherit", 5D)]
    [InlineData(false, "initial", 15D)]
    [InlineData(false, "5px", 10D)]
    [InlineData(true, "", 15D)]
    [InlineData(true, "inherit", 5D)]
    [InlineData(true, "unset", 15D)]
    [InlineData(true, "5px", 10D)]
    public void SvgNestedBaselineShiftsRequireALocalDeclaration(bool css, string childShift, double expectedInnerY) {
        string outer = css ? "style='baseline-shift:10px'" : "baseline-shift='10px'";
        string inner = childShift.Length == 0 ? "" : css
            ? $"style='baseline-shift:{childShift}'" : $"baseline-shift='{childShift}'";
        string html = "<svg width='240' height='80'><text x='2' y='35' font-size='10'>"
            + $"<tspan {outer}>Outer<tspan id='child' {inner}>Inner<tspan>Deep</tspan></tspan>After</tspan></text></svg>";

        HtmlComputedStyle child = HtmlComputedStyleEngine.Compute(html).Single(pair => pair.Key.Id == "child").Value;
        if (childShift.Length == 0) Assert.False(child.Properties.ContainsKey("baseline-shift"));
        OfficeDrawingText[] runs = InlineSvgTextRuns(html);
        Assert.Equal(new[] { "Outer", "Inner", "Deep", "After" }, runs.Select(run => run.Text));
        Assert.Equal(new[] { 15D, expectedInnerY, expectedInnerY, 15D }, runs.Select(run => run.Y));
    }

    private static OfficeDrawingText[] InlineSvgTextRuns(string html) {
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html);
        return EnumerateRenderVisuals(rendered.Pages.SelectMany(page => page.Visuals))
            .OfType<HtmlRenderDrawing>().SelectMany(visual => DrawingTestTraversal.Elements(visual.Drawing))
            .OfType<OfficeDrawingText>().ToArray();
    }
}
