using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlRendering_FormControlsProduceVectorPaintAndSearchableStaticValues() {
        const string html = """
            <form>
              <label>Name <input id='name' value='Ada Lovelace'></label>
              <label>Lookup <input id='lookup' placeholder='Search records'></label>
              <label>Password <input id='password' type='password' value='never-export-this'></label>
              <label><input id='enabled' type='checkbox' checked> Enabled</label>
              <label><input id='preferred' type='radio' checked> Preferred</label>
              <select id='plan'><option>Basic</option><option selected>Premium plan</option></select>
              <textarea id='notes' rows='2'>Ready
            For review</textarea>
              <button id='export'>Export report</button>
              <progress id='progress' max='100' value='72'></progress>
              <meter id='meter' min='0' max='10' value='8'></meter>
              <input id='hidden-value' type='hidden' value='must-not-render'>
              <span hidden>also-must-not-render</span>
            </form>
            """;

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(
            html,
            new HtmlRenderOptions { ViewportWidth = 720D, Margins = HtmlRenderMargins.All(12D) });

        string text = rendered.Text;
        Assert.Contains("Ada Lovelace", text, StringComparison.Ordinal);
        Assert.Contains("Search records", text, StringComparison.Ordinal);
        Assert.Contains("Premium plan", text, StringComparison.Ordinal);
        Assert.Contains("Ready", text, StringComparison.Ordinal);
        Assert.Contains("For review", text, StringComparison.Ordinal);
        Assert.Contains("Export report", text, StringComparison.Ordinal);
        Assert.Contains("72%", text, StringComparison.Ordinal);
        Assert.Contains("80%", text, StringComparison.Ordinal);
        Assert.Contains(new string('*', "never-export-this".Length), text, StringComparison.Ordinal);
        Assert.DoesNotContain("never-export-this", text, StringComparison.Ordinal);
        Assert.DoesNotContain("must-not-render", text, StringComparison.Ordinal);
        Assert.DoesNotContain("also-must-not-render", text, StringComparison.Ordinal);

        HtmlRenderVisual[] visuals = rendered.Pages.SelectMany(page => page.Visuals).ToArray();
        Assert.Contains(visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "input#enabled:checked");
        Assert.Contains(visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "input#preferred:checked");
        Assert.Contains(visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "progress#progress:value");
        Assert.Contains(visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "meter#meter:value");
        Assert.All(
            visuals.OfType<HtmlRenderText>().Where(value => value.Source != null && value.Source.Contains('#')),
            value => Assert.Equal("form-control", value.SemanticRole));
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Severity == HtmlDiagnosticSeverity.Error);
    }

    [Fact]
    public void HtmlRendering_FormControlsRespectAuthoredDimensionsAndCommonInputKinds() {
        const string html = """
            <div>
              <input id='sized' value='Sized field' style='box-sizing:border-box;width:240px;height:40px'>
              <input id='range' type='range' min='0' max='10' value='6'>
              <input id='color' type='color' value='#ef476f'>
              <input id='file' type='file' value='report.pdf'>
              <select id='multiple' multiple size='3'>
                <option selected>North</option><option>South</option><option selected>West</option>
              </select>
              <input id='disabled' value='Disabled value' disabled>
            </div>
            """;

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(
            html,
            new HtmlRenderOptions { ViewportWidth = 800D, Margins = HtmlRenderMargins.All(8D) });
        HtmlRenderVisual[] visuals = rendered.Pages.SelectMany(page => page.Visuals).ToArray();

        HtmlRenderShape sizedBox = Assert.Single(
            visuals.OfType<HtmlRenderShape>(),
            shape => shape.Source == "input#sized"
                && shape.Shape.FillColor != null
                && Math.Abs(shape.Width - 240D) < 0.001D);
        Assert.Equal(40D, sizedBox.Height, 3);
        Assert.Contains(visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "input#range:thumb");
        Assert.Contains(visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "input#color:swatch");
        Assert.Contains("Choose file", rendered.Text, StringComparison.Ordinal);
        Assert.DoesNotContain("report.pdf", rendered.Text, StringComparison.Ordinal);
        Assert.Contains("North", rendered.Text, StringComparison.Ordinal);
        Assert.Contains("West", rendered.Text, StringComparison.Ordinal);
        Assert.Contains("Disabled value", rendered.Text, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlRendering_SizedSingleSelectWithoutSelectionRendersBlank() {
        const string html = "<select size='2'><option>First choice</option><option>Second choice</option></select>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions());

        Assert.DoesNotContain("First choice", rendered.Text, StringComparison.Ordinal);
        Assert.DoesNotContain("Second choice", rendered.Text, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlRendering_SingleSelectDefaultsToFirstEnabledOption() {
        const string html = "<select><option disabled>Disabled</option><optgroup disabled><option>Group disabled</option></optgroup><option label='Enabled label'>Fallback text</option></select>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions());

        Assert.Contains("Enabled label", rendered.Text, StringComparison.Ordinal);
        Assert.DoesNotContain("Disabled", rendered.Text, StringComparison.Ordinal);
        Assert.DoesNotContain("Fallback text", rendered.Text, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlRendering_UsesEffectiveInputTypesAndSanitizedRangePositions() {
        HtmlRenderDocument invalidType = HtmlRenderTestDriver.Render(
            "<input id='choice' type=' checkbox ' value='Text state' checked>",
            new HtmlRenderOptions());
        Assert.Contains("Text state", invalidType.Text, StringComparison.Ordinal);
        Assert.DoesNotContain(
            invalidType.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderShape>(),
            shape => shape.Source == "input#choice:checked");

        double defaultPosition = RangeThumbX("<input id='range' type='range'>");
        double explicitDefaultPosition = RangeThumbX("<input id='range' type='range' value='50'>");
        double clampedPosition = RangeThumbX("<input id='range' type='range' min='10' max='20' value='200'>");
        double maximumPosition = RangeThumbX("<input id='range' type='range' min='10' max='20' value='20'>");
        double extremeMidpointPosition = RangeThumbX("<input id='range' type='range' min='-1e308' max='1e308'>");
        double steppedPosition = RangeThumbX("<input id='range' type='range' min='0' max='10' step='3' value='8'>");
        double roundedPosition = RangeThumbX("<input id='range' type='range' min='0' max='10' step='3' value='9'>");

        Assert.Equal(explicitDefaultPosition, defaultPosition, 6);
        Assert.Equal(maximumPosition, clampedPosition, 6);
        Assert.Equal(explicitDefaultPosition, extremeMidpointPosition, 6);
        Assert.Equal(roundedPosition, steppedPosition, 6);

        HtmlRenderDocument radios = HtmlRenderTestDriver.Render(
            "<input id='first' type='radio' name='choice' checked><input id='last' type='radio' name='choice' checked>",
            new HtmlRenderOptions());
        HtmlRenderShape[] radioMarks = radios.Pages.SelectMany(page => page.Visuals)
            .OfType<HtmlRenderShape>()
            .Where(shape => shape.Source.EndsWith(":checked", StringComparison.Ordinal))
            .ToArray();
        Assert.DoesNotContain(radioMarks, shape => shape.Source == "input#first:checked");
        Assert.Contains(radioMarks, shape => shape.Source == "input#last:checked");
    }

    [Fact]
    public void HtmlRendering_DoesNotRenderInapplicablePlaceholdersForTypedInputs() {
        const string html = """
            <input type='date' value='invalid' placeholder='Birthday'>
            <input type='number' value='invalid' placeholder='Quantity'>
            <input type='text' placeholder='Search records'>
            """;

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions());

        Assert.DoesNotContain("Birthday", rendered.Text, StringComparison.Ordinal);
        Assert.DoesNotContain("Quantity", rendered.Text, StringComparison.Ordinal);
        Assert.Contains("Search records", rendered.Text, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlPdf_FormControlSnapshotRemainsSearchableWithoutExposingPasswordValues() {
        const string html = """
            <h1>Approval</h1>
            <label>Owner <input value='Grace Hopper'></label>
            <select><option selected>Approved</option></select>
            <input type='password' value='classified-value'>
            <progress max='100' value='64'></progress>
            <button>Archive record</button>
            """;

        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdf(new HtmlPdfSaveOptions());
        string text = PdfCore.PdfReadDocument.Open(pdf).ExtractText();

        Assert.Contains("Grace Hopper", text, StringComparison.Ordinal);
        Assert.Contains("Approved", text, StringComparison.Ordinal);
        Assert.Contains("64%", text, StringComparison.Ordinal);
        Assert.Contains("Archive record", text, StringComparison.Ordinal);
        Assert.DoesNotContain("classified-value", text, StringComparison.Ordinal);
        Assert.True(PdfCore.PdfInspector.Inspect(pdf).PageCount >= 1);
    }

    [Fact]
    public void HtmlRendering_ImageInputUsesItsImageSourceAndAlternativeText() {
        const string pixelPng = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNgYAAAAAMAASsJTYQAAAAASUVORK5CYII=";
        string html =
            "<form><input id='save-image' type='image' src='data:image/png;base64," +
            pixelPng +
            "' alt='Save changes' style='width:36px;height:24px'></form>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(
            html,
            new HtmlRenderOptions { ViewportWidth = 200D, Margins = HtmlRenderMargins.All(8D) });
        HtmlRenderImage image = Assert.Single(
            rendered.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderImage>());

        Assert.Equal("Save changes", image.AlternativeText);
        Assert.Equal(36D, image.Width, 3);
        Assert.Equal(24D, image.Height, 3);
        Assert.DoesNotContain("Button", rendered.Text, StringComparison.Ordinal);
    }

    private static double RangeThumbX(string html) {
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions());
        return Assert.Single(
            rendered.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderShape>(),
            shape => shape.Source == "input#range:thumb").X;
    }
}
