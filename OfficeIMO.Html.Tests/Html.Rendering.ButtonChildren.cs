using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("")]
    [InlineData("inline-block")]
    [InlineData("inline-flex")]
    [InlineData("inline-grid")]
    public void HtmlRendering_ButtonChildrenKeepTheirLayoutStylesAndVisibility(string display) {
        string displayDeclaration = display.Length == 0 ? string.Empty : "display:" + display + ";";
        string html = "<div style='font:16px Arial'><button id='rich' style='" + displayDeclaration
            + "width:132px;flex-direction:column;text-align:left'>"
            + "<strong style='display:block;font-size:16px;font-weight:bold'>Primary</strong>"
            + "<small style='display:block;font-size:10px'>domain controllers</small>"
            + "<span style='display:none'>hidden detail</span></button><span>After</span></div>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions { ViewportWidth = 320D });
        HtmlRenderText[] text = rendered.Pages.SelectMany(p => EnumerateRenderVisuals(p.Scene)).OfType<HtmlRenderText>().ToArray();
        HtmlRenderText title = Assert.Single(text, t => t.Text == "Primary");
        HtmlRenderText subtitle = Assert.Single(text, t => t.Text == "domain controllers");
        Assert.Equal(16D, title.Font.Size, 3);
        Assert.True((title.Font.Style & OfficeFontStyle.Bold) != 0);
        Assert.Equal(10D, subtitle.Font.Size, 3);
        Assert.True(subtitle.Y >= title.Y + title.Height - 0.01D);
        Assert.DoesNotContain("hidden detail", rendered.Text, StringComparison.Ordinal);
        Assert.Contains("After", rendered.Text, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlRendering_ButtonChildLayoutHonorsTransparentBackgroundAndZeroPadding() {
        const string html = "<button id='rich' style='width:100px;border:0;padding:0;background:transparent;text-align:left'>"
            + "<span>Visible label</span></button>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions { ViewportWidth = 200D, Margins = HtmlRenderMargins.All(0D) });
        HtmlRenderVisual[] visuals = rendered.Pages.SelectMany(p => EnumerateRenderVisuals(p.Scene)).ToArray();
        Assert.DoesNotContain(visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "button#rich" && shape.Shape.FillColor != null);
        HtmlRenderText label = Assert.Single(visuals.OfType<HtmlRenderText>(), t => t.Text == "Visible label");
        Assert.Equal(0D, label.X, 3);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void HtmlPdf_ButtonChildrenPreserveStaticContentAlongsideInputSnapshots(bool interactive) {
        const string html = "<button style='display:inline-flex;flex-direction:column;width:132px'>"
            + "<strong>Primary</strong><small style='font-size:10px'>domain controllers</small>"
            + "<span hidden>omitted label</span></button><input value='Actual input'>";
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdfBytes(new HtmlToPdfOptions { InteractiveFormControls = interactive });
        string text = OfficeIMO.Pdf.PdfReadDocument.Open(pdf).ExtractText();
        Assert.Contains("Primary", text, StringComparison.Ordinal);
        Assert.Contains("domain controllers", text, StringComparison.Ordinal);
        Assert.Contains("Actual input", text, StringComparison.Ordinal);
        Assert.DoesNotContain("omitted label", text, StringComparison.Ordinal);
    }
}
