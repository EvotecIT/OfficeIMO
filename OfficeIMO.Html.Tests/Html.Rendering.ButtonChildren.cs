using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("block")]
    [InlineData("flex")]
    public void HtmlRendering_AutoWidthRichButtonKeepsStyledChildIntrinsicSize(string display) {
        string html = "<div style='display:flex'><button id='rich' style='display:" + display
            + ";padding:0 6px;border:1px solid #000'><span style='display:inline-block;width:48px;height:24px;background:red'></span>"
            + "<span hidden>hidden label</span></button><span>After</span></div>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html,
            new HtmlRenderOptions { ViewportWidth = 320D, Margins = HtmlRenderMargins.All(0D) });
        HtmlRenderShape button = Assert.Single(rendered.Pages.SelectMany(page => page.Visuals)
            .OfType<HtmlRenderShape>(), shape => shape.Source == "button#rich" && shape.Shape.FillColor.HasValue);
        Assert.Equal(62D, button.Width, 2);
        Assert.DoesNotContain("hidden label", rendered.Text, StringComparison.Ordinal);
        Assert.Contains("After", rendered.Text, StringComparison.Ordinal);
    }

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

    [Theory]
    [InlineData("")]
    [InlineData("display:inline-block")]
    [InlineData("display:inline-flex")]
    [InlineData("display:inline-grid")]
    public void HtmlRendering_ButtonChildrenHonorHiddenButtonInInlineAndIntrinsicLayout(string containerStyle) {
        string html = "<p><span style='" + containerStyle + "'>Before<button hidden><span>hidden label</span>"
            + "</button>After</span>Tail</p>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html);
        Assert.DoesNotContain("hidden label", rendered.Text, StringComparison.Ordinal);
        Assert.Contains("Before", rendered.Text, StringComparison.Ordinal);
        Assert.Contains("After", rendered.Text, StringComparison.Ordinal);
        Assert.Contains("Tail", rendered.Text, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("button", "")]
    [InlineData("button", "display:inline-block;")]
    [InlineData("span", "display:inline-block;")]
    public void HtmlRendering_ButtonChildrenAndInlineBlocksSizeStyledDescendants(string tag, string display) {
        string html = "<p><" + tag + " style='" + display + "font:16px Arial;padding:3px 9px;border:2px solid;margin-right:7px'>"
            + "<span style='font-size:48px;white-space:nowrap'>WIDE</span></" + tag + "><span>After</span></p>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions { ViewportWidth = 640D });
        HtmlRenderText[] text = rendered.Pages.SelectMany(p => EnumerateRenderVisuals(p.Scene)).OfType<HtmlRenderText>().ToArray();
        HtmlRenderText label = Assert.Single(text, t => t.Text == "WIDE");
        HtmlRenderText after = Assert.Single(text, t => t.Text == "After");
        Assert.Equal(48D, label.Font.Size, 3);
        Assert.True(after.X >= label.X + label.TextAdvanceWidth!.Value + 18D - 0.01D);
    }

    [Theory]
    [InlineData("", "height")]
    [InlineData("display:inline-block;", "height")]
    [InlineData("", "min-height")]
    [InlineData("display:inline-block;", "min-height")]
    public void HtmlRendering_ButtonChildrenRetainBlockAxisCentering(string display, string heightProperty) {
        string dimensions = "width:120px;" + heightProperty + ":80px";
        string html = "<p><button style='" + dimensions + "'>Save</button><button style='"
            + display + dimensions + "'><span>Save</span></button></p>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html);
        HtmlRenderText[] labels = rendered.Pages.SelectMany(p => EnumerateRenderVisuals(p.Scene))
            .OfType<HtmlRenderText>().Where(t => t.Text == "Save").ToArray();
        Assert.Equal(2, labels.Length);
        Assert.Equal(labels[0].Y, labels[1].Y, 3);
    }

    [Theory]
    [InlineData("inline-flex", "flex-start")]
    [InlineData("inline-grid", "start")]
    public void HtmlRendering_ButtonChildrenRetainAuthoredContainerAlignment(string display, string alignment) {
        string html = "<button id='rich' style='display:" + display + ";align-items:" + alignment
            + ";width:120px;height:80px'><span>Save</span></button>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions { Margins = HtmlRenderMargins.All(0D) });
        HtmlRenderText label = Assert.Single(rendered.Pages.SelectMany(p => EnumerateRenderVisuals(p.Scene)).OfType<HtmlRenderText>(), t => t.Text == "Save");
        Assert.Equal(5D, label.Y, 3);
    }
}
