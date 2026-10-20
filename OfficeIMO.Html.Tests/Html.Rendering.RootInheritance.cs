using OfficeIMO.Drawing;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData(HtmlRenderMode.Continuous, false)]
    [InlineData(HtmlRenderMode.Paged, false)]
    [InlineData(HtmlRenderMode.Continuous, true)]
    [InlineData(HtmlRenderMode.Paged, true)]
    public void HtmlRendering_BodyRootUsesHtmlTypographyAndRetainsAuthoredOverrides(HtmlRenderMode mode, bool bodyOverride) {
        string html = "<style>html{--report-size:11px;--report-color:#123456;"
            + "font-family:Courier;font-size:var(--report-size);line-height:1.3;"
            + "font-weight:bold;font-style:italic;color:var(--report-color)}body,p{margin:0}"
            + (bodyOverride ? "body{font-family:Arial;font-size:150%;line-height:1.5;font-weight:normal;font-style:normal;color:#654321}" : string.Empty)
            + "</style><p>RootTypography</p>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions {
            Mode = mode,
            DefaultFontSize = 18D,
            ViewportWidth = 400D,
            ViewportHeight = 200D,
            Margins = HtmlRenderMargins.All(0D)
        });

        HtmlRenderText text = Assert.Single(rendered.Pages.SelectMany(page => EnumerateRenderVisuals(page.Scene))
            .OfType<HtmlRenderText>(), item => item.Text == "RootTypography");
        Assert.Equal(bodyOverride ? 16.5D : 11D, text.Font.Size, 6);
        Assert.Equal(bodyOverride ? 24.75D : 14.3D, text.LineHeight, 6);
        Assert.Equal(bodyOverride ? "Arial" : "Courier", text.Font.FamilyName);
        Assert.Equal(bodyOverride ? 400 : 700, text.FontDescriptor.Weight);
        Assert.Equal(bodyOverride ? OfficeFontSlant.Normal : OfficeFontSlant.Italic, text.FontDescriptor.Slant);
        Assert.Equal(bodyOverride ? OfficeColor.FromRgb(101, 67, 33) : OfficeColor.FromRgb(18, 52, 86), text.Color);
    }

    [Theory]
    [InlineData(HtmlRenderMode.Continuous)]
    [InlineData(HtmlRenderMode.Paged)]
    public void HtmlRendering_BodyRootInheritsComputedPercentageLineHeightAcrossFontChanges(HtmlRenderMode mode) {
        const string html = "<style>html{font-size:20px;line-height:150%}body{margin:0;font-size:50%}"
            + "p{margin:0;font-size:2em}</style><p>RootLineHeight</p>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions {
            Mode = mode,
            ViewportWidth = 400D,
            ViewportHeight = 200D,
            Margins = HtmlRenderMargins.All(0D)
        });

        HtmlRenderText text = Assert.Single(rendered.Pages.SelectMany(page => EnumerateRenderVisuals(page.Scene))
            .OfType<HtmlRenderText>(), item => item.Text == "RootLineHeight");
        Assert.Equal(20D, text.Font.Size, 6);
        Assert.Equal(30D, text.LineHeight, 6);
    }
}
