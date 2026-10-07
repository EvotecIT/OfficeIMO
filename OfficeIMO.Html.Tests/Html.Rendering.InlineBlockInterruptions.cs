using OfficeIMO.Drawing;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlInlineBlockInterruptionTests {
    [Theory]
    [InlineData("map", false)]
    [InlineData("span", false)]
    [InlineData("a", false)]
    [InlineData("span", true)]
    public void BlockDescendant_InterruptsInlineFlowAndRetainsItsBox(string tag, bool nested) {
        string child = "Lead<span id='child' style='display:block;border-left:6px solid green;padding-left:10px;height:40px'>Inside</span>Tail";
        if (nested) child = "<em>" + child + "</em>";
        string html = $"<style>*{{margin:0}}body{{font-size:16px;line-height:20px}}</style><div>Before<{tag}>{child}</{tag}>After</div>";
        HtmlRenderDocument result = HtmlRenderEngine.Render(HtmlConversionDocument.Parse(html), Options());
        var visuals = result.Pages.SelectMany(page => page.Visuals).ToList();
        HtmlRenderText Text(string text) => Assert.Single(visuals.OfType<HtmlRenderText>(), value => value.Text == text);
        Assert.Equal(Text("Before").Y, Text("Lead").Y, 3);
        Assert.Equal(20D, Text("Inside").Y, 3);
        Assert.Equal(16D, Text("Inside").X, 3);
        Assert.Equal(60D, Text("Tail").Y, 3);
        Assert.Equal(Text("Tail").Y, Text("After").Y, 3);
        HtmlRenderShape border = Assert.Single(visuals.OfType<HtmlRenderShape>(), value => value.Source == "span#child:border-left");
        Assert.Equal(40D, border.Height, 3);
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(result.Pages[0].CreateDrawing());
        Assert.Equal(OfficeColor.Green, raster.GetPixel(2, 35));
        Assert.NotEqual(OfficeColor.Green, raster.GetPixel(8, 35));
    }

    [Fact]
    public void BlockDescendant_InheritsRelativePaintOffsetAndAncestorLink() {
        const string html = "<style>*{margin:0}body{font-size:16px;line-height:20px}</style>"
            + "<div>Before<a href='https://example.org/' style='position:relative;left:7px;top:5px'>Lead"
            + "<span id='child' style='display:block;height:40px'>Inside</span>Tail</a>After</div>";
        HtmlRenderDocument result = HtmlRenderEngine.Render(HtmlConversionDocument.Parse(html), Options());
        var visuals = result.Pages.SelectMany(page => page.Visuals).ToList();
        HtmlRenderText inside = Assert.Single(visuals.OfType<HtmlRenderText>(), value => value.Text == "Inside");
        Assert.Equal(7D, inside.X, 3);
        Assert.Equal(25D, inside.Y, 3);
        Assert.Contains(visuals, visual => visual.LinkUri == "https://example.org/" && visual.Y == 25D && visual.Height >= 40D);
        Assert.Equal(60D, Assert.Single(visuals.OfType<HtmlRenderText>(), value => value.Text == "After").Y, 3);
    }

    [Fact]
    public void BlockDescendant_PreservesForcedPageBreak() {
        const string html = "<style>*{margin:0}body{font-size:16px;line-height:20px}</style>"
            + "<div>Before<span>Lead<span style='display:block;break-before:page;height:40px'>Inside</span>Tail</span>After</div>";
        HtmlRenderOptions options = Options();
        options.Mode = HtmlRenderMode.Paged;
        options.PageSize = new OfficePageSize(640D / 96D, 200D / 96D);
        HtmlRenderDocument result = HtmlRenderEngine.Render(HtmlConversionDocument.Parse(html), options);
        Assert.Equal(2, result.Pages.Count);
        Assert.Contains(result.Pages[0].Visuals.OfType<HtmlRenderText>(), text => text.Text == "Lead");
        Assert.DoesNotContain(result.Pages[0].Visuals.OfType<HtmlRenderText>(), text => text.Text == "Inside");
        Assert.Contains(result.Pages[1].Visuals.OfType<HtmlRenderText>(), text => text.Text == "Inside");
        Assert.Contains(result.Pages[1].Visuals.OfType<HtmlRenderText>(), text => text.Text == "After");
    }

    [Theory]
    [InlineData("")]
    [InlineData(" \n  ")]
    public void ConsecutiveBlockDescendants_CollapseAdjoiningMargins(string separator) {
        string html = "<style>*{margin:0}body{font-size:16px;line-height:20px}</style>"
            + "<div><span><span style='display:block;height:20px;margin-bottom:30px'>First</span>" + separator
            + "<span style='display:block;height:20px;margin-top:20px'>Second</span></span></div>";
        HtmlRenderDocument result = HtmlRenderEngine.Render(HtmlConversionDocument.Parse(html), Options());
        HtmlRenderText second = Assert.Single(result.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderText>(), text => text.Text == "Second");
        Assert.Equal(50D, second.Y, 3);
    }

    [Fact]
    public void BlockDescendant_RetainsAncestorOpacity() {
        const string html = "<style>*{margin:0}</style><div><span style='opacity:0.5'>"
            + "<span style='display:block;width:40px;height:40px;background:red'></span></span></div>";
        HtmlRenderDocument result = HtmlRenderEngine.Render(HtmlConversionDocument.Parse(html), Options());
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(result.Pages[0].CreateDrawing());
        OfficeColor pixel = raster.GetPixel(10, 10);
        Assert.Equal((byte)255, pixel.R);
        Assert.InRange(pixel.G, (byte)127, (byte)128);
        Assert.InRange(pixel.B, (byte)127, (byte)128);
    }

    [Fact]
    public void AncestorOpacity_CompositesOverlappingInlineAndBlockContentOnce() {
        const string html = "<style>*{margin:0}body{line-height:20px}</style><div><span id='group' style='opacity:.5'>"
            + "<span style='display:inline-block;width:40px;height:20px;background:blue'></span>"
            + "<span style='display:block;position:relative;top:-20px;width:40px;height:40px;background:red'></span>"
            + "</span></div>";
        HtmlRenderOptions options = Options();
        options.BackgroundColor = OfficeColor.Transparent;
        HtmlRenderDocument result = HtmlRenderEngine.Render(HtmlConversionDocument.Parse(html), options);
        HtmlRenderEffectGroup group = Assert.Single(result.Pages[0].Visuals.OfType<HtmlRenderEffectGroup>(), item => item.Source == "span#group");
        Assert.Equal(.5D, group.Opacity);
        OfficeColor pixel = OfficeDrawingRasterRenderer.Render(result.Pages[0].CreateDrawing()).GetPixel(10, 10);
        Assert.Equal((byte)255, pixel.R);
        Assert.Equal((byte)0, pixel.B);
        Assert.InRange(pixel.A, (byte)127, (byte)128);
    }

    [Fact]
    public void InlineBackground_DoesNotPaintAcrossInterveningBlock() {
        const string html = "<style>*{margin:0}body{font-size:16px;line-height:20px}</style><div>"
            + "<span style='background:green'>Lead<span style='display:block;height:40px'></span>Tail</span></div>";
        HtmlRenderOptions options = Options();
        options.BackgroundColor = OfficeColor.Transparent;
        HtmlRenderDocument result = HtmlRenderEngine.Render(HtmlConversionDocument.Parse(html), options);
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(result.Pages[0].CreateDrawing());
        Assert.Equal(OfficeColor.Transparent, raster.GetPixel(10, 30));
        Assert.Equal(OfficeColor.Green, raster.GetPixel(1, 1));
        Assert.Equal(OfficeColor.Green, raster.GetPixel(1, 61));
    }

    [Fact]
    public void BlockDescendant_InsideListRetainsForcedPageBreak() {
        const string html = "<style>*{margin:0}body{font-size:16px;line-height:20px}</style>"
            + "<ul><li><span>Lead<span style='display:block;break-before:page;height:40px'>Inside</span>Tail</span></li></ul>";
        HtmlRenderOptions options = Options();
        options.Mode = HtmlRenderMode.Paged;
        options.PageSize = new OfficePageSize(640D / 96D, 200D / 96D);
        HtmlRenderDocument result = HtmlRenderEngine.Render(HtmlConversionDocument.Parse(html), options);
        Assert.Equal(2, result.Pages.Count);
        Assert.Contains(result.Pages[1].Visuals.OfType<HtmlRenderText>(), text => text.Text == "Inside");
    }

    [Fact]
    public void EmptyInlineWrapper_PreservesOuterMarginCollapse() {
        const string prefix = "<style>*{margin:0}body{font-size:16px;line-height:20px}</style>"
            + "<div style='margin-bottom:40px'>Before</div><div>";
        const string block = "<span style='display:block;margin-top:20px;margin-bottom:30px'>Inside</span>";
        const string suffix = "</div><div style='margin-top:10px'>After</div>";
        HtmlRenderDocument direct = HtmlRenderEngine.Render(HtmlConversionDocument.Parse(prefix + block + suffix), Options());
        HtmlRenderDocument wrapped = HtmlRenderEngine.Render(HtmlConversionDocument.Parse(prefix + "<span>" + block + "</span>" + suffix), Options());
        foreach (string text in new[] { "Inside", "After" }) {
            double expected = Assert.Single(direct.Pages[0].Visuals.OfType<HtmlRenderText>(), item => item.Text == text).Y;
            Assert.Equal(expected, Assert.Single(wrapped.Pages[0].Visuals.OfType<HtmlRenderText>(), item => item.Text == text).Y, 3);
        }
    }

    [Theory]
    [InlineData("vertical-rl")]
    [InlineData("vertical-lr")]
    public void TransparentInlineWrapper_PreservesVerticalBlockFlow(string writingMode) {
        string prefix = "<style>*{margin:0}body{font-size:16px;line-height:20px}</style>"
            + $"<div style='writing-mode:{writingMode};width:200px;height:180px'>";
        const string content = "Lead<span style='display:block'>Inside</span>Tail";
        HtmlRenderDocument direct = HtmlRenderEngine.Render(HtmlConversionDocument.Parse(prefix + content + "</div>"), Options());
        HtmlRenderDocument wrapped = HtmlRenderEngine.Render(HtmlConversionDocument.Parse(prefix + "<span>" + content + "</span></div>"), Options());
        Assert.Equal(
            OfficeDrawingRasterRenderer.ToPng(direct.Pages[0].CreateDrawing()),
            OfficeDrawingRasterRenderer.ToPng(wrapped.Pages[0].CreateDrawing()));
    }

    [Theory]
    [InlineData("<div style='position:absolute;width:20px;height:20px;background:red'></div>")]
    [InlineData("<div style='display:none;float:left'>Hidden</div>")]
    public void ContainerWithoutInFlowBlocks_RemainsRenderable(string child) {
        HtmlRenderDocument result = HtmlRenderEngine.Render(
            HtmlConversionDocument.Parse("<div style='position:relative'>" + child + "</div><p>Following</p>"), Options());
        Assert.Contains(result.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderText>(), text => text.Text == "Following");
    }

    private static HtmlRenderOptions Options() => new() {
        ViewportWidth = 640D, ViewportHeight = 200D, Margins = HtmlRenderMargins.All(0D)
    };
}
