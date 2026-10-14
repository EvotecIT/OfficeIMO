using System.Globalization;
using System.Text;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlUsedLayoutHeightTests {
    // Byte-backed native controls in used-heights.html establish these used
    // dimensions. They qualify this HTML layout policy, not a universal CSS unit.
    private const double UsedParagraphLineHeight = 23.09375D;
    private const double UsedImageHeight = 53.9375D;
    private const string Tail = "<div id='tail' style='height:4px;background:#d06c24'></div>";
    private const string ImageSvg = "<svg xmlns='http://www.w3.org/2000/svg' width='128' height='64' "
        + "viewBox='0 0 128 64'><rect width='128' height='64' fill='#2463a0'/></svg>";

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void WrappedLinesAndFollowingFlowShareTheUsedHeight(bool floated) {
        string content = "<div class='box' id='subject' style='width:" + (floated ? "70" : "50") + "px'>"
            + (floated ? "<i style='float:left;width:20px;height:50px'></i>" : "")
            + "Hgyp Hgyp Hgyp</div>" + Tail;
        HtmlRenderDocument rendered = Render(Source(content));
        HtmlRenderText[] text = Visuals(rendered).OfType<HtmlRenderText>().OrderBy(item => item.Y).ToArray();

        Assert.Equal(3, text.Length);
        Assert.All(text, item => Assert.Equal("Hgyp", item.Text.Trim()));
        Assert.Equal(3D * UsedParagraphLineHeight, Shape(rendered, "div#subject").Height, 6);
        Assert.Equal(3D * UsedParagraphLineHeight, Shape(rendered, "div#tail").Y, 6);
        for (int index = 0; index < text.Length; index++) {
            Assert.Equal(16D + index * UsedParagraphLineHeight, text[index].Y + text[index].Font.Size, 6);
            Assert.Equal(floated ? 20D : 0D, text[index].X, 6);
        }
    }

    [Fact]
    public void NumericMultiplierUsesTheSameHeightForItsBoxAndFollowingBlock() {
        HtmlRenderDocument rendered = Render(Source(
            "<div class='box' id='subject' style='font-size:16px;line-height:1.2'>Hgyp</div>" + Tail));

        Assert.Equal(19.1875D, Shape(rendered, "div#subject").Height, 6);
        Assert.Equal(19.1875D, Shape(rendered, "div#tail").Y, 6);
        HtmlRenderText text = Assert.Single(Visuals(rendered).OfType<HtmlRenderText>());
        Assert.Equal(16D, text.Font.Size, 6);
        Assert.Equal(15D, text.Y + text.Font.Size, 6);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void EmptyCellAndCaptionDoNotCreateAnUnrequestedLine(bool caption, bool explicitZeroMinimum) {
        string minimum = explicitZeroMinimum ? " style='min-height:0'" : "";
        string content = "<div class='box' id='empty'></div><table id='subject'" + minimum + ">"
            + (caption ? "<caption id='caption'" + minimum + "></caption>" : "")
            + "<tbody><tr><td id='cell'" + minimum + "></td></tr></tbody></table>" + Tail;
        HtmlRenderDocument rendered = Render(Source(content));

        // The native empty ordinary/table/caption/row/cell boxes all have zero
        // height. A painted table grid is separate from its caption wrapper.
        Assert.Equal(0D, Shape(rendered, "div#tail").Y, 6);
        Assert.Empty(Visuals(rendered).OfType<HtmlRenderText>());
    }

    [Fact]
    public void PopulatedCaptionAndCellEachContributeOneUsedLine() {
        HtmlRenderDocument rendered = Render(Source("<table id='subject'><caption id='caption'>Hgyp</caption>"
            + "<tbody><tr><td id='cell'>Hgyp</td></tr></tbody></table>" + Tail));
        HtmlRenderShape caption = Shape(rendered, "caption#caption");
        HtmlRenderShape cell = Shape(rendered, "td#cell");
        HtmlRenderShape grid = Shape(rendered, "table#subject");

        Assert.Equal(UsedParagraphLineHeight, caption.Height, 6);
        Assert.Equal(UsedParagraphLineHeight, cell.Height, 6);
        Assert.Equal(caption.Y + caption.Height, cell.Y, 6);
        Assert.Equal(cell.Y, grid.Y, 6);
        Assert.Equal(cell.Height, grid.Height, 6);
        Assert.Equal(2D * UsedParagraphLineHeight, Shape(rendered, "div#tail").Y, 6);
        Assert.Equal(2, Visuals(rendered).OfType<HtmlRenderText>().Count());
    }

    [Theory]
    [InlineData("1.65", 33D)]
    [InlineData("23.1px", UsedParagraphLineHeight)]
    [InlineData("165%", UsedParagraphLineHeight)]
    public void InheritedNumberRescalesButLengthAndPercentageKeepTheComputedLength(
        string parentLineHeight, double expectedHeight) {
        string content = "<div style='font-size:14px;line-height:" + parentLineHeight + "'>"
            + "<div class='box' id='subject' style='font-size:20px'>Hgyp</div></div>" + Tail;
        HtmlRenderDocument rendered = Render(Source(content));

        Assert.Equal(expectedHeight, Shape(rendered, "div#subject").Height, 6);
        Assert.Equal(expectedHeight, Shape(rendered, "div#tail").Y, 6);
        Assert.Equal(20D, Assert.Single(Visuals(rendered).OfType<HtmlRenderText>()).Font.Size, 6);
    }

    [Theory]
    [InlineData("width:64px;height:53.94px", 64D)]
    [InlineData("width:64px;height:53.949px", 64D)]
    [InlineData("width:107.88px;height:auto", 107.875D)]
    [InlineData("width:107.906px;height:auto", 107.890625D)]
    public void ReplacedUsedSizeControlsRatioNeighborPlacementAndFollowingFlow(string dimensions, double expectedWidth) {
        string imageSource = "data:image/svg+xml;base64," + Convert.ToBase64String(Encoding.UTF8.GetBytes(ImageSvg));
        string content = "<div class='box' id='subject' style='width:250px'><img id='image' style='"
            + dimensions + "' alt='' src='" + imageSource + "'><span>Hgyp</span></div>" + Tail;
        HtmlRenderDocument rendered = Render(Source(content));
        HtmlRenderVisual image = Assert.Single(Visuals(rendered), visual =>
            visual.Source == "img#image" && visual is HtmlRenderImage or HtmlRenderDrawing);
        HtmlRenderText neighbor = Assert.Single(Visuals(rendered).OfType<HtmlRenderText>());

        Assert.Equal(expectedWidth, image.Width, 6);
        Assert.Equal(UsedImageHeight, image.Height, 6);
        Assert.Equal(0D, image.Y, 6);
        Assert.Equal(expectedWidth, neighbor.X, 6);
        Assert.Equal(image.Y + image.Height, neighbor.Y + neighbor.Font.Size, 6);
        Assert.Equal(61.03125D, Shape(rendered, "div#subject").Height, 6);
        Assert.Equal(61.03125D, Shape(rendered, "div#tail").Y, 6);
    }

    [Theory]
    [InlineData(1.65D)]
    [InlineData(2D)]
    public void CustomDefaultMultiplierSurvivesNormalInheritanceAndPdfOptionsCopy(double multiplier) {
        HtmlRenderOptions options = Options(multiplier);
        const string content = "<div class='box' id='subject' style='font-size:20px'>Hgyp</div>";
        string number = multiplier.ToString(CultureInfo.InvariantCulture);
        HtmlRenderDocument reference = Render(Source(content + Tail, number), options);

        foreach (string? lineHeight in new string?[] { null, "normal" }) {
            HtmlConversionDocument document = HtmlConversionDocument.Parse(Source(content + Tail, lineHeight));
            HtmlRenderDocument actual = HtmlRenderTestDriver.Render(document, options);
            HtmlRenderDocument pdfScene = document.RenderToPdfResult(HtmlRenderRequest.Create(
                HtmlRenderIntentProfile.PrintPaged, HtmlRenderEncoder.Pdf, new HtmlToPdfOptions(options)))
                .RenderResult.Document;

            Assert.Equal(Shape(reference, "div#subject").Height, Shape(actual, "div#subject").Height, 6);
            Assert.Equal(Shape(reference, "div#tail").Y, Shape(actual, "div#tail").Y, 6);
            Assert.Equal(Shape(actual, "div#tail").Y, Shape(pdfScene, "div#tail").Y, 6);
            HtmlRenderText text = Assert.Single(Visuals(actual).OfType<HtmlRenderText>());
            HtmlRenderText pdfText = Assert.Single(Visuals(pdfScene).OfType<HtmlRenderText>());
            Assert.Equal(20D, text.Font.Size, 6);
            Assert.Equal(text.Y + text.Font.Size, pdfText.Y + pdfText.Font.Size, 6);
        }
    }

    [Theory]
    [InlineData("1.2")]
    [InlineData("23.1px")]
    [InlineData("165%")]
    public void AuthoredLineHeightKeepsPrecedenceOverTheCustomDefault(string lineHeight) {
        string html = Source("<div class='box' id='subject' style='font-size:20px'>Hgyp</div>" + Tail, lineHeight);
        HtmlRenderDocument ordinaryDefault = Render(html, Options());
        HtmlRenderDocument customDefault = Render(html, Options(2D));

        Assert.Equal(Shape(ordinaryDefault, "div#subject").Height, Shape(customDefault, "div#subject").Height, 6);
        Assert.Equal(Shape(ordinaryDefault, "div#tail").Y, Shape(customDefault, "div#tail").Y, 6);
    }

    private static string Source(string content, string? lineHeight = "1.65") =>
        "<!doctype html><style>body{margin:0;padding:0;font-family:Pinned;font-size:14px;font-weight:400;"
        + (lineHeight == null ? "" : "line-height:" + lineHeight + ";")
        + "}.box,td,caption{background:#eee}table{border-collapse:collapse;border-spacing:0;"
        + "margin:0;padding:0;width:90px;background:#eee}td,caption{border:0;padding:0;margin:0;"
        + "vertical-align:baseline}img{vertical-align:baseline;margin:0;padding:0;border:0}</style>" + content;

    private static HtmlRenderOptions Options(double defaultLineHeight = 1.2D) {
        var options = new HtmlRenderOptions {
            ViewportWidth = 400D, ViewportHeight = 300D,
            Margins = HtmlRenderMargins.All(0D),
            UserAgentStyles = HtmlRenderUserAgentStyleMode.Browser,
            DefaultLineHeight = defaultLineHeight
        };
        options.Fonts.Add("Pinned", File.ReadAllBytes(Path.Combine(RepositoryTestPaths.Find(),
            "OfficeIMO.TestAssets", "Fonts", "OfficeIMOBaselineSans-Regular.ttf")));
        return options;
    }

    private static HtmlRenderDocument Render(string html, HtmlRenderOptions? options = null) =>
        HtmlRenderTestDriver.Render(html, options ?? Options());

    private static HtmlRenderShape Shape(HtmlRenderDocument document, string source) =>
        Assert.Single(Visuals(document).OfType<HtmlRenderShape>(), shape =>
            shape.Source == source && shape.Shape.FillColor.HasValue);

    private static IEnumerable<HtmlRenderVisual> Visuals(HtmlRenderDocument document) =>
        document.Pages.SelectMany(page => Walk(page.Scene));

    private static IEnumerable<HtmlRenderVisual> Walk(IEnumerable<HtmlRenderVisual> visuals) {
        foreach (HtmlRenderVisual visual in visuals) {
            yield return visual;
            IEnumerable<HtmlRenderVisual>? children = visual switch {
                HtmlRenderClipGroup group => group.Visuals,
                HtmlRenderPathClipGroup group => group.Visuals,
                HtmlRenderSemanticGroup group => group.Visuals,
                HtmlRenderEffectGroup group => group.Visuals,
                HtmlRenderLogicalTextGroup group => group.Visuals,
                _ => null
            };
            if (children != null) foreach (HtmlRenderVisual child in Walk(children)) yield return child;
        }
    }
}
