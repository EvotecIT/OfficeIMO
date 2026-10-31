using System.Xml.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("<br>", "", false, 120D)]
    [InlineData("<br>", "", true, 120D)]
    [InlineData("\n", "white-space:pre-wrap", false, 120D)]
    [InlineData("\n", "white-space:pre-wrap", true, 120D)]
    [InlineData("\n", "white-space:pre", false, 120D)]
    [InlineData("\n", "white-space:pre", true, 120D)]
    [InlineData("<br><br>", "", false, 180D)]
    [InlineData("<br><br>", "", true, 180D)]
    public void HtmlForcedBreak_InlineStrutParticipatesOnTheCommittedAndEmptyLines(
        string separator, string whitespace, bool floated, double expectedHeight) {
        // Native Baseline Sans controls put the first baseline at 43px in a
        // 60px line, including a break before the large inline's first glyph.
        string html = ForcedBreakSource(separator + "AFTER", whitespace, floated);
        HtmlConversionDocument document = HtmlConversionDocument.Parse(html);
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(document, ForcedBreakOptions());
        HtmlRenderText[] texts = InlineEdgeTexts(rendered);

        Assert.Equal(new[] { "BEFORE", "AFTER", "FOLLOW" }, texts.Select(text => text.Text));
        Assert.Equal(67D, texts[0].Y + texts[0].Font.Size, 6);
        Assert.Equal(24D + expectedHeight - 17D, texts[1].Y + texts[1].Font.Size, 6);
        Assert.Equal(floated ? 44D : 24D, texts[0].X, 6);
        Assert.Equal(floated && expectedHeight == 120D ? 44D : 24D, texts[1].X, 6);
        Assert.Equal(expectedHeight, TableGeometryShape(rendered, "div#case").Height, 6);
        Assert.Equal(24D + expectedHeight, texts[2].LayoutY, 6);

        using var pdf = UglyToad.PdfPig.PdfDocument.Open((document.ToPdf(
            new HtmlToPdfOptions(ForcedBreakOptions()))).ToBytes());
        var page = pdf.GetPage(1);
        Assert.Equal("BEFOREAFTERFOLLOW", string.Concat(page.Letters.Select(letter => letter.Value)));
        Assert.Equal(67D, (page.Height - page.Letters.First().StartBaseLine.Y) / 0.75D, 3);
        Assert.Equal(24D + expectedHeight - 17D,
            (page.Height - page.Letters.First(letter => letter.Value == "A").StartBaseLine.Y) / 0.75D, 3);
    }

    [Theory]
    [InlineData("<br>", 60D)]
    [InlineData("INNER<br>AFTER", 120D)]
    public void HtmlForcedBreak_TrailingAndAlreadyParticipatingStrutsKeepTheirLineHeight(
        string content, double expectedHeight) {
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(ForcedBreakSource(content, "", false),
            ForcedBreakOptions());

        Assert.Equal(expectedHeight, TableGeometryShape(rendered, "div#case").Height, 6);
        HtmlRenderText[] texts = InlineEdgeTexts(rendered);
        Assert.Equal(67D, texts[0].Y + texts[0].Font.Size, 6);
        Assert.Equal(24D + expectedHeight, texts.Last().LayoutY, 6);
        Assert.DoesNotContain(texts, text => text.Text.Length == 0 || text.Text.Contains('\u2028'));
    }

    [Theory]
    [InlineData("slice", false, 24D)]
    [InlineData("clone", false, 35D)]
    [InlineData("slice", true, 44D)]
    [InlineData("clone", true, 55D)]
    public void HtmlForcedBreak_EmptyBreakFragmentConsumesOnlyAuthoredInlineEdges(
        string decorationBreak, bool floated, double expectedAfterX) {
        string html = ForcedBreakSource("<br>AFTER", "padding:0 7px 0 11px;background:lime;box-decoration-break:"
            + decorationBreak, floated);
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, ForcedBreakOptions());
        HtmlRenderText[] texts = InlineEdgeTexts(rendered);
        HtmlRenderText after = Assert.Single(texts, text => text.Text == "AFTER");

        Assert.Equal(new[] { "AFTER", "BEFORE", "FOLLOW" }, texts.Select(text => text.Text).OrderBy(text => text));
        Assert.Equal(expectedAfterX, after.X, 6);
        Assert.Equal(120D, TableGeometryShape(rendered, "div#case").Height, 6);
        Assert.Equal(127D, after.Y + after.Font.Size, 6);

        HtmlConversionDocument document = HtmlConversionDocument.Parse(html);
        XElement[] paintedFragments = XDocument.Parse(document.ToSvg(ForcedBreakOptions()))
            .Descendants().Where(node => node.Attribute("fill")?.Value == "#00FF00").ToArray();
        Assert.Equal(2, paintedFragments.Length);
        Assert.True(OfficePngReader.TryDecode(document.ExportImage(OfficeImageExportFormat.Png,
            ForcedBreakOptions()).Bytes, out OfficeRasterImage? image));
        HtmlRenderText before = Assert.Single(texts, text => text.Text == "BEFORE");
        int edgeX = (int)Math.Ceiling(before.X + before.TextAdvanceWidth!.Value);
        Assert.Equal(OfficeColor.Lime, image!.GetPixel(edgeX + 5, 40));
        Assert.Equal(decorationBreak == "clone" ? OfficeColor.Lime : OfficeColor.FromRgb(238, 238, 238),
            image.GetPixel(edgeX + 15, 40));
        byte[] pdf = (document.ToPdf(new HtmlToPdfOptions(ForcedBreakOptions()))).ToBytes();
        Assert.Equal(2, OfficeIMO.Pdf.PdfDocument.Load(new System.IO.MemoryStream(pdf)).Render.Drawing(1).Shapes.Count(shape =>
            shape.Shape.FillColor == OfficeColor.Lime));
    }

    private static HtmlRenderOptions ForcedBreakOptions() {
        HtmlRenderOptions options = OrdinaryBaselineOptions();
        options.ViewportHeight = 400D;
        return options;
    }

    private static string ForcedBreakSource(string content, string whitespace, bool floated) =>
        "<!doctype html><style>body{margin:0;padding:24px;font:16px/24px Pinned}</style>"
        + "<div id='case' style='display:flow-root;width:300px;background:#eee'>"
        + (floated ? "<i style='float:left;width:20px;height:100px'></i>" : "")
        + "BEFORE<span style='font-size:40px;line-height:60px;" + whitespace + "'>"
        + content + "</span></div><div>FOLLOW</div>";
}
