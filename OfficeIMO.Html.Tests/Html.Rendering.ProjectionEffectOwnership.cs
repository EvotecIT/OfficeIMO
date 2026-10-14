using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData(0)]
    [InlineData(25)]
    public void SnapshotTextOwnershipRequiresSharedPageAndAncestorClip(int shift) {
        string html = "<style>html,body{height:300px;margin:0}"
            + ".clip{position:absolute;top:150px;height:20px;width:300px;overflow:hidden;"
            + "transform:translateY(" + shift + "px);transform-origin:0 0}"
            + ".clip span{position:relative;top:-150px;font:200px/200px Courier}</style>"
            + "<div class='clip'><span>A</span></div>";
        var options = new HtmlToPdfOptions {
            PageSize = new OfficePageSize(300D / 96D, 100D / 96D),
            Margins = HtmlRenderMargins.All(0), ViewportWidth = 300, ViewportHeight = 300,
            HonorCssPageRules = false, AllowSystemFontFallback = false
        };
        byte[] bytes = HtmlConversionDocument.Parse(html).RenderToPdfBytes(
            HtmlRenderRequest.Create(HtmlRenderIntentProfile.ScreenSnapshotPaged, HtmlRenderEncoder.Pdf, options));
        var pdf = PdfReadDocument.Open(bytes);
        Assert.Equal(string.Empty, pdf.Pages[0].ExtractText().Trim());
        Assert.Equal("A", pdf.Pages[1].ExtractText().Trim());
        Assert.Equal("A", pdf.ExtractText().Trim());
    }

    [Theory]
    [InlineData("<div style='transform:translateY(150px);transform-origin:0 0'>MovedMarker</div>", 1)]
    [InlineData("<div style='height:150px'></div><div style='transform:translateY(-140px);transform-origin:0 0'>MovedMarker</div>", 0)]
    [InlineData("<div style='transform:translateY(50px);transform-origin:0 0'><div style='transform:translateY(100px);transform-origin:0 0'>MovedMarker</div></div>", 1)]
    public void SnapshotProjectionPreservesTranslatedText(string content, int expectedPage) {
        string html = "<style>html,body{margin:0;height:300px}div{font:16px/20px Courier}</style>" + content;
        var options = new HtmlToPdfOptions {PageSize=new OfficePageSize(300d/96d,100d/96d), Margins=HtmlRenderMargins.All(0),
            ViewportWidth=300, ViewportHeight=300, HonorCssPageRules=false, AllowSystemFontFallback=false};
        byte[] pdf = HtmlConversionDocument.Parse(html).RenderToPdfBytes(
            HtmlRenderRequest.Create(HtmlRenderIntentProfile.ScreenSnapshotPaged, HtmlRenderEncoder.Pdf, options));
        var read = PdfReadDocument.Open(pdf);
        Assert.Equal(3, read.Pages.Count);
        Assert.Contains("MovedMarker", read.Pages[expectedPage].ExtractText());
        Assert.Equal(1, read.ExtractText().Split(new[]{"MovedMarker"},StringSplitOptions.None).Length-1);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SnapshotProjectionRetainsTranslatedAncestorClips(bool moveClip) {
        string transform = "transform:translateY(150px);transform-origin:0 0;";
        string html = "<style>html,body{margin:0;height:300px}div{font:16px/20px Courier}</style>"+
            "<div style='height:20px;overflow:hidden;"+(moveClip?transform:"")+"'><div style='"+
            (moveClip?"":transform)+"'>ClippedMarker</div></div>";
        var options = new HtmlToPdfOptions { PageSize=new OfficePageSize(300d/96d,100d/96d), Margins=HtmlRenderMargins.All(0),
            ViewportWidth=300, ViewportHeight=300, HonorCssPageRules=false, AllowSystemFontFallback=false };
        byte[] pdf = HtmlConversionDocument.Parse(html).RenderToPdfBytes(
            HtmlRenderRequest.Create(HtmlRenderIntentProfile.ScreenSnapshotPaged, HtmlRenderEncoder.Pdf, options));
        var read = PdfReadDocument.Open(pdf);
        if (moveClip) Assert.Contains("ClippedMarker", read.Pages[1].ExtractText());
        Assert.Equal(moveClip?1:0, read.ExtractText().Split(new[]{"ClippedMarker"},StringSplitOptions.None).Length-1);
    }

    [Theory]
    [InlineData("transform:translateX(300px);transform-origin:0 0", true)]
    [InlineData("clip-path:inset(20px 990px 60px 0)", false)]
    [InlineData("clip:rect(20px,10px,40px,0)", false)]
    [InlineData("overflow:hidden;clip-path:inset(20px 990px 60px 0)", false)]
    public void EqualUriClippedAnchorDoesNotCoverAnotherAnchorsPositionedImage(string paintEffect, bool translated) {
        string png = Convert.ToBase64String(OfficeIMO.Tests.Pdf.PdfPngTestImages.CreateRgbPng(2,2));
        string html = "<style>body{margin:0}a{position:absolute;left:0;top:0;height:20px}"+
            ".wide{top:-20px;height:100px;width:1000px;"+paintEffect+"}.narrow{width:10px}"+
            "img{position:absolute;left:20px;top:0;width:100px;height:20px}</style>"+
            "<a class='wide' href='https://example.test/same'>Translated</a><a class='narrow' href='https://example.test/same'>"+
            "<img src='data:image/png;base64,"+png+"'></a>";
        byte[] bytes = HtmlConversionDocument.Parse(html).ToPdfBytes(new HtmlToPdfOptions {Margins=HtmlRenderMargins.All(0), AutoFitWidePrintContent=false});
        var links = PdfReadDocument.Open(bytes).Pages[0].GetLinkAnnotations();
        Assert.Contains(links, link => link.Uri=="https://example.test/same" && link.X1 >= 14 && link.X1 <= 16 && link.X2 >= 89);
        if (translated) Assert.Contains(links, link => link.Uri=="https://example.test/same" && link.X1 >= 200);
        else Assert.Contains(links, link => link.Uri=="https://example.test/same" && link.X1 == 0 && link.X2 <= 7.51);
    }
}
