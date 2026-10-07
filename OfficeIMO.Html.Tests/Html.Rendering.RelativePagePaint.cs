using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData(-80, 0, false, false)]
    [InlineData(80, 0, false, false)]
    [InlineData(-80, 40, false, false)]
    [InlineData(-80, 0, true, false)]
    [InlineData(-80, 0, false, true)]
    public void RelativePagePaint_PreservesLabelsAcrossPhysicalContentWindows(
        int offset, int margin, bool varyingHeight, bool earlyBreak) {
        string pageRules = varyingHeight
            ? "@page{size:400px 600px;margin:0}@page:first{size:400px 400px}"
            : "@page{size:400px 400px;margin:" + margin + "px}";
        string rules = pageRules + "html,body{margin:0;font:16px Arial;line-height:20px}p{margin:0}"
            + (earlyBreak ? "#target p:nth-child(6){break-before:page}" : "");
        string rows = string.Concat(Enumerable.Range(1, 18).Select(index => "<p>Plain row " + index.ToString("D2") + "</p>"));
        string content = "<div style='height:240px'>Leading label</div>"
            + "<section id='target' style='position:relative;top:" + offset + "px'>" + rows + "</section>"
            + "<div style='height:40px'>Trailing label</div>";
        HtmlRenderDocument baseline = HtmlRenderTestDriver.Render("<style>" + rules + "</style>" + content.Replace("top:" + offset + "px", "top:0"),
            new HtmlRenderOptions { Mode = HtmlRenderMode.Paged });
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render("<style>" + rules + "</style>" + content,
            new HtmlRenderOptions { Mode = HtmlRenderMode.Paged });

        for (int index = 1; index <= 18; index++) {
            string marker = "Plain row " + index.ToString("D2");
            var occurrences = RelativeTextPlacements(rendered, marker).ToArray();
            var actual = Assert.Single(occurrences);
            Assert.InRange(actual.Text.Y, actual.Page.Margins.Top, actual.Page.Height - actual.Page.Margins.Bottom);
            // This is the independent print fixture's decisive transfer, rather
            // than a same-page offset assertion that permits off-page loss.
            if (index == 9 && offset < 0 && margin == 0 && !earlyBreak) Assert.Equal(1, actual.Page.PageNumber);
            if (earlyBreak && index == 6) Assert.Equal(1, actual.Page.PageNumber);
        }
        var baselineTail = Assert.Single(RelativeTextPlacements(baseline, "Trailing label"));
        var movedTail = Assert.Single(RelativeTextPlacements(rendered, "Trailing label"));
        Assert.Equal(baselineTail.Page.PageNumber, movedTail.Page.PageNumber);
        Assert.Equal(baselineTail.Text.Y, movedTail.Text.Y, 3);
        if (varyingHeight) {
            Assert.Equal(400D, rendered.Pages[0].Height);
            Assert.Equal(600D, rendered.Pages[1].Height);
        }
    }

    [Fact]
    public void RelativePagePaint_EnforcesProjectionBudget() {
        string rows = string.Concat(Enumerable.Range(1, 18).Select(index => "<p style='margin:0'>BudgetRow" + index + "</p>"));
        string html = "<style>@page{size:400px 400px;margin:0}body{margin:0;line-height:20px}</style>"
            + "<div style='height:240px'></div><section style='position:relative;top:-80px'>" + rows + "</section>";
        InvalidOperationException error = Assert.Throws<InvalidOperationException>(() => HtmlRenderTestDriver.Render(html,
            new HtmlRenderOptions { Mode = HtmlRenderMode.Paged, MaxProjectedVisuals = 1 }));
        Assert.Contains("MaxProjectedVisuals", error.Message);
    }

    [Theory]
    [InlineData(-80)]
    [InlineData(80)]
    public void RelativePagePaint_MovesOversizedImageSegmentClipsAndLinks(int offset) {
        // Independent PNG used by the archived browser-print control: three
        // 300px bands, rather than a renderer-produced image round trip.
        byte[] image = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Images", "relative-three-band.png"));
        string html = "<style>@page{size:400px 400px;margin:0}body{margin:0;font:16px Arial;line-height:20px}"
            + "img{display:block;width:200px;height:900px}a{display:block}</style>"
            + "<div style='height:240px'>Leading label</div><section style='position:relative;top:" + offset + "px'>"
            + "<a href='https://example.test/relative-image'><img src='data:image/png;base64," + Convert.ToBase64String(image)
            + "'></a><p style='margin:0'>Image caption</p></section><p style='margin:0'>Trailing label</p>";
        var input = HtmlConversionDocument.Parse(html);
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(input, new HtmlRenderOptions { Mode = HtmlRenderMode.Paged });
        Assert.Equal(4, rendered.Pages.Count);
        if (offset < 0) {
            OfficeRasterImage first = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());
            Assert.Equal(OfficeColor.FromRgb(255, 0, 0), first.GetPixel(100, 360));
        }
        OfficeRasterImage last = OfficeDrawingRasterRenderer.Render(rendered.Pages[3].CreateDrawing());
        Assert.Equal(OfficeColor.FromRgb(0, 0, 255), last.GetPixel(100, offset < 0 ? 10 : 150));

        byte[] pdf = input.ToPdfBytes(new HtmlToPdfOptions {
            ResourcePolicy = PdfCore.PdfResourcePolicy.CreatePortableDeterministic()
        });
        var links = PdfCore.PdfInspector.Inspect(pdf).LinkAnnotations.Where(link => link.Uri == "https://example.test/relative-image").ToArray();
        Assert.Equal(offset < 0 ? new[] { 1, 2, 3, 4 } : new[] { 2, 3, 4 },
            links.Select(link => link.PageNumber!.Value).Distinct().OrderBy(page => page));
        foreach (var pageLinks in links.GroupBy(link => link.PageNumber)) {
            var parts = pageLinks.OrderBy(link => link.Y1).ToArray();
            for (int index = 1; index < parts.Length; index++) {
                Assert.Equal(parts[index - 1].Y2, parts[index].Y1, 3);
            }
        }
        Assert.All(links, link => {
            Assert.InRange(link.Y1, 0D, 300D);
            Assert.InRange(link.Y2, 0D, 300D);
            Assert.True(link.Y2 > link.Y1);
        });
    }

    [Fact]
    public void RelativePagePaint_TransfersInlineTextDecorationAndAnchorTogether() {
        const string uri = "https://example.test/relative-inline";
        string html = "<style>@page{size:400px 400px;margin:0}body{margin:0;font:16px Arial;line-height:20px}p{margin:0}</style>"
            + "<div style='height:400px'></div><p><a href='" + uri + "' style='position:relative;top:-80px;"
            + "background:#ffff00;text-shadow:2px 2px red'>InlineMarker</a><span>TailMarker</span></p>";
        var input = HtmlConversionDocument.Parse(html);
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(input, new HtmlRenderOptions { Mode = HtmlRenderMode.Paged });
        var marker = Assert.Single(RelativeTextPlacements(rendered, "InlineMarker"), item => !item.Text.IsPaintOnly);
        Assert.Equal(1, marker.Page.PageNumber);
        Assert.InRange(marker.Text.Y, 319D, 321D);
        Assert.Equal(2, Assert.Single(RelativeTextPlacements(rendered, "TailMarker")).Page.PageNumber);
        OfficeRasterImage first = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());
        Assert.Equal(OfficeColor.FromRgb(255, 255, 0), first.GetPixel(1, 321));
        var links = PdfCore.PdfInspector.Inspect(input.ToPdfBytes(new HtmlToPdfOptions {
            ResourcePolicy = PdfCore.PdfResourcePolicy.CreatePortableDeterministic()
        })).LinkAnnotations.Where(link => link.Uri == uri).ToArray();
        Assert.Single(links);
        Assert.Equal(1, links[0].PageNumber);
        Assert.InRange(links[0].Y1, 40D, 65D);
    }

    [Fact]
    public void RelativePagePaint_PreservesAuthoredClippingAndDoesNotPaginatePositionedOverlays() {
        string html = "<style>@page{size:400px 400px;margin:0}body{margin:0;font:16px Arial;line-height:20px}p{margin:0}</style>"
            + "<div style='height:80px;overflow:hidden'><div style='position:relative;top:60px;height:60px;background:red'></div></div>"
            + "<div style='position:relative;top:10px;height:20px'>FlowMarker"
            + "<div style='position:absolute;top:700px'>OverlayMarker</div></div>"
            + "<div style='position:fixed;top:700px'><span style='position:relative;top:800px'>FixedMarker</span></div>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions { Mode = HtmlRenderMode.Paged });
        Assert.Single(rendered.Pages);
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());
        Assert.Equal(OfficeColor.FromRgb(255, 0, 0), raster.GetPixel(300, 70));
        Assert.Equal(OfficeColor.White, raster.GetPixel(300, 85));
    }

    [Theory]
    [InlineData("translateY(20px)", 140)]
    [InlineData("translateY(-20px)", 100)]
    [InlineData("scale(1.5)", 220)]
    public void RelativePagePaint_PreservesEffectCoordinatesAndPhysicalWindowClipping(string transform, int finalBottom) {
        string html = "<style>@page{size:400px 400px;margin:0}body{margin:0}</style><div style='height:400px'></div>"
            + "<div style='position:relative;top:-80px;transform:" + transform
            + ";transform-origin:0 0;width:200px;height:200px;background:red'></div>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions { Mode = HtmlRenderMode.Paged });
        Assert.Equal(2, rendered.Pages.Count);
        OfficeRasterImage first = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());
        OfficeRasterImage last = OfficeDrawingRasterRenderer.Render(rendered.Pages[1].CreateDrawing());
        Assert.Equal(OfficeColor.FromRgb(255, 0, 0), first.GetPixel(100, 350));
        Assert.Equal(OfficeColor.FromRgb(255, 0, 0), last.GetPixel(100, finalBottom - 10));
        Assert.Equal(OfficeColor.White, last.GetPixel(100, finalBottom + 10));
    }

    [Fact]
    public void RelativePagePaint_RasterKeepsTextMadeVisibleByItsEnclosingEffect() {
        string html = "<style>@page{size:400px 400px;margin:0}body{margin:0;font:16px Arial}</style>"
            + "<div style='position:relative;top:-80px;transform:translateY(80px);transform-origin:0 0'>VisibleEffectMarker</div>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions { Mode = HtmlRenderMode.Paged });
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(Assert.Single(rendered.Pages).CreateDrawing());
        Assert.Contains(Enumerable.Range(0, 24), y => Enumerable.Range(0, 200)
            .Any(x => raster.GetPixel(x, y) != OfficeColor.White));
    }

    [Fact]
    public void RelativePagePaint_SingularEffectDoesNotCreateOverflowPagesOrRejectTheDocument() {
        string html = "<style>@page{size:400px 400px;margin:0}body{margin:0}</style><div style='height:400px'></div>"
            + "<div style='position:relative;top:-80px;transform:matrix(1,1,1,1,0,0);transform-origin:0 0;"
            + "width:200px;height:200px;background:red'></div>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions { Mode = HtmlRenderMode.Paged });
        Assert.Equal(2, rendered.Pages.Count);
        foreach (HtmlRenderPage page in rendered.Pages) {
            OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(page.CreateDrawing());
            Assert.Equal(OfficeColor.White, raster.GetPixel(100, 100));
        }
    }

    [Fact]
    public void RelativePagePaint_ProjectsNamedFontAscentAndItsShadowAboveTheFrame() {
        byte[] font = OfficeIMO.TestAssets.ManagedTextShapingTestAssets.CreateFontWithVerticalMetrics('A', 1069, -200, 1040);
        var options = new HtmlToPdfOptions { Margins = HtmlRenderMargins.All(0D) };
        options.ResourcePolicy.AllowDocumentFontEmbedding = true;
        options.PdfOptions.RegisterNamedFontFamily(new PdfCore.PdfEmbeddedFontFamily("Arial", font));
        string html = "<style>@page{size:400px 400px;margin:0}body{margin:0}</style><div style='height:400px'></div>"
            + "<div style='position:relative;top:50px;font:20px/20px Arial'>"
            + "<span style='font-size:100px;text-shadow:0 0 red'>A</span></div>";
        HtmlPdfRenderResult result = HtmlPdfRenderedConverter.Convert(HtmlConversionDocument.Parse(html), options);
        HtmlRenderDocument rendered = result.RenderResult!.Document;
        Assert.Equal(2, rendered.Pages.Count);
        Assert.Contains(RelativeTextPlacements(rendered, "A"), placement => placement.Page.PageNumber == 1);
        Assert.Contains(RelativeTextPlacements(rendered, "A"), placement => placement.Page.PageNumber == 2);
        Assert.Equal("A", PdfCore.PdfReadDocument.Open(result.Document.ToBytes()).ExtractText().Trim());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void RelativePagePaint_ProjectsNamedFontAndShadowBelowThePage(bool precedingNormalText) {
        byte[] font = OfficeIMO.TestAssets.ManagedTextShapingTestAssets.CreateFontWithVerticalMetrics('A', 800, -100, 700);
        var options = new HtmlToPdfOptions { Margins = HtmlRenderMargins.All(0D) };
        options.ResourcePolicy.AllowDocumentFontEmbedding = true;
        options.PdfOptions.RegisterNamedFontFamily(new PdfCore.PdfEmbeddedFontFamily("Arial", font));
        string html = "<style>@page{size:400px 400px;margin:0}body{margin:0}</style>"
            + "<div style='position:relative;top:365px;font:20px/20px Arial'>"
            + (precedingNormalText ? "A" : string.Empty)
            + "<span style='font-size:100px;text-shadow:0 0 red'>A</span></div>";

        HtmlPdfRenderResult result = HtmlPdfRenderedConverter.Convert(HtmlConversionDocument.Parse(html), options);

        Assert.Equal(2, result.RenderResult!.Document.Pages.Count);
        Assert.Contains(RelativeTextPlacements(result.RenderResult.Document, "A"),
            placement => placement.Page.PageNumber == 2 && placement.Text.Font.Size == 100D);
        string extracted = PdfCore.PdfReadDocument.Open(result.Document.ToBytes()).ExtractText();
        Assert.Equal(precedingNormalText ? "AA" : "A", string.Concat(extracted.Where(c => !char.IsWhiteSpace(c))));
    }

    [Fact]
    public void RelativePagePaint_KeepsOverflowGlyphPaintTranslatedIntoThePage() {
        string html = "<style>@page{size:400px 400px;margin:0}body{margin:0;font:16px Arial}</style>"
            + "<div style='position:relative;top:10px;width:200px;white-space:nowrap;transform:translateX(-200px)'>"
            + string.Concat(Enumerable.Repeat("ABCDEFGHIJKLMNO ", 8)) + "</div>";
        var input = HtmlConversionDocument.Parse(html);
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(input, new HtmlRenderOptions { Mode = HtmlRenderMode.Paged });
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(Assert.Single(rendered.Pages).CreateDrawing());
        Assert.Contains(Enumerable.Range(300, 100), x => Enumerable.Range(10, 24)
            .Any(y => raster.GetPixel(x, y) != OfficeColor.White));

    }

    private static IEnumerable<(HtmlRenderPage Page, HtmlRenderText Text)> RelativeTextPlacements(HtmlRenderDocument document, string marker) {
        foreach (HtmlRenderPage page in document.Pages) {
            foreach (HtmlRenderVisual visual in Leaves(page.Scene)) {
                if (visual is HtmlRenderText text && text.Text == marker) yield return (page, text);
            }
        }

        static IEnumerable<HtmlRenderVisual> Leaves(IEnumerable<HtmlRenderVisual> visuals) {
            foreach (HtmlRenderVisual visual in visuals) {
                if (visual.PaintChildren is { } children) {
                    foreach (HtmlRenderVisual child in Leaves(children)) yield return child;
                } else yield return visual;
            }
        }
    }
}
