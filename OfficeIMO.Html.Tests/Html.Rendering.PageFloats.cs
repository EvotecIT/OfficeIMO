using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("top", 10D, 50D)]
    [InlineData("bottom", 130D, 10D)]
    [InlineData("snap", 10D, 50D)]
    public void HtmlPageFloats_EdgePlacementReservesBodySpaceWithoutAdvancingTheAnchor(string side, double floatY, double firstY) {
        string html = "<style>@page{size:240px 180px;margin:10px}body,p,aside{margin:0;font:12px/20px Arial}"
            + ".edge-float{float:" + side + ";float-reference:page;height:40px;background:red}</style>"
            + "<p>One</p><aside id='edge-float' class='edge-float'>Floating</aside>"
            + "<p><a href='https://example.test/two'>Two</a></p>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged, HonorCssPageRules = true
        });
        Assert.Single(rendered.Pages);
        HtmlRenderText[] text = rendered.Pages[0].Visuals.OfType<HtmlRenderText>().ToArray();
        Assert.Equal(3, text.Length);
        Assert.Equal(floatY, Assert.Single(text, item => item.Text == "Floating").Y, 3);
        Assert.Equal(firstY, Assert.Single(text, item => item.Text == "One").Y, 3);
        HtmlRenderText last = Assert.Single(text, item => item.Text == "Two");
        Assert.Equal(firstY + 20D, last.Y, 3);
        Assert.Equal("https://example.test/two", last.LinkUri);
        Assert.DoesNotContain(rendered.Diagnostics, item => item.Code == HtmlRenderDiagnosticCodes.FloatValueUnsupported);
    }

    [Fact]
    public void HtmlPageFloats_SnapUsesTheAnchorBeforeItsOwnReservationAndConverges() {
        const string html = "<style>@page{size:240px 180px;margin:10px}body,p,aside{margin:0;font:12px/20px Arial}</style>"
            + "<p>One</p><p>Two</p><p>Three</p>"
            + "<aside style='float:snap;float-reference:page;height:40px'>Floating</aside><p>Four</p>";
        HtmlRenderDocument rendered = RenderPageFloatFixture(html);
        Assert.Single(rendered.Pages);
        HtmlRenderText[] text = rendered.Pages[0].Visuals.OfType<HtmlRenderText>().ToArray();
        Assert.Equal(5, text.Length);
        Assert.Equal(10D, Assert.Single(text, item => item.Text == "Floating").Y, 3);
        Assert.Equal(110D, Assert.Single(text, item => item.Text == "Four").Y, 3);
        Assert.DoesNotContain(rendered.Diagnostics, item => item.Severity == HtmlDiagnosticSeverity.Error);
    }

    [Fact]
    public void HtmlPageFloats_NestedAnchorAndForcedPageBreakPreserveWholeFloatAndBody() {
        const string html = "<style>@page{size:240px 120px;margin:10px}body,p,aside,section{margin:0;font:12px/20px Arial}</style>"
            + "<p style='break-after:page'>Prelude</p><section><p>One"
            + "<span style='float:top;float-reference:page;height:40px;background:red'>Floating</span></p>"
            + "<p>Two</p><p>Three</p><p>Four</p></section>";
        HtmlRenderDocument rendered = RenderPageFloatFixture(html);
        HtmlRenderText[] all = rendered.Pages.SelectMany(page => EnumerateRenderVisuals(page.Scene)).OfType<HtmlRenderText>().ToArray();
        foreach (string marker in new[] { "Prelude", "One", "Floating", "Two", "Three", "Four" }) Assert.Single(all, item => item.Text == marker);
        Assert.DoesNotContain(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), item => item.Text == "Floating");
        Assert.Equal(10D, Assert.Single(rendered.Pages[1].Visuals.OfType<HtmlRenderText>(), item => item.Text == "Floating").Y, 3);
        Assert.DoesNotContain(rendered.Diagnostics, item => item.Code == HtmlRenderDiagnosticCodes.ForcedFragment || item.Severity == HtmlDiagnosticSeverity.Error);
        Assert.All(rendered.Pages, page => Assert.All(page.Visuals.OfType<HtmlRenderText>(), item => {
            Assert.True(item.Y >= 10D - 0.0001D);
            Assert.True(item.Y + item.Height <= 110D + 0.0001D);
        }));
    }

    [Fact]
    public void HtmlPageFloats_MultipleEdgesAndFootnoteHaveDisjointReservedAreas() {
        const string html = "<style>@page{size:240px 180px;margin:10px}body,p,aside,span{margin:0;font:12px/20px Arial}</style>"
            + "<aside style='float:top;float-reference:page;height:20px'>Top</aside>"
            + "<p>Body<span style='float:footnote;font:10px/12px Arial'>Note</span></p>"
            + "<aside style='float:bottom;float-reference:page;height:20px'>Bottom</aside><p>End</p>";
        HtmlRenderDocument rendered = RenderPageFloatFixture(html);
        Assert.Single(rendered.Pages);
        HtmlRenderText[] text = rendered.Pages[0].Visuals.OfType<HtmlRenderText>().ToArray();
        HtmlRenderText top = Assert.Single(text, item => item.Text == "Top");
        HtmlRenderText body = Assert.Single(text, item => item.Text == "Body");
        HtmlRenderText bottom = Assert.Single(text, item => item.Text == "Bottom");
        HtmlRenderText note = Assert.Single(text, item => item.Text == "Note");
        Assert.Equal(10D, top.Y, 3);
        Assert.True(body.Y >= top.Y + top.Height);
        Assert.True(bottom.Y >= body.Y + body.Height);
        Assert.True(note.Y >= bottom.Y + bottom.Height);
        Assert.True(note.Y + note.Height <= 170D);
        Assert.DoesNotContain(rendered.Diagnostics, item => item.Severity == HtmlDiagnosticSeverity.Error);
    }

    [Fact]
    public void HtmlPageFloats_FullPageOrOversizedFloatUsesObservableContentPreservingFallback() {
        const string html = "<style>@page{size:240px 120px;margin:10px}body,p,aside{margin:0;font:12px/20px Arial}</style>"
            + "<aside style='float:top;float-reference:page;height:100px'>Floating</aside><p>After</p>";
        HtmlRenderDocument rendered = RenderPageFloatFixture(html);
        HtmlRenderText[] text = rendered.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderText>().ToArray();
        Assert.Single(text, item => item.Text == "Floating");
        Assert.Single(text, item => item.Text == "After");
        Assert.Contains(rendered.Diagnostics, item => item.Code == HtmlRenderDiagnosticCodes.FloatValueUnsupported);
    }

    private static HtmlRenderDocument RenderPageFloatFixture(string html) => HtmlRenderTestDriver.Render(html,
        new HtmlRenderOptions { Mode = HtmlRenderMode.Paged, HonorCssPageRules = true });

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void HtmlPageFloats_TerminalAnchorSurvivesAFragmentedSection(bool pathClipped) {
        string html = "<style>@page{size:240px 120px;margin:10px}body,p,aside,section{margin:0;font:12px/20px Arial}"
            + (pathClipped ? "section{clip-path:inset(0)}" : string.Empty) + "</style>"
            + "<section>" + string.Concat(Enumerable.Range(1, 6).Select(i => "<p>Line" + i + "</p>"))
            + "<aside style='float:top;float-reference:page;height:20px'>Floating</aside></section>";
        HtmlRenderDocument rendered = RenderPageFloatFixture(html);
        Assert.DoesNotContain(EnumerateRenderVisuals(rendered.Pages[0].Scene).OfType<HtmlRenderText>(), item => item.Text == "Floating");
        Assert.Contains(EnumerateRenderVisuals(rendered.Pages[1].Scene).OfType<HtmlRenderText>(), item => item.Text == "Floating" && item.Y == 10D);
        HtmlRenderText[] all = rendered.Pages.SelectMany(page => EnumerateRenderVisuals(page.Scene)).OfType<HtmlRenderText>().ToArray();
        foreach (string marker in Enumerable.Range(1, 6).Select(i => "Line" + i).Append("Floating")) Assert.Single(all, item => item.Text == marker);
        Assert.DoesNotContain(rendered.Diagnostics, item => item.Severity == HtmlDiagnosticSeverity.Error);
    }

    [Fact]
    public void HtmlPageFloats_DeferredFloatSkipsAnInsufficientLeftPage() {
        const string html = "<style>@page{size:240px 120px;margin:10px}@page:left{margin-top:40px;margin-bottom:40px}"
            + "body,p,aside{margin:0;font:12px/20px Arial}</style>"
            + "<aside style='float:top;float-reference:page;height:60px'>FirstFloat</aside>"
            + "<aside style='float:bottom;float-reference:page;height:60px'>SecondFloat</aside><p>Body</p>";
        HtmlRenderDocument rendered = RenderPageFloatFixture(html);
        Assert.Equal(3, rendered.Pages.Count);
        Assert.Contains(rendered.Pages[2].Visuals.OfType<HtmlRenderText>(), item => item.Text == "SecondFloat" && item.Y == 50D);
        HtmlRenderText[] all = rendered.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderText>().ToArray();
        foreach (string marker in new[] { "FirstFloat", "SecondFloat", "Body" }) Assert.Single(all, item => item.Text == marker);
        Assert.DoesNotContain(rendered.Diagnostics, item => item.Severity == HtmlDiagnosticSeverity.Error);
    }

    [Fact]
    public void HtmlPageFloats_ContainedFootnoteDefersWithoutMovingItsFloatOrOverlapping() {
        const string html = "<style>@page{size:240px 120px;margin:10px}body,p,aside,span{margin:0;font:12px/20px Arial}</style>"
            + "<aside style='float:top;float-reference:page;height:80px'>Floating"
            + "<span style='float:footnote'>Note</span></aside><p>Body</p>";
        HtmlRenderDocument rendered = RenderPageFloatFixture(html);
        Assert.Contains(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), item => item.Text == "Floating" && item.Y == 10D);
        Assert.DoesNotContain(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), item => item.Text == "Note");
        Assert.Contains(rendered.Pages[1].Visuals.OfType<HtmlRenderText>(), item => item.Text == "Note");
        HtmlRenderText[] all = rendered.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderText>().ToArray();
        foreach (string marker in new[] { "Floating", "Note", "Body" }) Assert.Single(all, item => item.Text == marker);
        Assert.DoesNotContain(rendered.Diagnostics, item => item.Severity == HtmlDiagnosticSeverity.Error);
        Assert.All(rendered.Pages, page => Assert.All(page.Visuals.OfType<HtmlRenderText>(), item => Assert.True(item.Y + item.Height <= 110D + 0.0001D)));
    }

    [Theory]
    [InlineData("top")]
    [InlineData("bottom")]
    [InlineData("snap")]
    public void HtmlPageFloats_InlineSourceRetainsLogicalOrderWhenPaintedAtThePageEdge(string side) {
        string html = "<style>@page{size:240px 180px;margin:10px}body,p,span{margin:0;font:12px/20px Arial}</style>"
            + "<p>Before <span style='float:" + side + ";float-reference:page;height:20px'>Floating </span>After</p>";
        HtmlRenderDocument rendered = RenderPageFloatFixture(html);
        Assert.Equal("Before Floating After", string.Join(" ", rendered.Text.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries)));
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdfBytes(new OfficeIMO.Html.Pdf.HtmlToPdfOptions { AutoFitWidePrintContent = false });
        string text = OfficeIMO.Pdf.PdfReadDocument.Open(pdf).ExtractText();
        Assert.Equal("Before Floating After", string.Join(" ", text.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries)));
    }

    [Fact]
    public void HtmlPageFloats_InterleavedContainerRetainsParagraphSeparatorsArtifactsAndLinks() {
        const string html = "<style>@page{size:300px 220px;margin:10px}body,p,span,section{margin:0;font:12px/20px Arial}</style>"
            + "<section><p>First</p><p><a href='https://example.test/before'>Before </a>"
            + "<span style='float:top;float-reference:page;height:20px'><a href='https://example.test/float'>Floating </a></span>"
            + "<span style='-officeimo-pdf-tag-type:artifact'>Hidden </span>After</p><p>Last</p></section>";
        HtmlRenderDocument rendered = RenderPageFloatFixture(html);
        Assert.Equal("First Before Floating After Last", string.Join(" ", rendered.Text.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries)));
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdfBytes(new HtmlToPdfOptions { AutoFitWidePrintContent = false });
        string text = OfficeIMO.Pdf.PdfReadDocument.Open(pdf).ExtractText();
        Assert.Equal("First Before Floating After Last", string.Join(" ", text.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries)));
        var info = OfficeIMO.Pdf.PdfInspector.Inspect(pdf);
        Assert.Contains("https://example.test/before", info.LinkUris);
        Assert.Contains("https://example.test/float", info.LinkUris);
    }

    [Fact]
    public void HtmlPageFloats_TwoInlineFloatsShareOneSourceTextOwner() {
        const string html = "<style>@page{size:300px 220px;margin:10px}body,p,span{margin:0;font:12px/20px Arial}</style>"
            + "<p>Before <span style='float:top;float-reference:page;height:20px'>Top </span>Between "
            + "<span style='float:bottom;float-reference:page;height:20px'>Bottom </span>After</p>";
        HtmlRenderDocument rendered = RenderPageFloatFixture(html);
        Assert.Equal("Before Top Between Bottom After", string.Join(" ", rendered.Text.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries)));
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdfBytes(new HtmlToPdfOptions { AutoFitWidePrintContent = false });
        string text = OfficeIMO.Pdf.PdfReadDocument.Open(pdf).ExtractText();
        Assert.Equal("Before Top Between Bottom After", string.Join(" ", text.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries)));
    }

    [Fact]
    public void HtmlPageFloats_MultipleWholeFloatsDeferWithoutLosingTheirAnchorsOrText() {
        const string html = "<style>@page{size:240px 120px;margin:10px}body,p,aside{margin:0;font:12px/20px Arial}</style>"
            + "<aside style='float:top;float-reference:page;height:60px'>FirstFloat</aside>"
            + "<aside style='float:bottom;float-reference:page;height:60px'>SecondFloat</aside>"
            + "<p>Body</p><p>End</p>";
        HtmlRenderDocument rendered = RenderPageFloatFixture(html);
        Assert.Equal(2, rendered.Pages.Count);
        Assert.Contains(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), item => item.Text == "FirstFloat" && item.Y == 10D);
        Assert.Contains(rendered.Pages[1].Visuals.OfType<HtmlRenderText>(), item => item.Text == "SecondFloat" && item.Y == 50D);
        HtmlRenderText[] text = rendered.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderText>().ToArray();
        foreach (string marker in new[] { "FirstFloat", "SecondFloat", "Body", "End" }) Assert.Single(text, item => item.Text == marker);
        Assert.DoesNotContain(rendered.Diagnostics, item => item.Code == HtmlRenderDiagnosticCodes.ForcedFragment || item.Severity == HtmlDiagnosticSeverity.Error);
    }

    [Fact]
    public void HtmlPageFloats_StylesheetCascadeRetainsImportantReferenceAndSideInTheirLayer() {
        const string html = "<style>@page{size:240px 180px;margin:10px}body,p,aside{margin:0;font:12px/20px Arial}"
            + "@layer first,second;@layer first{aside{float:top!important;float-reference:page!important;height:40px}}"
            + "@layer second{aside{float:bottom!important;float-reference:inline!important}}</style>"
            + "<p>One</p><aside>Floating</aside><p>Two</p>";
        HtmlRenderDocument rendered = RenderPageFloatFixture(html);
        HtmlRenderText[] text = Assert.Single(rendered.Pages).Visuals.OfType<HtmlRenderText>().ToArray();
        Assert.Equal(10D, Assert.Single(text, item => item.Text == "Floating").Y, 3);
        Assert.Equal(50D, Assert.Single(text, item => item.Text == "One").Y, 3);
        Assert.DoesNotContain(rendered.Diagnostics, item => item.Code == HtmlRenderDiagnosticCodes.FloatValueUnsupported);
    }

    [Fact]
    public void HtmlPageFloats_NestedPageFloatKeepsItsBodyWithAnExplicitFallback() {
        const string html = "<style>@page{size:240px 180px;margin:10px}body,p,aside{margin:0;font:12px/20px Arial}</style>"
            + "<aside style='float:top;float-reference:page;height:60px'>Outer"
            + "<aside style='float:bottom;float-reference:page;height:20px'>Inner</aside></aside><p>End</p>";
        HtmlRenderDocument rendered = RenderPageFloatFixture(html);
        HtmlRenderText[] text = rendered.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderText>().ToArray();
        foreach (string marker in new[] { "Outer", "Inner", "End" }) Assert.Single(text, item => item.Text == marker);
        Assert.Contains(rendered.Diagnostics, item => item.Code == HtmlRenderDiagnosticCodes.FloatValueUnsupported);
        Assert.DoesNotContain(rendered.Diagnostics, item => item.Severity == HtmlDiagnosticSeverity.Error);
    }
}
