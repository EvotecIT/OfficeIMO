using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("top", 10D, 50D)]
    [InlineData("bottom", 70D, 10D)]
    [InlineData("block-start", 10D, 50D)]
    [InlineData("block-end", 70D, 10D)]
    public void HtmlColumnEdgeFloats_ReserveOnlyTheirOriginatingColumn(string side, double floatY, double bodyY) {
        string html = ColumnEdgeStyle() + "<section class='columns'><p>Before</p>"
            + "<aside id='figure' style='float:" + side + ";float-reference:column;height:40px'>Figure</aside>"
            + string.Concat(Enumerable.Range(0, 7).Select(i => "<p>Body" + i + "</p>"))
            + "</section><p><a href='#figure'>After</a></p>";
        HtmlRenderDocument document = HtmlRenderTestDriver.Render(html,
            new HtmlRenderOptions { Mode = HtmlRenderMode.Paged, HonorCssPageRules = true });
        HtmlRenderText[] text = document.Pages.SelectMany(p => p.Visuals).OfType<HtmlRenderText>().ToArray();
        HtmlRenderText figure = Assert.Single(text, t => t.Text == "Figure");
        Assert.Equal(10D, figure.X, 3);
        Assert.Equal(floatY, figure.Y, 3);
        Assert.Equal(bodyY, Assert.Single(text, t => t.Text == "Before").Y, 3);
        HtmlRenderText neighbor = text.First(t => t.Text.StartsWith("Body") && t.X > 110D);
        Assert.Equal(10D, neighbor.Y, 3);
        for (int i = 0; i < 7; i++) Assert.Single(text, t => t.Text == "Body" + i);
        Assert.Single(text, t => t.Text == "After");
        Assert.Single(document.Pages.SelectMany(p => p.Visuals).OfType<HtmlRenderNamedDestination>(),
            d => d.Name == "figure");
        Assert.DoesNotContain(document.Diagnostics, d => d.Code == HtmlRenderDiagnosticCodes.FloatValueUnsupported);
    }

    [Fact]
    public void HtmlColumnEdgeFloats_MultipleEdgesAndColumnNoteReserveDisjointAreas() {
        string html = ColumnEdgeStyle() + "<section class='columns'>"
            + "<aside style='float:top;float-reference:column;height:20px'>Top</aside>"
            + "<p>Call<span style='float:footnote;float-reference:column;font:10px/12px Arial'>Note</span></p>"
            + "<aside style='float:bottom;float-reference:column;height:20px'>Bottom</aside><p>Body</p></section><p>End</p>";
        HtmlRenderDocument document = RenderColumnEdgeFixture(html);
        HtmlRenderText[] text = document.Pages.SelectMany(p => p.Visuals).OfType<HtmlRenderText>().ToArray();
        HtmlRenderText top = Assert.Single(text, t => t.Text == "Top");
        HtmlRenderText call = Assert.Single(text, t => t.Text == "Call");
        HtmlRenderText bottom = Assert.Single(text, t => t.Text == "Bottom");
        HtmlRenderText note = Assert.Single(text, t => t.Text == "Note");
        Assert.Equal(10D, top.Y, 3);
        Assert.Equal(30D, call.Y, 3);
        Assert.True(bottom.Y >= call.Y + call.Height);
        Assert.True(note.Y >= bottom.Y + 20D);
        Assert.True(note.Y + note.Height <= 110D + 0.001D);
        foreach (string label in new[] { "Body", "End" }) Assert.Single(text, t => t.Text == label);
        Assert.DoesNotContain(document.Diagnostics, d => d.Severity == HtmlDiagnosticSeverity.Error);
    }

    [Fact]
    public void HtmlColumnEdgeFloats_DeferredWholeFiguresPreserveBodyAndColumnLimit() {
        string html = ColumnEdgeStyle() + "<section class='columns'><p>Before</p>"
            + "<aside style='float:top;float-reference:column;height:60px'>First</aside>"
            + "<aside style='float:bottom;float-reference:column;height:60px'>Second</aside>"
            + string.Concat(Enumerable.Range(0, 20).Select(i => "<p>Body" + i + "</p>")) + "</section><p>End</p>";
        HtmlRenderDocument document = RenderColumnEdgeFixture(html);
        HtmlRenderText[] text = document.Pages.SelectMany(p => p.Visuals).OfType<HtmlRenderText>().ToArray();
        HtmlRenderText first = Assert.Single(text, t => t.Text == "First");
        HtmlRenderText second = Assert.Single(text, t => t.Text == "Second");
        Assert.Equal(10D, first.X, 3);
        Assert.Equal(130D, second.X, 3);
        Assert.Equal(50D, second.Y, 3);
        for (int i = 0; i < 20; i++) Assert.Single(text, t => t.Text == "Body" + i);
        Assert.Single(text, t => t.Text == "End");
        Assert.All(document.Pages, page => Assert.All(page.Visuals.OfType<HtmlRenderText>(), t => {
            Assert.True(t.X >= 10D && t.X + t.Width <= 230D + 0.001D);
            Assert.True(t.Y >= 10D && t.Y + t.Height <= 230D + 0.001D);
        }));
        HtmlDomLimitException error = Assert.Throws<HtmlDomLimitException>(() => HtmlRenderTestDriver.Render(html,
            new HtmlRenderOptions { Mode = HtmlRenderMode.Paged, HonorCssPageRules = true, MaxColumnCount = 2 }));
        Assert.Equal(HtmlRenderDiagnosticCodes.MultiColumnLimitExceeded, error.Code);
    }

    [Fact]
    public void HtmlColumnEdgeFloats_TallInlineAnchorCanPrecedeDeferredFigure() {
        string html = ColumnEdgeStyle() + "<section class='columns'><p style='line-height:60px'>Call"
            + "<span style='float:top;float-reference:column;height:70px'>Figure</span></p><p>End</p></section>";
        HtmlRenderDocument document = RenderColumnEdgeFixture(html);
        HtmlRenderText[] text = document.Pages.SelectMany(p => p.Visuals).OfType<HtmlRenderText>().ToArray();
        Assert.Equal(10D, Assert.Single(text, t => t.Text == "Call").X, 3);
        HtmlRenderText figure = Assert.Single(text, t => t.Text == "Figure");
        Assert.Equal(130D, figure.X, 3);
        Assert.Equal(10D, figure.Y, 3);
        Assert.Single(text, t => t.Text == "End");
        Assert.DoesNotContain(document.Diagnostics, d => d.Severity == HtmlDiagnosticSeverity.Error);
    }

    [Theory]
    [InlineData(100)]
    [InlineData(120)]
    public void HtmlColumnEdgeFloats_OversizedFiguresRetainObservableNormalFlowFallback(int height) {
        string html = ColumnEdgeStyle() + "<section class='columns'><aside style='float:bottom;float-reference:column;height:"
            + height + "px'>Figure</aside><p>After</p></section>";
        HtmlRenderDocument document = RenderColumnEdgeFixture(html);
        HtmlRenderText[] text = document.Pages.SelectMany(p => p.Visuals).OfType<HtmlRenderText>().ToArray();
        Assert.Single(text, t => t.Text == "Figure");
        Assert.Single(text, t => t.Text == "After");
        Assert.Contains(document.Diagnostics, d => d.Code == HtmlRenderDiagnosticCodes.FloatValueUnsupported);
    }

    [Fact]
    public void HtmlColumnEdgeFloats_ContinuousModeRetainsSourceFlow() {
        string html = ColumnEdgeStyle() + "<section class='columns'><p>Before</p>"
            + "<aside style='float:top;float-reference:column;height:40px'>Figure</aside><p>After</p></section>";
        HtmlRenderDocument document = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions { Mode = HtmlRenderMode.Continuous });
        HtmlRenderText[] text = document.Pages.SelectMany(p => p.Visuals).OfType<HtmlRenderText>().ToArray();
        Assert.True(Assert.Single(text, t => t.Text == "Figure").Y > Assert.Single(text, t => t.Text == "Before").Y);
        Assert.Single(text, t => t.Text == "After");
        Assert.Contains(document.Diagnostics, d => d.Code == HtmlRenderDiagnosticCodes.FloatValueUnsupported);
    }

    [Fact]
    public void HtmlColumnEdgeFloats_BalancedColumnsIncludeFigureReservation() {
        string html = ColumnEdgeStyle() + "<section class='columns' style='height:auto;column-fill:balance'>"
            + "<aside style='float:top;float-reference:column;height:40px'>Figure</aside>"
            + string.Concat(Enumerable.Range(0, 8).Select(i => "<p>Body" + i + "</p>")) + "</section><p>After</p>";
        HtmlRenderDocument document = RenderColumnEdgeFixture(html);
        Assert.Single(document.Pages);
        HtmlRenderText[] text = document.Pages[0].Visuals.OfType<HtmlRenderText>().ToArray();
        Assert.Equal(10D, Assert.Single(text, t => t.Text == "Figure").Y, 3);
        for (int n = 0; n < 8; n++) Assert.Single(text, t => t.Text == "Body" + n);
        HtmlRenderText after = Assert.Single(text, t => t.Text == "After");
        Assert.True(after.Y >= text.Where(t => t.Text.StartsWith("Body")).Max(t => t.Y + 20D));
        Assert.DoesNotContain(document.Diagnostics, d => d.Severity == HtmlDiagnosticSeverity.Error);
    }

    [Fact]
    public void HtmlColumnEdgeFloats_NotesInsideFigureRemainVisible() {
        string html = ColumnEdgeStyle() + "<section class='columns'>"
            + "<aside style='float:top;float-reference:column;height:40px'>Figure"
            + "<span style='float:footnote;float-reference:column;font:10px/12px Arial'>FigureNote</span></aside>"
            + "<p>Body</p></section><p>After</p>";
        HtmlRenderDocument document = RenderColumnEdgeFixture(html);
        foreach (string label in new[] { "Figure", "FigureNote", "Body", "After" })
            Assert.Single(document.Pages.SelectMany(p => p.Visuals).OfType<HtmlRenderText>(), t => t.Text == label);
        Assert.DoesNotContain(document.Diagnostics, d => d.Severity == HtmlDiagnosticSeverity.Error);
    }

    [Fact]
    public void HtmlColumnEdgeFloats_InnerColumnContainerOwnsItsFigure() {
        string html = ColumnEdgeStyle() + "<section class='columns' style='height:160px'>"
            + "<section class='columns' style='column-gap:10px;height:80px'>"
            + "<p>Before</p><aside style='float:top;float-reference:column;height:20px'>Inner</aside>"
            + "<p>AfterInner</p></section><p>Outer</p></section><p>End</p>";
        HtmlRenderDocument document = RenderColumnEdgeFixture(html);
        HtmlRenderText[] text = document.Pages.SelectMany(p => p.Visuals).OfType<HtmlRenderText>().ToArray();
        Assert.Equal(10D, Assert.Single(text, t => t.Text == "Inner").Y, 3);
        Assert.Equal(30D, Assert.Single(text, t => t.Text == "Before").Y, 3);
        foreach (string label in new[] { "AfterInner", "Outer", "End" }) Assert.Single(text, t => t.Text == label);
        Assert.DoesNotContain(document.Diagnostics, d => d.Code == HtmlRenderDiagnosticCodes.FloatValueUnsupported);
    }

    [Fact]
    public void HtmlColumnEdgeFloats_BalancingProbesCanExceedTheFinalColumnLimit() {
        string html = ColumnEdgeStyle() + "<section class='columns' style='height:auto;column-fill:balance'>"
            + "<aside style='float:top;float-reference:column;height:40px'>Figure</aside>"
            + "<p style='height:60px;break-inside:avoid'>First</p><p style='height:60px;break-inside:avoid'>Second</p></section>";
        HtmlRenderDocument document = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged, HonorCssPageRules = true, MaxColumnCount = 2
        });
        HtmlRenderText[] text = document.Pages.SelectMany(p => p.Visuals).OfType<HtmlRenderText>().ToArray();
        foreach (string label in new[] { "Figure", "First", "Second" }) Assert.Single(text, t => t.Text == label);
        Assert.Equal(10D, Assert.Single(text, t => t.Text == "Figure").X, 3);
        Assert.Equal(130D, Assert.Single(text, t => t.Text == "Second").X, 3);
    }

    [Theory]
    [InlineData("top;float-reference:page")]
    [InlineData("footnote")]
    public void HtmlColumnEdgeFloats_ExtractedAncestorUsesContentPreservingNestedFallback(string outerFloat) {
        string html = ColumnEdgeStyle() + "<section class='columns'><p>Body<span style='float:" + outerFloat + "'>"
            + "Outer<span style='float:top;float-reference:column;height:20px'>Nested</span>Tail</span></p>"
            + "<p>After</p></section>";
        HtmlRenderDocument document = RenderColumnEdgeFixture(html);
        string text = string.Join(" ", document.Pages.SelectMany(p => p.Visuals).OfType<HtmlRenderText>().Select(t => t.Text));
        foreach (string label in new[] { "Body", "Outer", "Nested", "Tail", "After" }) Assert.Contains(label, text);
        Assert.Contains(document.Diagnostics, d => d.Code == HtmlRenderDiagnosticCodes.FloatValueUnsupported);
        Assert.DoesNotContain(document.Diagnostics, d => d.Severity == HtmlDiagnosticSeverity.Error);
    }

    [Theory]
    [InlineData("column")]
    [InlineData("page")]
    public void HtmlColumnEdgeFloats_InlineDestinationMovesOnceAndPdfNavigationRemainsValid(string reference) {
        string html = ColumnEdgeStyle() + "<section class='columns'><p style='line-height:60px'>Call"
            + "<span id='figure' style='float:top;float-reference:" + reference + ";height:70px'>Figure</span></p>"
            + "<p><a href='#figure'>Link</a></p></section>";
        HtmlRenderDocument document = RenderColumnEdgeFixture(html);
        HtmlRenderNamedDestination destination = Assert.Single(document.Pages.SelectMany(p => p.Visuals)
            .OfType<HtmlRenderNamedDestination>(), d => d.Name == "figure");
        HtmlRenderText figure = Assert.Single(document.Pages.SelectMany(p => p.Visuals).OfType<HtmlRenderText>(), t => t.Text == "Figure");
        Assert.Equal(figure.X, destination.X, 3);
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdfBytes(new HtmlToPdfOptions { AutoFitWidePrintContent = false });
        Assert.Contains("Figure", OfficeIMO.Pdf.PdfReadDocument.Open(pdf).ExtractText());
    }

    private static HtmlRenderDocument RenderColumnEdgeFixture(string html) => HtmlRenderTestDriver.Render(html,
        new HtmlRenderOptions { Mode = HtmlRenderMode.Paged, HonorCssPageRules = true });

    private static string ColumnEdgeStyle() =>
        "<style>@page{size:240px 240px;margin:10px}body,p,aside{margin:0;font:10px/20px Arial}"
        + ".columns{column-count:2;column-gap:20px;column-fill:auto;height:100px}</style>";
}
