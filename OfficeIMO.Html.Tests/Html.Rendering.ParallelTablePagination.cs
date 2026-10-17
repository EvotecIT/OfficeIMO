using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("flex")]
    [InlineData("grid")]
    public void HtmlTables_ParallelFlowsRepeatTheirOwnSectionsAndMatchNativeRowPages(string display) {
        string html = ParallelReportStyle(display) + "<div>ReportStart</div>"
            + "<div class='columns'>" + ParallelReportTable("Left", 18) + ParallelReportTable("Right", 27)
            + "</div><div>ReportEnd</div>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions { Mode = HtmlRenderMode.Paged });

        Assert.Equal(4, rendered.Pages.Count);
        int[][] left = { Enumerable.Range(0, 7).ToArray(), Enumerable.Range(7, 8).ToArray(), Enumerable.Range(15, 3).ToArray(), Array.Empty<int>() };
        int[][] right = { Enumerable.Range(0, 7).ToArray(), Enumerable.Range(7, 8).ToArray(), Enumerable.Range(15, 8).ToArray(), Enumerable.Range(23, 4).ToArray() };
        for (int page = 0; page < rendered.Pages.Count; page++) {
            string[] text = rendered.Pages[page].Visuals.OfType<HtmlRenderText>().Select(item => item.Text).ToArray();
            Assert.Equal(left[page].Select(row => "Left" + row.ToString("D2")), text.Where(item => IsParallelReportRow(item, "Left")));
            Assert.Equal(right[page].Select(row => "Right" + row.ToString("D2")), text.Where(item => IsParallelReportRow(item, "Right")));
            Assert.Equal(left[page].Length > 0 ? 1 : 0, text.Count(item => item == "LeftHeader"));
            Assert.Equal(left[page].Length > 0 ? 1 : 0, text.Count(item => item == "LeftFooter"));
            Assert.Single(text, item => item == "RightHeader");
            Assert.Single(text, item => item == "RightFooter");
        }
        Assert.Contains(rendered.Pages[3].Visuals.OfType<HtmlRenderText>(), item => item.Text == "ReportEnd");
    }

    [Theory]
    [InlineData("flex")]
    [InlineData("grid")]
    public void HtmlTables_ParallelFlowKeepsItsFollowingContentWithoutWaitingForItsNeighbor(string display) {
        string html = ParallelReportStyle(display)
            + "<div class='columns'><section><div>LeftIntro</div>" + ParallelReportTable("Left", 4)
            + "<div>LeftAfter</div></section><section>" + ParallelReportTable("Right", 27)
            + "<div>RightAfter</div></section></div><div>OutsideAfter</div>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions { Mode = HtmlRenderMode.Paged });

        AssertParallelReportRowsAndSections(rendered, "Left", 4);
        AssertParallelReportRowsAndSections(rendered, "Right", 27);
        Assert.Contains(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), item => item.Text == "LeftAfter");
        HtmlRenderPage last = rendered.Pages.Last();
        HtmlRenderText rightAfter = Assert.Single(last.Visuals.OfType<HtmlRenderText>(), item => item.Text == "RightAfter");
        HtmlRenderText outsideAfter = Assert.Single(last.Visuals.OfType<HtmlRenderText>(), item => item.Text == "OutsideAfter");
        Assert.True(outsideAfter.Y >= rightAfter.Y + rightAfter.Height - 0.0001D);
        Assert.Single(rendered.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderText>(), item => item.Text == "LeftAfter");
    }

    [Theory]
    [InlineData("flex")]
    [InlineData("grid")]
    public void HtmlTables_LaterParallelRegionPreservesEarlierProjectedContent(string display) {
        string html = ParallelReportStyle(display) + "<div>ReportStart</div>"
            + "<div class='columns'>" + ParallelReportTable("FirstLeft", 18) + ParallelReportTable("FirstRight", 27)
            + "</div><div>BetweenRegions</div><div class='columns'>"
            + ParallelReportTable("SecondLeft", 18) + ParallelReportTable("SecondRight", 27)
            + "</div><div>ReportEnd</div>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions { Mode = HtmlRenderMode.Paged });

        foreach ((string name, int rows) in new[] { ("FirstLeft", 18), ("FirstRight", 27), ("SecondLeft", 18), ("SecondRight", 27) })
            AssertParallelReportRowsAndSections(rendered, name, rows);
        string[] allText = rendered.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderText>().Select(item => item.Text).ToArray();
        Assert.Single(allText, text => text == "BetweenRegions");
        Assert.Single(allText, text => text == "ReportEnd");
    }

    [Theory]
    [InlineData("flex")]
    [InlineData("grid")]
    public void HtmlTables_ParallelRepeatsObserveTheProjectedPageLimit(string display) {
        string html = ParallelReportStyle(display)
            + "<style>thead th,tfoot td{height:40px;line-height:40px}tbody td{height:10px;line-height:10px;font-size:8px}</style>"
            + "<div class='columns'>" + ParallelReportTable("Left", 27) + ParallelReportTable("Right", 27) + "</div>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions { Mode = HtmlRenderMode.Paged });
        Assert.Equal(3, rendered.Pages.Count);
        InvalidOperationException error = Assert.Throws<InvalidOperationException>(() =>
            HtmlRenderTestDriver.Render(html, new HtmlRenderOptions { Mode = HtmlRenderMode.Paged, MaxPageCount = 2 }));
        Assert.Contains("maximum page count", error.Message, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("flex")]
    [InlineData("grid")]
    public void HtmlTables_ParallelGeometryFallbackRetainsEveryRow(string display) {
        string html = ParallelReportStyle(display)
            + "<style>@page:left{margin-top:30px}</style><div class='columns'>"
            + ParallelReportTable("Left", 18) + ParallelReportTable("Right", 27) + "</div>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions { Mode = HtmlRenderMode.Paged });

        string[] allText = rendered.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderText>().Select(item => item.Text).ToArray();
        foreach (string name in new[] { "Left", "Right" }) {
            int count = name == "Left" ? 18 : 27;
            foreach (int row in Enumerable.Range(0, count)) Assert.Single(allText, text => text == name + row.ToString("D2"));
        }
        Assert.Contains(rendered.Diagnostics, item => item.Code == HtmlRenderDiagnosticCodes.PagePseudoGeometryPending);
    }

    [Theory]
    [InlineData("flex", 40, 5, 0)]
    [InlineData("grid", 40, 5, 0)]
    [InlineData("flex", 40, 5, 1)]
    [InlineData("grid", 40, 5, 1)]
    [InlineData("flex", 40, 5, 2)]
    [InlineData("grid", 40, 5, 2)]
    [InlineData("flex", -10, 7, 0)]
    [InlineData("grid", -10, 7, 0)]
    [InlineData("flex", -10, 7, 2)]
    [InlineData("grid", -10, 7, 2)]
    public void HtmlTables_ParallelBandsUseTheCollapsedMarginOrigin(
        string display, int margin, int firstPageRows, int wrapperDepth) {
        string open = string.Concat(Enumerable.Repeat("<section style='margin:0;width:auto'>", wrapperDepth));
        string close = string.Concat(Enumerable.Repeat("</section>", wrapperDepth));
        string html = ParallelReportStyle(display) + "<style>.intro{margin-bottom:" + margin
            + "px}.columns{margin-top:" + margin + "px}</style><main><div class='intro'>Start</div>"
            + open + "<div class='columns'>" + ParallelReportTable("Left", 18) + ParallelReportTable("Right", 27)
            + "</div>" + close + "<div>End</div></main>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions { Mode = HtmlRenderMode.Paged });

        Assert.Equal(4, rendered.Pages.Count);
        Assert.Equal(Enumerable.Range(0, firstPageRows).Select(row => "Left" + row.ToString("D2")),
            rendered.Pages[0].Visuals.OfType<HtmlRenderText>().Select(text => text.Text)
                .Where(text => IsParallelReportRow(text, "Left")));
        AssertParallelReportRowsAndSections(rendered, "Left", 18);
        AssertParallelReportRowsAndSections(rendered, "Right", 27);
    }

    [Theory]
    [InlineData("flex", HtmlRenderIntentProfile.PrintPaged, false)]
    [InlineData("grid", HtmlRenderIntentProfile.PrintPaged, false)]
    [InlineData("flex", HtmlRenderIntentProfile.ScreenMediaPaged, false)]
    [InlineData("grid", HtmlRenderIntentProfile.ScreenMediaPaged, false)]
    [InlineData("flex", HtmlRenderIntentProfile.PrintPaged, true)]
    [InlineData("grid", HtmlRenderIntentProfile.PrintPaged, true)]
    [InlineData("flex", HtmlRenderIntentProfile.ScreenMediaPaged, true)]
    [InlineData("grid", HtmlRenderIntentProfile.ScreenMediaPaged, true)]
    public void HtmlTables_ParallelItemFooterCannotConsumeNestedNeighborContent(
        string display, HtmlRenderIntentProfile profile, bool paragraphNeighbor) {
        string floats = paragraphNeighbor
            ? "<div class='left'>" + ParallelReportTable("A", 9) + "</div><p class='right'>"
                + string.Join("<br>", Enumerable.Range(0, 12).Select(row => "Tail" + row.ToString("D2"))) + "</p>"
            : "<div class='left'>" + ParallelReportTable("A", 3) + "</div><div class='right'>"
                + ParallelReportTable("B", 9) + "</div>";
        string html = ParallelReportStyle(display)
            + "<style>.columns{grid-template-columns:210px 80px}.nested{display:flow-root;width:210px}"
            + ".left{float:left;width:100px}.right{float:right;width:100px;margin:0}"
            + "table{width:100px}.other{width:80px}</style><div>Start</div><div class='columns'>"
            + "<section class='nested'>" + floats + "</section><section class='other'>"
            + ParallelReportTable("C", 12) + "</section></div><div>End</div>";
        var options = new HtmlRenderOptions { Mode = HtmlRenderMode.Paged };
        HtmlRenderDocument rendered = HtmlRenderEngine.Execute(HtmlConversionDocument.Parse(html),
            HtmlRenderRequest.Create(profile, HtmlRenderEncoder.DisplayList, options)).Document;
        string[] painted = rendered.Pages.SelectMany(page => EnumerateRenderVisuals(page.Scene))
            .OfType<HtmlRenderText>().Select(text => text.Text).ToArray();
        string[] expected = Enumerable.Range(0, paragraphNeighbor ? 9 : 3).Select(row => "A" + row.ToString("D2"))
            .Concat(Enumerable.Range(0, paragraphNeighbor ? 12 : 9)
                .Select(row => (paragraphNeighbor ? "Tail" : "B") + row.ToString("D2")))
            .Concat(Enumerable.Range(0, 12).Select(row => "C" + row.ToString("D2"))).Append("End").ToArray();
        string pdf = string.Concat(FloatTablePaginationPdfText(html, profile, options));
        foreach (string marker in expected) {
            Assert.Single(painted, text => text == marker);
            Assert.Equal(1, CountFloatPaginationMarker(pdf, marker));
        }
        string loss = paragraphNeighbor ? HtmlRenderDiagnosticCodes.TableFooterRepeatSuppressed
            : HtmlRenderDiagnosticCodes.TableHeaderRepeatSuppressed;
        Assert.Contains(rendered.Diagnostics, diagnostic => diagnostic.Code == loss
            && diagnostic.LossKind == OfficeIMO.OfficeConversionLossKind.Approximation);
        Assert.Throws<HtmlConversionException>(() => rendered.RequireNoLoss());
    }

    private static string ParallelReportStyle(string display) => "<style>@page{size:320px 220px;margin:10px}"
        + "body{margin:0;font-size:12px;line-height:20px}th,td{padding:0;font-weight:normal;height:20px}"
        + "table{border-collapse:collapse;width:140px}section{width:140px}"
        + ".columns{" + (display == "flex" ? "display:flex;align-items:flex-start" : "display:grid;grid-template-columns:140px 140px;align-items:start")
        + ";gap:10px}</style>";

    private static string ParallelReportTable(string name, int count) => "<table id='" + name + "'><thead><tr><th>"
        + name + "Header</th></tr></thead><tbody>"
        + string.Concat(Enumerable.Range(0, count).Select(row => "<tr><td>" + name + row.ToString("D2") + "</td></tr>"))
        + "</tbody><tfoot><tr><td>" + name + "Footer</td></tr></tfoot></table>";

    private static bool IsParallelReportRow(string text, string name) =>
        text.StartsWith(name, StringComparison.Ordinal) && int.TryParse(text.Substring(name.Length), out _);

    private static void AssertParallelReportRowsAndSections(HtmlRenderDocument rendered, string name, int count) {
        string[] allRows = rendered.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderText>()
            .Select(item => item.Text).Where(text => IsParallelReportRow(text, name)).ToArray();
        Assert.Equal(Enumerable.Range(0, count).Select(row => name + row.ToString("D2")), allRows);
        foreach (HtmlRenderPage page in rendered.Pages) {
            string[] text = page.Visuals.OfType<HtmlRenderText>().Select(item => item.Text).ToArray();
            int expected = text.Any(item => IsParallelReportRow(item, name)) ? 1 : 0;
            Assert.Equal(expected, text.Count(item => item == name + "Header"));
            Assert.Equal(expected, text.Count(item => item == name + "Footer"));
        }
    }
}
