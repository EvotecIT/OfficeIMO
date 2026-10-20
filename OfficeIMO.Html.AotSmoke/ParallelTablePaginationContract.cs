using System.Globalization;
using OfficeIMO.Html;

/// <summary>Exercises independent table continuation through the trimmed native layout entry point.</summary>
internal static class ParallelTablePaginationContract {
    internal static void Verify() {
        foreach (string display in new[] { "flex", "grid" }) {
            string columns = display == "flex"
                ? "display:flex;align-items:flex-start"
                : "display:grid;grid-template-columns:140px 140px;align-items:start";
            string html = "<style>@page{size:320px 220px;margin:10px}"
                + "body{margin:0;font:12px/20px Arial}th,td{padding:0;font-weight:normal;height:20px}"
                + "table{border-collapse:collapse;width:140px}.columns{" + columns + ";gap:10px}</style>"
                + "<div>ReportStart</div><div class='columns'>"
                + Table("Left", 18) + Table("Right", 27) + "</div><div>ReportEnd</div>";
            HtmlRenderDocument rendered = HtmlRenderEngine.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
                Mode = HtmlRenderMode.Paged,
                DefaultFontFamily = "Arial",
                UserAgentStyles = HtmlRenderUserAgentStyleMode.Browser
            });
            if (rendered.Pages.Count != 4) {
                throw new InvalidOperationException($"The NativeAOT {display} parallel tables lost their four-page layout.");
            }

            VerifyTable(rendered, display, "Left", 18, new[] { 7, 8, 3, 0 });
            VerifyTable(rendered, display, "Right", 27, new[] { 7, 8, 8, 4 });
            string[] text = rendered.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderText>()
                .Select(item => item.Text).ToArray();
            if (text.Count(value => value == "ReportStart") != 1
                || text.Count(value => value == "ReportEnd") != 1
                || !rendered.Pages[3].Visuals.OfType<HtmlRenderText>().Any(item => item.Text == "ReportEnd")) {
                throw new InvalidOperationException($"The NativeAOT {display} parallel tables lost their following content.");
            }
        }
    }

    private static void VerifyTable(HtmlRenderDocument rendered, string display, string name, int count, int[] pageRows) {
        string[] expected = Enumerable.Range(0, count).Select(row => RowLabel(name, row)).ToArray();
        string[] actual = rendered.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderText>()
            .Select(item => item.Text).Where(text => IsRow(text, name)).ToArray();
        if (!actual.SequenceEqual(expected)) {
            throw new InvalidOperationException($"The NativeAOT {display} {name} table lost exact-once source row order.");
        }
        for (int page = 0; page < rendered.Pages.Count; page++) {
            string[] text = rendered.Pages[page].Visuals.OfType<HtmlRenderText>().Select(item => item.Text).ToArray();
            int repeats = pageRows[page] > 0 ? 1 : 0;
            if (text.Count(value => IsRow(value, name)) != pageRows[page]
                || text.Count(value => value == name + "Header") != repeats
                || text.Count(value => value == name + "Footer") != repeats) {
                throw new InvalidOperationException(
                    $"The NativeAOT {display} {name} table lost its independent row/header/footer continuation on page {page + 1}.");
            }
        }
    }

    private static string Table(string name, int count) => "<table><thead><tr><th>" + name
        + "Header</th></tr></thead><tbody>"
        + string.Concat(Enumerable.Range(0, count).Select(row => "<tr><td>" + RowLabel(name, row) + "</td></tr>"))
        + "</tbody><tfoot><tr><td>" + name + "Footer</td></tr></tfoot></table>";

    private static string RowLabel(string name, int row) => name + row.ToString("D2", CultureInfo.InvariantCulture);

    private static bool IsRow(string text, string name) => text.StartsWith(name, StringComparison.Ordinal)
        && int.TryParse(text.AsSpan(name.Length), NumberStyles.None, CultureInfo.InvariantCulture, out _);
}
