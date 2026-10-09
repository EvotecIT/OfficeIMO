using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using Xunit;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData(HtmlRenderIntentProfile.PrintPaged, "")]
    [InlineData(HtmlRenderIntentProfile.ScreenMediaPaged, "")]
    [InlineData(HtmlRenderIntentProfile.PrintPaged, "text")]
    [InlineData(HtmlRenderIntentProfile.ScreenMediaPaged, "text")]
    [InlineData(HtmlRenderIntentProfile.PrintPaged, "float")]
    [InlineData(HtmlRenderIntentProfile.ScreenMediaPaged, "float")]
    [InlineData(HtmlRenderIntentProfile.PrintPaged, "image")]
    [InlineData(HtmlRenderIntentProfile.ScreenMediaPaged, "image")]
    public void HtmlFloatPagination_KeepsEveryTableRowAndNeighborVisible(
        HtmlRenderIntentProfile profile, string neighbor) {
        string side = neighbor switch {
            "text" => "<p style='margin:0;font-size:12px;line-height:20px'>"
                + string.Join("<br>", Enumerable.Range(0, 7).Select(index => "TEXT_" + (char)('A' + index))) + "</p>",
            "float" => "<div style='float:right;width:100px'><p style='margin:0'>SIDE_A<br>SIDE_B</p></div>",
            "image" => "<div style='float:right'><svg width='80' height='100'>"
                + "<rect width='80' height='100' fill='orange'/></svg></div>",
            _ => ""
        };
        string html = FloatTablePaginationHtml(FloatTableRows(1, 7), side);
        HtmlRenderDocument rendered = RenderFloatTablePagination(html, profile);

        Assert.True(rendered.Pages.Count >= 2);
        var markers = Enumerable.Range(1, 7).Select(index => $"ROW_{index:00}").ToList();
        if (neighbor == "text") markers.AddRange(Enumerable.Range(0, 7).Select(index => "TEXT_" + (char)('A' + index)));
        if (neighbor == "float") markers.AddRange(new[] { "SIDE_A", "SIDE_B" });
        AssertFloatPaginationMarkers(rendered, markers);

        if (neighbor == "image") {
            HtmlRenderDrawing image = Assert.Single(rendered.Pages
                .SelectMany(page => EnumerateRenderVisuals(page.Scene)).OfType<HtmlRenderDrawing>());
            Assert.Equal(100D, image.Height, 6);
        }
        byte[] pdf = HtmlConversionDocument.Parse(html).RenderToPdfResult(
            HtmlRenderRequest.Create(profile, HtmlRenderEncoder.Pdf, FloatTablePaginationOptions())).ToBytes();
        string pdfText = string.Concat(PdfCore.PdfReadDocument.Open(pdf).ExtractText().Where(character => !char.IsWhiteSpace(character)));
        foreach (string marker in markers) {
            int first = pdfText.IndexOf(marker, StringComparison.Ordinal);
            Assert.True(first >= 0, "The PDF lost " + marker + ".");
            Assert.Equal(-1, pdfText.IndexOf(marker, first + marker.Length, StringComparison.Ordinal));
        }
    }

    [Theory]
    [InlineData(HtmlRenderIntentProfile.PrintPaged)]
    [InlineData(HtmlRenderIntentProfile.ScreenMediaPaged)]
    public void HtmlFloatPagination_KeepsAFittingRowGroupTogetherAfterPrelude(HtmlRenderIntentProfile profile) {
        string rows = "<tbody style='break-inside:avoid'>" + FloatTableRows(1, 3) + "</tbody>"
            + "<tbody>" + FloatTableRows(4, 4) + "</tbody>";
        string html = FloatTablePaginationHtml(rows, "", "<div style='height:70px'>LEAD</div>");
        HtmlRenderDocument rendered = RenderFloatTablePagination(html, profile);

        AssertFloatPaginationMarkers(rendered, Enumerable.Range(1, 7).Select(index => $"ROW_{index:00}").Append("LEAD"));
        int[] rowPages = rendered.Pages.SelectMany(page => EnumerateRenderVisuals(page.Scene)
            .OfType<HtmlRenderText>().Where(text => text.Text is "ROW_01" or "ROW_02" or "ROW_03")
            .Select(_ => page.PageNumber)).ToArray();
        Assert.Equal(3, rowPages.Length);
        Assert.Single(rowPages.Distinct());
        Assert.True(rowPages[0] > 1);
    }

    private static string FloatTableRows(int first, int count) => string.Concat(
        Enumerable.Range(first, count).Select(index => $"<tr><td>ROW_{index:00}</td></tr>"));

    [Theory]
    [InlineData(HtmlRenderIntentProfile.PrintPaged, "anonymous")]
    [InlineData(HtmlRenderIntentProfile.ScreenMediaPaged, "anonymous")]
    [InlineData(HtmlRenderIntentProfile.PrintPaged, "nested-inline")]
    [InlineData(HtmlRenderIntentProfile.PrintPaged, "outside-list")]
    [InlineData(HtmlRenderIntentProfile.PrintPaged, "interrupted-inline")]
    public void HtmlFloatPagination_RepeatsTableHeaderAndFooter(HtmlRenderIntentProfile profile, string wrapper) {
        string rows = "<thead><tr><th>HEAD</th></tr></thead><tbody>" + FloatTableRows(1, 7)
            + "</tbody><tfoot><tr><td>FOOT</td></tr></tfoot>";
        string html = FloatTablePaginationHtml(rows, "", wrapper: wrapper);
        void Check(string input) {
            HtmlRenderDocument rendered = RenderFloatTablePagination(input, profile);
            HtmlRenderPage[] bodyPages = rendered.Pages.Where(page => EnumerateRenderVisuals(page.Scene)
                .OfType<HtmlRenderText>().Any(text => text.Text.StartsWith("ROW_", StringComparison.Ordinal))).ToArray();
            Assert.True(bodyPages.Length > 1);
            foreach (HtmlRenderPage page in bodyPages) {
                HtmlRenderText[] text = EnumerateRenderVisuals(page.Scene).OfType<HtmlRenderText>().ToArray();
                HtmlRenderText head = Assert.Single(text, item => item.Text == "HEAD");
                HtmlRenderText foot = Assert.Single(text, item => item.Text == "FOOT");
                HtmlRenderText[] body = text.Where(item => item.Text.StartsWith("ROW_", StringComparison.Ordinal)).ToArray();
                Assert.True(head.Y + head.Height <= body.Min(item => item.Y) + 0.0001D);
                Assert.True(foot.Y >= body.Max(item => item.Y + item.Height) - 0.0001D);
                OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(page.CreateDrawing(), 1D, OfficeColor.White);
                Assert.True(FloatPaintContainsGlyphInk(raster, head));
                Assert.True(FloatPaintContainsGlyphInk(raster, foot));
            }
            AssertFloatPaginationMarkers(rendered, Enumerable.Range(1, 7).Select(index => $"ROW_{index:00}"));
            IReadOnlyList<string> pdfPages = FloatTablePaginationPdfText(input, profile);
            Assert.Equal(rendered.Pages.Count, pdfPages.Count);
            foreach (string pageText in pdfPages.Where(text => text.Contains("ROW_", StringComparison.Ordinal))) {
                Assert.Equal(1, CountFloatPaginationMarker(pageText, "HEAD"));
                Assert.Equal(1, CountFloatPaginationMarker(pageText, "FOOT"));
            }
            foreach (int index in Enumerable.Range(1, 7)) Assert.Equal(1, CountFloatPaginationMarker(string.Concat(pdfPages), $"ROW_{index:00}"));
        }
        Check(html.Replace("float:left;", ""));
        Check(html);
    }

    [Theory]
    [InlineData(HtmlRenderIntentProfile.PrintPaged, "before")]
    [InlineData(HtmlRenderIntentProfile.ScreenMediaPaged, "before")]
    [InlineData(HtmlRenderIntentProfile.PrintPaged, "after")]
    [InlineData(HtmlRenderIntentProfile.ScreenMediaPaged, "after")]
    public void HtmlFloatPagination_PreservesForcedRowBreak(HtmlRenderIntentProfile profile, string breakPosition) {
        string rows = breakPosition == "before"
            ? FloatTableRows(1, 1) + "<tr style='break-before:page'><td>ROW_02</td></tr>" + FloatTableRows(3, 1)
            : "<tr style='break-after:page'><td>ROW_01</td></tr>" + FloatTableRows(2, 2);
        string html = FloatTablePaginationHtml(rows, "");
        void Check(string input) {
            HtmlRenderDocument rendered = RenderFloatTablePagination(input, profile);
            Assert.Equal(2, rendered.Pages.Count);
            Assert.Contains(EnumerateRenderVisuals(rendered.Pages[0].Scene).OfType<HtmlRenderText>(), text => text.Text == "ROW_01");
            Assert.DoesNotContain(EnumerateRenderVisuals(rendered.Pages[0].Scene).OfType<HtmlRenderText>(), text => text.Text == "ROW_02");
            Assert.Contains(EnumerateRenderVisuals(rendered.Pages[1].Scene).OfType<HtmlRenderText>(), text => text.Text == "ROW_02");
            AssertFloatPaginationMarkers(rendered, Enumerable.Range(1, 3).Select(index => $"ROW_{index:00}"));
            IReadOnlyList<string> pdfPages = FloatTablePaginationPdfText(input, profile);
            Assert.Equal(2, pdfPages.Count);
            Assert.Contains("ROW_01", pdfPages[0], StringComparison.Ordinal);
            Assert.DoesNotContain("ROW_02", pdfPages[0], StringComparison.Ordinal);
            Assert.Contains("ROW_02", pdfPages[1], StringComparison.Ordinal);
        }
        Check(html.Replace("float:left;", ""));
        Check(html);
    }

    [Theory]
    [InlineData(1, 9)]
    [InlineData(3, 9)]
    [InlineData(9, 9)]
    public void HtmlFloatPagination_ParallelTablesRetainEveryBodyRow(int leftCount, int rightCount) {
        static string Rows(string prefix, int count) => string.Concat(Enumerable.Range(1, count)
            .Select(index => $"<tr><td>{prefix}{index:00}</td></tr>"));
        static string Table(string prefix, int count) => "<table id='" + prefix + "'><thead><tr><th>" + prefix + "HEAD</th></tr></thead><tbody>"
            + Rows(prefix, count) + "</tbody><tfoot><tr><td>" + prefix + "FOOT</td></tr></tfoot></table>";
        string html = FloatTablePaginationHtml("", "")
            .Replace("<div class='float'><table></table></div>",
                "<div class='float'>" + Table("A", leftCount) + "</div>"
                + "<div style='float:right;width:100px'>" + Table("B", rightCount) + "</div>");
        HtmlRenderOptions options = FloatTablePaginationOptions();
        if (leftCount == 1) {
            options.PageSize = new OfficePageSize(320D / HtmlRenderOptions.CssPixelsPerInch, 400D / HtmlRenderOptions.CssPixelsPerInch);
            options.ViewportHeight = 400D;
        }
        HtmlRenderDocument rendered = HtmlRenderEngine.Execute(HtmlConversionDocument.Parse(html),
            HtmlRenderRequest.Create(HtmlRenderIntentProfile.PrintPaged, HtmlRenderEncoder.DisplayList, options)).Document;
        Assert.Contains(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.TableHeaderRepeatSuppressed
            && diagnostic.LossKind == OfficeConversionLossKind.Approximation);
        Assert.Throws<HtmlConversionException>(() => rendered.RequireNoLoss());
        if (leftCount == 1) {
            Assert.Single(rendered.Pages, page => EnumerateRenderVisuals(page.Scene)
                .OfType<HtmlRenderText>().Any(text => text.Text is "AHEAD" or "A01" or "AFOOT"));
            Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Source == "table#A"
                && diagnostic.Detail?.StartsWith("parallel-table-repetition;", StringComparison.Ordinal) == true);
        }
        foreach (string marker in Enumerable.Range(1, leftCount).Select(index => $"A{index:00}")
            .Concat(Enumerable.Range(1, rightCount).Select(index => $"B{index:00}"))) {
            Assert.Single(rendered.Pages.SelectMany(page => EnumerateRenderVisuals(page.Scene)).OfType<HtmlRenderText>(), text => text.Text == marker);
        }
        string pdfText = string.Concat(FloatTablePaginationPdfText(html, HtmlRenderIntentProfile.PrintPaged, options));
        foreach (string marker in Enumerable.Range(1, leftCount).Select(index => $"A{index:00}")
            .Concat(Enumerable.Range(1, rightCount).Select(index => $"B{index:00}"))) Assert.Equal(1, CountFloatPaginationMarker(pdfText, marker));
    }

    [Theory]
    [InlineData(HtmlRenderIntentProfile.PrintPaged, false)]
    [InlineData(HtmlRenderIntentProfile.ScreenMediaPaged, false)]
    [InlineData(HtmlRenderIntentProfile.PrintPaged, true)]
    public void HtmlFloatPagination_FooterCannotConsumeNeighboringParagraph(
        HtmlRenderIntentProfile profile, bool rootFlow) {
        string neighbor = "<p>" + string.Join("<br>", Enumerable.Range(1, 12).Select(index => $"TAIL_{index:00}")) + "</p>";
        string rows = "<thead><tr><th>HEAD</th></tr></thead><tbody>" + FloatTableRows(1, 4)
            + "</tbody><tfoot><tr><td>FOOT</td></tr></tfoot>";
        string html = FloatTablePaginationHtml(rows, neighbor)
            .Replace("font:16px Pinned", "font:16px/20px Pinned")
            .Replace("td{padding:5px}", "p{margin:0}td{padding:5px 0}th{padding:0}");
        if (rootFlow) html = html.Replace("<div style='overflow:hidden'>", "<div style='display:contents'>");
        HtmlRenderDocument rendered = RenderFloatTablePagination(html, profile);
        Assert.Contains(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.TableFooterRepeatSuppressed
            && diagnostic.LossKind == OfficeConversionLossKind.Approximation);
        Assert.Throws<HtmlConversionException>(() => rendered.RequireNoLoss());
        var markers = Enumerable.Range(1, 12).Select(index => $"TAIL_{index:00}")
            .Concat(Enumerable.Range(1, 4).Select(index => $"ROW_{index:00}")).Append("FOOT");
        string pdfText = string.Concat(FloatTablePaginationPdfText(html, profile));
        foreach (string marker in markers) {
            Assert.Single(rendered.Pages.SelectMany(page => EnumerateRenderVisuals(page.Scene))
                .OfType<HtmlRenderText>(), text => text.Text == marker);
            Assert.Equal(1, CountFloatPaginationMarker(pdfText, marker));
        }
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.VisualFragmentUnsupported);
    }

    private static IReadOnlyList<string> FloatTablePaginationPdfText(string html, HtmlRenderIntentProfile profile, HtmlRenderOptions? options = null) {
        byte[] pdf = HtmlConversionDocument.Parse(html).RenderToPdfResult(
            HtmlRenderRequest.Create(profile, HtmlRenderEncoder.Pdf, options ?? FloatTablePaginationOptions())).ToBytes();
        return PdfCore.PdfTextExtractor.ExtractTextByPage(pdf)
            .Select(text => string.Concat(text.Where(character => !char.IsWhiteSpace(character)))).ToArray();
    }

    private static int CountFloatPaginationMarker(string text, string marker) =>
        (text.Length - text.Replace(marker, "").Length) / marker.Length;

    private static string FloatTablePaginationHtml(string rows, string neighbor, string prelude = "", string wrapper = "anonymous") {
        string table = "<table>" + rows + "</table>";
        string floated = wrapper switch {
            "nested-inline" => "<span><span class='float'>" + table + "</span></span>",
            "outside-list" => "<ul style='margin:0;padding:0'><li><span><span class='float'>" + table + "</span></span></li></ul>",
            "interrupted-inline" => "<span><span class='float'>" + table + "</span><div>AFTER</div></span>",
            _ => "<div class='float'>" + table + "</div>"
        };
        return "<style>body{margin:0;font:16px Pinned}table{border-collapse:collapse}td{padding:5px}"
            + ".float{float:left;width:100px}</style><div style='overflow:hidden'>" + prelude + floated + neighbor + "</div>";
    }

    private static HtmlRenderOptions FloatTablePaginationOptions() {
        HtmlRenderOptions options = TableIntrinsicOptions();
        options.AllowSystemFontFallback = false;
        options.PageSize = new OfficePageSize(320D / HtmlRenderOptions.CssPixelsPerInch, 230D / HtmlRenderOptions.CssPixelsPerInch);
        options.ViewportWidth = 320D;
        options.ViewportHeight = 230D;
        options.Margins = HtmlRenderMargins.All(48D);
        options.HonorCssPageRules = false;
        return options;
    }

    private static HtmlRenderDocument RenderFloatTablePagination(string html, HtmlRenderIntentProfile profile) =>
        HtmlRenderEngine.Execute(HtmlConversionDocument.Parse(html),
            HtmlRenderRequest.Create(profile, HtmlRenderEncoder.DisplayList, FloatTablePaginationOptions())).Document;

    private static void AssertFloatPaginationMarkers(HtmlRenderDocument rendered, IEnumerable<string> markers) {
        foreach (string marker in markers) {
            HtmlRenderPage page = Assert.Single(rendered.Pages, item =>
                EnumerateRenderVisuals(item.Scene).OfType<HtmlRenderText>().Any(text => text.Text == marker));
            HtmlRenderText text = Assert.Single(EnumerateRenderVisuals(page.Scene).OfType<HtmlRenderText>(), item => item.Text == marker);
            OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(page.CreateDrawing(), 1D, OfficeColor.White);
            Assert.True(FloatPaintContainsGlyphInk(raster, text), "Retained text must also paint " + marker + ".");
        }
        rendered.RequireNoLoss();
    }
}
