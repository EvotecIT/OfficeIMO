using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfColumnContentBalancingTests {
    [Theory]
    [InlineData("list", 21)]
    [InlineData("list", 43)]
    [InlineData("table", 21)]
    [InlineData("table", 43)]
    public void Columns_BalancesItemsAndRowsOnTheFinalPhysicalPage(string kind, int count) {
        string prefix = kind == "list" ? "List" : "Table";
        byte[] bytes = PdfDocument.Create(Options()).Columns(content => {
            if (kind == "list") content.RichNumbered(Enumerable.Range(1, count)
                .Select(index => new PdfListItem(prefix + index.ToString("D3"))), style: ListStyle());
            else content.Table(Enumerable.Range(1, count).Select(index => new[] {
                new PdfTableCell(prefix + index.ToString("D3"))
            }), style: TableStyle());
        }, new PdfMultiColumnOptions { Gap = 20, BalanceLastPage = true })
            .Paragraph(p => p.Text("AfterColumns"), style: ParagraphStyle()).ToBytes();
        using var read = UglyToad.PdfPig.PdfDocument.Open(bytes);
        int pages = count == 21 ? 1 : 2;
        int leftLast = count == 21 ? 11 : 38;
        Assert.Equal(pages, read.NumberOfPages);
        Assert.InRange(X(read, pages, prefix + leftLast.ToString("D3")), 40, kind == "list" ? 80 : 40.1);
        Assert.InRange(X(read, pages, prefix + (leftLast + 1).ToString("D3")), 260, kind == "list" ? 300 : 260.1);
        Assert.InRange(X(read, pages, "AfterColumns"), 39.9, 40.1);
        int firstOnFinalPage = count == 21 ? 1 : 33;
        double finalColumnHeight = count == 21 ? 220 : 120;
        Assert.InRange(Y(read, pages, prefix + firstOnFinalPage.ToString("D3")) - Y(read, pages, "AfterColumns"),
            finalColumnHeight - 0.1, finalColumnHeight + 0.1);
        string rendered = string.Concat(Enumerable.Range(1, pages).Select(page => read.GetPage(page).Text));
        var markers = System.Text.RegularExpressions.Regex.Matches(rendered, prefix + @"(\d{3})")
            .Cast<System.Text.RegularExpressions.Match>().Select(match => int.Parse(match.Groups[1].Value));
        Assert.Equal(Enumerable.Range(1, count), markers);
        if (kind == "list") {
            var numbers = System.Text.RegularExpressions.Regex.Matches(rendered, @"(\d+)\.\s*List(\d{3})")
                .Cast<System.Text.RegularExpressions.Match>().Select(match => int.Parse(match.Groups[1].Value));
            Assert.Equal(Enumerable.Range(1, count), numbers);
        }
    }

    [Theory]
    [InlineData(21)]
    [InlineData(43)]
    public void Columns_BalancesAContinuingListItemWithoutRepeatingItsMarker(int lines) {
        var item = new PdfListItem(string.Join("\n", Enumerable.Range(1, lines).Select(index => "Line" + index.ToString("D3"))),
            bookmarkName: "ListStart", marker: "OnlyMarker");
        byte[] bytes = PdfDocument.Create(Options()).Columns(content => content.RichNumbered(new[] { item }, style: ListStyle()),
            new PdfMultiColumnOptions { Gap = 20, BalanceLastPage = true })
            .Paragraph(p => p.Text("AfterColumns"), style: ParagraphStyle()).ToBytes();
        using var read = UglyToad.PdfPig.PdfDocument.Open(bytes);
        int pages = lines == 21 ? 1 : 2;
        int leftLast = lines == 21 ? 11 : 38;
        Assert.Equal(pages, read.NumberOfPages);
        Assert.InRange(X(read, pages, "Line" + leftLast.ToString("D3")), 40, 150);
        Assert.InRange(X(read, pages, "Line" + (leftLast + 1).ToString("D3")), 260, 370);
        string rendered = string.Concat(Enumerable.Range(1, pages).Select(page => read.GetPage(page).Text));
        Assert.Single(System.Text.RegularExpressions.Regex.Matches(rendered, "OnlyMarker").Cast<System.Text.RegularExpressions.Match>());
        Assert.Equal(Enumerable.Range(1, lines), System.Text.RegularExpressions.Regex.Matches(rendered, @"Line(\d{3})")
            .Cast<System.Text.RegularExpressions.Match>().Select(match => int.Parse(match.Groups[1].Value)));
        Assert.Contains("ListStart", System.Text.Encoding.ASCII.GetString(bytes));
    }

    [Theory]
    [InlineData(21)]
    [InlineData(43)]
    public void Columns_BalancedTableReservesTheRepeatedHeaderInEachFrame(int bodyRows) {
        var style = TableStyle();
        style.HeaderRowCount = 1;
        style.RepeatHeaderRowCount = 1;
        style.HeaderBold = false;
        style.HeaderFontSize = 12;
        byte[] bytes = PdfDocument.Create(Options()).Columns(content => content.Table(
            new[] { new[] { new PdfTableCell("Header") } }.Concat(Enumerable.Range(1, bodyRows)
                .Select(index => new[] { new PdfTableCell("Table" + index.ToString("D3")) })), style: style),
            new PdfMultiColumnOptions { Gap = 20, BalanceLastPage = true })
            .Paragraph(p => p.Text("AfterColumns"), style: ParagraphStyle()).ToBytes();
        using var read = UglyToad.PdfPig.PdfDocument.Open(bytes);
        int pages = bodyRows == 21 ? 1 : 2;
        int firstFinal = bodyRows == 21 ? 1 : 31;
        int leftLast = bodyRows == 21 ? 11 : 37;
        Assert.Equal(pages, read.NumberOfPages);
        Assert.Equal(2, read.GetPage(pages).GetWords().Count(word => word.Text == "Header"));
        Assert.InRange(X(read, pages, "Table" + leftLast.ToString("D3")), 39.9, 40.1);
        Assert.InRange(X(read, pages, "Table" + (leftLast + 1).ToString("D3")), 259.9, 260.1);
        var headers = read.GetPage(pages).GetWords().Where(word => word.Text == "Header").ToArray();
        Assert.InRange(Math.Abs(headers[0].BoundingBox.Bottom - headers[1].BoundingBox.Bottom), 0, 0.1);
        Assert.InRange(headers[0].BoundingBox.Bottom - Y(read, pages, "Table" + firstFinal.ToString("D3")), 19.9, 20.1);
        string rendered = string.Concat(Enumerable.Range(1, pages).Select(page => read.GetPage(page).Text));
        Assert.Equal(Enumerable.Range(1, bodyRows), System.Text.RegularExpressions.Regex.Matches(rendered, @"Table(\d{3})")
            .Cast<System.Text.RegularExpressions.Match>().Select(match => int.Parse(match.Groups[1].Value)));
    }

    [Theory]
    [InlineData("list")]
    [InlineData("table")]
    public void Columns_BalancingPreservesAnAtomicKeptElement(string kind) {
        byte[] bytes = PdfDocument.Create(Options()).Columns(content => {
            content.Paragraph(p => p.Text("Before1\nBefore2\nBefore3"), style: ParagraphStyle());
            if (kind == "list") {
                var style = ListStyle(); style.KeepTogether = true;
                content.RichNumbered(Enumerable.Range(1, 5).Select(index => new PdfListItem("Kept" + index)), style: style);
            } else {
                var style = TableStyle(); style.KeepTogether = true;
                content.Table(Enumerable.Range(1, 5).Select(index => new[] { new PdfTableCell("Kept" + index) }), style: style);
            }
        }, new PdfMultiColumnOptions { Gap = 20, BalanceLastPage = true }).ToBytes();
        using var read = UglyToad.PdfPig.PdfDocument.Open(bytes);
        Assert.Equal(1, read.NumberOfPages);
        Assert.InRange(X(read, 1, "Before3"), 39.9, 40.1);
        Assert.InRange(X(read, 1, "Kept1"), 260, kind == "list" ? 300 : 260.1);
        Assert.InRange(X(read, 1, "Kept5"), 260, kind == "list" ? 300 : 260.1);
    }

    private static double X(UglyToad.PdfPig.PdfDocument read, int page, string text) =>
        read.GetPage(page).GetWords().Single(word => word.Text == text || word.Text.EndsWith("." + text)).BoundingBox.Left;

    private static double Y(UglyToad.PdfPig.PdfDocument read, int page, string text) =>
        read.GetPage(page).GetWords().Single(word => word.Text == text || word.Text.EndsWith("." + text)).BoundingBox.Bottom;

    private static PdfOptions Options() => new() {
        PageWidth = 500, PageHeight = 400, MarginLeft = 40, MarginRight = 40,
        MarginTop = 40, MarginBottom = 40, DefaultFontSize = 12, CompressContentStreams = false
    };

    private static PdfParagraphStyle ParagraphStyle() => new() {
        LineSpacing = PdfLineSpacing.Exactly(20), SpacingBefore = 0, SpacingAfter = 0, WidowControl = false
    };

    private static PdfListStyle ListStyle() => new() {
        LineSpacing = PdfLineSpacing.Exactly(20), ItemSpacing = 0, SpacingBefore = 0, SpacingAfter = 0
    };

    private static PdfTableStyle TableStyle() => new() {
        HeaderRowCount = 0, CellPaddingX = 0, CellPaddingY = 0, FontSize = 12, LineHeight = 20D / 12D,
        BorderWidth = 0, RowSeparatorWidth = 0, SpacingBefore = 0, SpacingAfter = 0,
        MinimumBodyRowsOnFirstPage = 0, MinimumBodyRowsOnLastPage = 0
    };
}
