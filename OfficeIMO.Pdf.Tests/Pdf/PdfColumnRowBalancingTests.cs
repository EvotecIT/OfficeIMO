using OfficeIMO.Pdf;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfColumnRowBalancingTests {
    [Theory]
    [InlineData(21, true)]
    [InlineData(43, true)]
    [InlineData(21, false)]
    [InlineData(43, false)]
    public void Columns_BalanceInitialAndContinuingRowsOnlyWhenEnabled(int lines, bool balanceRowLines) {
        using var pdf = PdfPigDocument.Open(Render(new[] { new[] { Cell("Line", lines) } }, TableStyle(), balanceRowLines));
        int pages = lines == 21 && balanceRowLines ? 1 : 2;
        Assert.Equal(pages, pdf.NumberOfPages);
        int leftLast = lines == 21 ? (balanceRowLines ? 11 : 16) : (balanceRowLines ? 38 : 43);
        Assert.InRange(X(pdf, lines == 21 ? 1 : 2, $"Line{leftLast:D3}"), 39.9, 40.1);
        if (balanceRowLines || lines == 21)
            Assert.InRange(X(pdf, lines == 21 ? 1 : 2, $"Line{leftLast + 1:D3}"), 259.9, 260.1);
        Assert.InRange(X(pdf, pages, "AfterColumns"), 39.9, 40.1);
        AssertMarkers(pdf, "Line", lines);
    }

    [Theory]
    [InlineData(21)]
    [InlineData(43)]
    public void Columns_RowBalancingReservesPaddingInEachFragment(int lines) {
        var style = TableStyle();
        style.CellPaddingY = 4;
        using var pdf = PdfPigDocument.Open(Render(new[] { new[] { Cell("Padded", lines) } }, style));
        int pages = lines == 21 ? 1 : 2;
        int leftLast = lines == 21 ? 11 : 37; // Full-height columns hold fifteen padded lines each.
        Assert.Equal(pages, pdf.NumberOfPages);
        Assert.InRange(X(pdf, pages, $"Padded{leftLast:D3}"), 39.9, 40.1);
        Assert.InRange(X(pdf, pages, $"Padded{leftLast + 1:D3}"), 259.9, 260.1);
        Assert.InRange(Y(pdf, pages, $"Padded{leftLast:D3}") - Y(pdf, pages, "AfterColumns"), 23.9, 24.1);
        AssertMarkers(pdf, "Padded", lines);
    }

    [Theory]
    [InlineData(21)]
    [InlineData(43)]
    public void Columns_RowBalancingReservesARepeatedHeader(int lines) {
        var style = TableStyle();
        style.HeaderRowCount = 1; style.RepeatHeaderRowCount = 1;
        style.HeaderFontSize = 12; style.HeaderBold = false;
        var rows = new[] { new[] { Cell("Header", 1) }, new[] { Cell("Body", lines) } };
        using var pdf = PdfPigDocument.Open(Render(rows, style));
        int pages = lines == 21 ? 1 : 2;
        int leftLast = lines == 21 ? 11 : 37; // Repeated headers leave fifteen body lines in each full column.
        Assert.Equal(pages, pdf.NumberOfPages);
        Assert.Equal(2, pdf.GetPage(pages).GetWords().Count(word => word.Text == "Header001"));
        Assert.InRange(X(pdf, pages, $"Body{leftLast:D3}"), 39.9, 40.1);
        Assert.InRange(X(pdf, pages, $"Body{leftLast + 1:D3}"), 259.9, 260.1);
        AssertMarkers(pdf, "Body", lines);
    }

    [Fact]
    public void Columns_RowBalancingMeasuresAccumulatedCellHeightsInsteadOfSummingLineMaxima() {
        PdfTableCell MakeCell(string prefix, bool right) {
            var paragraphs = Enumerable.Range(1, 12).Select(index => {
                double height = right ? (index <= 6 && index % 2 == 1 ? 10 : 30)
                    : (index <= 6 && index % 2 == 1 ? 30 : 10);
                return new PdfTableCellParagraph(new[] { PdfTextRun.Normal($"{prefix}{index:D3}", fontSize: 8) },
                    fontSize: 8, lineSpacing: PdfLineSpacing.Exactly(height), widowControl: false);
            }).ToArray();
            return new PdfTableCell(paragraphs.SelectMany(paragraph => paragraph.Runs), paragraphs);
        }
        var rows = new[] { new[] { MakeCell("Left", false), MakeCell("Right", true) } };
        using var pdf = PdfPigDocument.Open(Render(rows, TableStyle()));
        using var sequential = PdfPigDocument.Open(PdfDocument.Create(Options()).Table(rows, style: TableStyle())
            .Paragraph(p => p.Text("AfterColumns"), style: ParagraphStyle()).ToBytes());
        Assert.Equal(1, pdf.NumberOfPages);
        Assert.InRange(X(pdf, 1, "Left007"), 39.9, 40.1);
        Assert.InRange(X(pdf, 1, "Left008"), 259.9, 260.1);
        Assert.InRange(Y(pdf, 1, "AfterColumns") - Y(sequential, 1, "AfterColumns"), 149.9, 150.1);
        AssertMarkers(pdf, "Left", 12); AssertMarkers(pdf, "Right", 12);
    }

    [Fact]
    public void Columns_RowBalancingPreservesAFittingKeptCellParagraph() {
        var first = new[] { PdfTextRun.Normal(string.Join("\n", Enumerable.Range(1, 10).Select(index => $"Kept{index:D3}"))) };
        var second = new[] { PdfTextRun.Normal("Following001\nFollowing002\nFollowing003") };
        var paragraphs = new[] {
            new PdfTableCellParagraph(first, fontSize: 12, lineSpacing: PdfLineSpacing.Exactly(20), widowControl: false, keepTogether: true),
            new PdfTableCellParagraph(second, fontSize: 12, lineSpacing: PdfLineSpacing.Exactly(20), widowControl: false)
        };
        var rows = new[] { new[] { new PdfTableCell(first.Concat(second), paragraphs) } };
        using var pdf = PdfPigDocument.Open(Render(rows, TableStyle()));
        Assert.Equal(1, pdf.NumberOfPages);
        Assert.InRange(X(pdf, 1, "Kept010"), 39.9, 40.1);
        Assert.InRange(X(pdf, 1, "Following001"), 259.9, 260.1);
        AssertMarkers(pdf, "Kept", 10); AssertMarkers(pdf, "Following", 3);
    }

    private static PdfTableCell Cell(string prefix, int lines) {
        var runs = new[] { PdfTextRun.Normal(string.Join("\n", Enumerable.Range(1, lines).Select(index => $"{prefix}{index:D3}"))) };
        return new PdfTableCell(runs, new[] { new PdfTableCellParagraph(runs, fontSize: 12,
            lineSpacing: PdfLineSpacing.Exactly(20), widowControl: false) });
    }

    [Theory]
    [InlineData(true, 7, 7, 8)]
    [InlineData(false, 7, 5, 7)]
    public void Columns_RowBalancingHonorsCellParagraphKeepNextAndWidowBoundaries(
        bool keepNext, int firstCount, int secondCount, int leftLast) {
        var first = new[] { PdfTextRun.Normal(string.Join("\n", Enumerable.Range(1, firstCount).Select(index => $"Boundary{index:D3}"))) };
        var second = new[] { PdfTextRun.Normal(string.Join("\n", Enumerable.Range(firstCount + 1, secondCount).Select(index => $"Boundary{index:D3}"))) };
        var paragraphs = new[] {
            new PdfTableCellParagraph(first, fontSize: 12, lineSpacing: PdfLineSpacing.Exactly(20),
                widowControl: !keepNext, keepWithNext: keepNext),
            new PdfTableCellParagraph(second, fontSize: 12, lineSpacing: PdfLineSpacing.Exactly(20), widowControl: false)
        };
        using var pdf = PdfPigDocument.Open(Render(new[] { new[] { new PdfTableCell(first.Concat(second), paragraphs) } }, TableStyle()));
        Assert.Equal(1, pdf.NumberOfPages);
        Assert.InRange(X(pdf, 1, $"Boundary{leftLast:D3}"), 39.9, 40.1);
        Assert.InRange(X(pdf, 1, $"Boundary{leftLast + 1:D3}"), 259.9, 260.1);
        AssertMarkers(pdf, "Boundary", firstCount + secondCount);
    }

    [Fact]
    public void Columns_RowBalancingKeepsAFittingTableWithItsFollowingBlock() {
        var tableStyle = TableStyle(); tableStyle.KeepWithNext = true;
        var paragraphStyle = ParagraphStyle(); paragraphStyle.KeepTogether = true;
        byte[] bytes = PdfDocument.Create(Options()).Columns(content => {
            content.Table(new[] { new[] { Cell("KeptTable", 3) } }, style: tableStyle);
            content.Paragraph(p => p.Text(string.Join("\n", Enumerable.Range(1, 5).Select(index => $"KeptAfter{index:D3}"))), style: paragraphStyle);
        }, new PdfMultiColumnOptions { Gap = 20, BalanceTableRowLines = true }).ToBytes();
        using var pdf = PdfPigDocument.Open(bytes);
        Assert.Equal(1, pdf.NumberOfPages);
        Assert.InRange(X(pdf, 1, "KeptTable003"), 39.9, 40.1);
        Assert.InRange(X(pdf, 1, "KeptAfter005"), 39.9, 40.1);
        AssertMarkers(pdf, "KeptTable", 3); AssertMarkers(pdf, "KeptAfter", 5);
    }

    [Theory]
    [InlineData(7, false)]
    [InlineData(13, true)]
    public void Columns_RowBalancingPreservesFittingBodyGroupsAndRelaxesOversizedGroups(int lines, bool secondColumn) {
        var style = TableStyle();
        style.MinimumBodyRowsOnFirstPage = 2; style.MinimumBodyRowsOnLastPage = 2;
        using var pdf = PdfPigDocument.Open(Render(new[] { new[] { Cell("FirstRow", lines) }, new[] { Cell("SecondRow", lines) } }, style));
        Assert.Equal(1, pdf.NumberOfPages);
        double secondX = secondColumn ? 260 : 40;
        Assert.InRange(X(pdf, 1, "SecondRow001"), secondX - .1, secondX + .1);
        Assert.InRange(X(pdf, 1, $"SecondRow{lines:D3}"), secondX - .1, secondX + .1);
        AssertMarkers(pdf, "FirstRow", lines); AssertMarkers(pdf, "SecondRow", lines);
    }

    [Fact]
    public void Columns_RowBalancingReservesTableSpacingAfterTheFinalFragment() {
        var style = TableStyle(); style.SpacingAfter = 40;
        using var pdf = PdfPigDocument.Open(Render(new[] { new[] { Cell("Spaced", 21) } }, style));
        Assert.Equal(1, pdf.NumberOfPages);
        Assert.InRange(X(pdf, 1, "Spaced012"), 39.9, 40.1);
        Assert.InRange(X(pdf, 1, "Spaced013"), 259.9, 260.1);
        // The left fragment is 240pt; the right fragment plus its 40pt table spacing is 220pt.
        Assert.InRange(Y(pdf, 1, "Spaced013") - Y(pdf, 1, "AfterColumns"), 239.9, 240.1);
        AssertMarkers(pdf, "Spaced", 21);
    }

    private static byte[] Render(PdfTableCell[][] rows, PdfTableStyle style, bool balance = true) =>
        PdfDocument.Create(Options()).Columns(content => content.Table(rows, style: style),
            new PdfMultiColumnOptions { Gap = 20, BalanceTableRowLines = balance })
            .Paragraph(p => p.Text("AfterColumns"), style: ParagraphStyle()).ToBytes();
    private static double X(PdfPigDocument pdf, int page, string marker) =>
        pdf.GetPage(page).GetWords().Single(word => word.Text == marker).BoundingBox.Left;
    private static double Y(PdfPigDocument pdf, int page, string marker) =>
        pdf.GetPage(page).GetWords().Single(word => word.Text == marker).BoundingBox.Bottom;
    private static void AssertMarkers(PdfPigDocument pdf, string prefix, int count) =>
        Assert.Equal(Enumerable.Range(1, count), System.Text.RegularExpressions.Regex.Matches(
            string.Concat(Enumerable.Range(1, pdf.NumberOfPages).Select(page => pdf.GetPage(page).Text)), prefix + @"(\d{3})")
            .Cast<System.Text.RegularExpressions.Match>().Select(match => int.Parse(match.Groups[1].Value)));
    private static PdfOptions Options() => new() { PageWidth = 500, PageHeight = 400,
        MarginLeft = 40, MarginRight = 40, MarginTop = 40, MarginBottom = 40, DefaultFontSize = 12 };
    private static PdfParagraphStyle ParagraphStyle() => new() { LineSpacing = PdfLineSpacing.Exactly(20),
        SpacingBefore = 0, SpacingAfter = 0, WidowControl = false };
    private static PdfTableStyle TableStyle() => new() { HeaderRowCount = 0, CellPaddingX = 0, CellPaddingY = 0,
        FontSize = 12, LineHeight = 20D / 12D, BorderWidth = 0, RowSeparatorWidth = 0, SpacingBefore = 0, SpacingAfter = 0,
        MinimumBodyRowsOnFirstPage = 0, MinimumBodyRowsOnLastPage = 0 };
}
