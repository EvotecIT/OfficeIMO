using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfSequentialColumnFlowTests {
    [Fact]
    public void Columns_SnapshotsExplicitDefinitionsAndRejectsAConflictingCountAtomically() {
        var definitions = new[] {
            new PdfFlowColumn(PdfColumnWidth.Fixed(100), 20), new PdfFlowColumn(PdfColumnWidth.Fixed(300))
        };
        var options = new PdfMultiColumnOptions { ColumnDefinitions = definitions, BalanceLastPage = false };
        definitions[0] = new PdfFlowColumn(PdfColumnWidth.Fixed(300), 20);
        definitions[1] = new PdfFlowColumn(PdfColumnWidth.Fixed(100));
        Assert.Throws<InvalidOperationException>(() => options.ColumnCount = 3);
        Assert.Equal(2, options.ColumnCount);
        byte[] bytes = PdfDocument.Create(Options()).Columns(content => {
            content.Paragraph(p => p.Text("First"), style: Paragraph());
            content.ColumnBreak();
            content.Paragraph(p => p.Text("Second"), style: Paragraph());
        }, options).ToBytes();
        using var read = UglyToad.PdfPig.PdfDocument.Open(bytes);
        Assert.Equal(1, read.NumberOfPages);
        Assert.InRange(X(read, 1, "Second"), 159.9, 160.1);
    }

    [Fact]
    public void Columns_BalancesTheLastPageOfAContinuingParagraph() {
        byte[] bytes = PdfDocument.Create(Options()).Columns(content => content.Paragraph(p =>
            p.Text(string.Join("\n", Enumerable.Range(1, 42).Select(index => "Line" + index.ToString("D2")))), style: Paragraph()),
            new PdfMultiColumnOptions { Gap = 20, BalanceLastPage = true })
            .Paragraph(p => p.Text("AfterColumns"), style: Paragraph()).ToBytes();
        using var read = UglyToad.PdfPig.PdfDocument.Open(bytes);
        Assert.Equal(2, read.NumberOfPages);
        Assert.InRange(X(read, 2, "Line37"), 39.9, 40.1);
        Assert.InRange(X(read, 2, "Line38"), 259.9, 260.1);
        Assert.InRange(X(read, 2, "AfterColumns"), 39.9, 40.1);
        Assert.InRange(Y(read, 2, "Line33") - Y(read, 2, "AfterColumns"), 99.9, 100.1);
    }

    [Fact]
    public void Columns_UnequalContinuationPreservesParagraphTabs() {
        var style = Paragraph();
        style.TabStops.Add(new PdfTabStop(50));
        string text = string.Join("\n", Enumerable.Range(1, 16).Select(index => "Prelude" + index)) + "\nLeft\tRight";
        byte[] bytes = PdfDocument.Create(Options()).Columns(content => content.Paragraph(p => p.Text(text), style: style), Unequal()).ToBytes();
        using var read = UglyToad.PdfPig.PdfDocument.Open(bytes);
        Assert.Equal(1, read.NumberOfPages);
        Assert.InRange(X(read, 1, "Right"), 209.9, 210.1);
    }

    [Fact]
    public void Columns_UnequalTableContinuationPreservesParagraphIndentsAndSpacing() {
        string text = string.Join(" ", Enumerable.Range(1, 80).Select(index => "m" + index.ToString("D3")));
        var paragraphs = new[] {
            new PdfTableCellParagraph(new[] { PdfTextRun.Normal(text) }, spacingAfter: 6, leftIndent: 6, rightIndent: 4,
                firstLineIndent: 10, fontSize: 12, lineSpacing: PdfLineSpacing.Exactly(20)),
            new PdfTableCellParagraph(new[] { PdfTextRun.Normal("SecondA\nSecondB") }, spacingBefore: 10, leftIndent: 6,
                firstLineIndent: 20, fontSize: 12, lineSpacing: PdfLineSpacing.Exactly(20))
        };
        var cell = new PdfTableCell(new[] { PdfTextRun.Normal(text + "\nSecondA\nSecondB") }, paragraphs);
        byte[] bytes = PdfDocument.Create(Options()).Columns(content => content.Table(new[] { new[] { cell } },
            style: new PdfTableStyle { HeaderRowCount = 0, CellPaddingX = 0, CellPaddingY = 0, FontSize = 12,
                LineHeight = 20D / 12D, BorderWidth = 0, SpacingBefore = 0, SpacingAfter = 0 }), Unequal()).ToBytes();
        using var read = UglyToad.PdfPig.PdfDocument.Open(bytes);
        Assert.Equal(1, read.NumberOfPages);
        int firstContinued = Enumerable.Range(1, 80).First(index => X(read, 1, "m" + index.ToString("D3")) > 150);
        Assert.InRange(X(read, 1, "m" + firstContinued.ToString("D3")), 165.9, 166.1);
        Assert.InRange(X(read, 1, "SecondA"), 185.9, 186.1);
        Assert.InRange(X(read, 1, "SecondB"), 165.9, 166.1);
        Assert.InRange(Y(read, 1, "m080") - Y(read, 1, "SecondA"), 35.9, 36.1);
    }

    [Fact]
    public void Columns_BalancesOddParagraphLineCountsWithoutAnExtraPage() {
        byte[] bytes = PdfDocument.Create(Options()).Columns(content => content.Paragraph(p =>
            p.Text("Line1\nLine2\nLine3\nLine4\nLine5"), style: Paragraph()),
            new PdfMultiColumnOptions { Gap = 20, BalanceLastPage = true }).ToBytes();
        using var read = UglyToad.PdfPig.PdfDocument.Open(bytes);
        Assert.Equal(1, read.NumberOfPages);
        Assert.InRange(X(read, 1, "Line3"), 39.9, 40.1);
        Assert.InRange(X(read, 1, "Line4"), 259.9, 260.1);
    }

    [Theory]
    [InlineData("heading", 100, 1)]
    [InlineData("list", 100, 1)]
    [InlineData("table", 100, 1)]
    [InlineData("heading", 250, 2)]
    [InlineData("list", 250, 2)]
    [InlineData("table", 250, 2)]
    public void Columns_UnequalFramesRewrapContinuingHeadingListAndTableText(string kind, int words, int pages) {
        string text = string.Join(" ", Enumerable.Range(1, words).Select(index => "m" + index.ToString("D3")));
        byte[] bytes = PdfDocument.Create(Options()).Columns(content => {
            if (kind == "heading") content.H1(text, style: new PdfHeadingStyle {
                FontSize = 12, LineSpacing = PdfLineSpacing.Exactly(20), SpacingBefore = 0, SpacingAfter = 0, KeepWithNext = false
            });
            if (kind == "list") content.RichNumbered(new[] { new PdfListItem(text) }, style: new PdfListStyle {
                LineSpacing = PdfLineSpacing.Exactly(20), ItemSpacing = 0, SpacingAfter = 0
            });
            if (kind == "table") content.Table(new[] { new[] { new PdfTableCell(text) } }, style: new PdfTableStyle {
                HeaderRowCount = 0, CellPaddingX = 0, CellPaddingY = 0, FontSize = 12, LineHeight = 20D / 12D,
                BorderWidth = 0, SpacingBefore = 0, SpacingAfter = 0
            });
        }, new PdfMultiColumnOptions {
            BalanceLastPage = false,
            ColumnDefinitions = new[] {
                new PdfFlowColumn(PdfColumnWidth.Fixed(100), 20), new PdfFlowColumn(PdfColumnWidth.Fixed(300))
            }
        }).ToBytes();
        using var read = UglyToad.PdfPig.PdfDocument.Open(bytes);
        Assert.Equal(pages, read.NumberOfPages);
        if (pages == 1) Assert.InRange(X(read, 1, "m100"), 160, 460);
        string rendered = string.Join(" ", Enumerable.Range(1, pages).Select(page => read.GetPage(page).Text));
        int[] markers = System.Text.RegularExpressions.Regex.Matches(rendered, @"m(\d{3})")
            .Cast<System.Text.RegularExpressions.Match>().Select(match => int.Parse(match.Groups[1].Value)).ToArray();
        Assert.Equal(Enumerable.Range(1, words), markers);
    }

    [Fact]
    public void Columns_RewrapsTheRemainingParagraphInAnUnequalWidthFrame() {
        byte[] bytes = PdfDocument.Create(Options()).Columns(content =>
            content.Paragraph(p => p.Text(string.Join(" ", Enumerable.Range(1, 100).Select(index => "m" + index.ToString("D3")))), style: Paragraph()),
            new PdfMultiColumnOptions {
                BalanceLastPage = false,
                ColumnDefinitions = new[] {
                    new PdfFlowColumn(PdfColumnWidth.Fixed(100), 20), new PdfFlowColumn(PdfColumnWidth.Fixed(300))
                }
            }).ToBytes();
        using var read = UglyToad.PdfPig.PdfDocument.Open(bytes);
        Assert.Equal(1, read.NumberOfPages);
        Assert.InRange(X(read, 1, "m001"), 39.9, 40.1);
        Assert.InRange(X(read, 1, "m100"), 160, 460);
        foreach (int index in Enumerable.Range(1, 100)) Assert.Contains("m" + index.ToString("D3"), read.GetPage(1).Text);
    }

    [Fact]
    public void Columns_ListContinuesInTheNextColumnBeforeStartingAnotherPage() {
        byte[] bytes = PdfDocument.Create(Options()).Columns(content =>
            content.RichNumbered(Enumerable.Range(1, 24).Select(index => new PdfListItem("List" + index.ToString("D2"))),
                style: new PdfListStyle { LineSpacing = PdfLineSpacing.Exactly(20), ItemSpacing = 0, SpacingAfter = 0 }),
            new PdfMultiColumnOptions { Gap = 20, BalanceLastPage = false }).ToBytes();
        using var read = UglyToad.PdfPig.PdfDocument.Open(bytes);
        Assert.Equal(1, read.NumberOfPages);
        Assert.InRange(X(read, 1, "List01"), 40, 80);
        Assert.InRange(X(read, 1, "List17"), 260, 300);
        Assert.Contains("List24", read.GetPage(1).Text);
    }

    [Fact]
    public void Columns_KeepWithNextReservesTheNextParagraphsFirstLine() {
        byte[] bytes = PdfDocument.Create(Options()).Columns(content => {
            content.Paragraph(p => p.Text(string.Join("\n", Enumerable.Range(1, 15).Select(index => "Prelude" + index))), style: Paragraph());
            var heading = Paragraph(); heading.KeepWithNext = true;
            content.Paragraph(p => p.Text("KeptHeading"), style: heading);
            content.Paragraph(p => p.Text("BodyFirst\nBodySecond\nBodyThird"), style: Paragraph());
        }, new PdfMultiColumnOptions { Gap = 20, BalanceLastPage = false }).ToBytes();
        using var read = UglyToad.PdfPig.PdfDocument.Open(bytes);
        Assert.Equal(1, read.NumberOfPages);
        Assert.InRange(X(read, 1, "KeptHeading"), 259.9, 260.1);
        Assert.InRange(X(read, 1, "BodyFirst"), 259.9, 260.1);
    }

    [Fact]
    public void Columns_TableContinuesInTheNextColumnBeforeStartingAnotherPage() {
        byte[] bytes = PdfDocument.Create(Options()).Columns(content =>
            content.Table(Enumerable.Range(1, 8).Select(row => new[] {
                new PdfTableCell(string.Join("\n", Enumerable.Range(1, 3).Select(line => "Row" + row + "Line" + line)))
            }), style: new PdfTableStyle {
                HeaderRowCount = 0, CellPaddingX = 0, CellPaddingY = 0,
                FontSize = 12, LineHeight = 20D / 12D,
                SpacingBefore = 0, SpacingAfter = 0, BorderWidth = 0, RowSeparatorWidth = 0
            }), new PdfMultiColumnOptions { Gap = 20, BalanceLastPage = false }).ToBytes();
        using var read = UglyToad.PdfPig.PdfDocument.Open(bytes);
        Assert.Equal(1, read.NumberOfPages);
        Assert.InRange(X(read, 1, "Row1Line1"), 39.9, 40.1);
        Assert.InRange(X(read, 1, "Row8Line3"), 259.9, 260.1);
    }

    private static PdfOptions Options() => new() {
        PageWidth = 500, PageHeight = 400, MarginLeft = 40, MarginRight = 40,
        MarginTop = 40, MarginBottom = 40, DefaultFontSize = 12, CompressContentStreams = false
    };

    private static PdfMultiColumnOptions Unequal() => new() {
        BalanceLastPage = false,
        ColumnDefinitions = new[] {
            new PdfFlowColumn(PdfColumnWidth.Fixed(100), 20), new PdfFlowColumn(PdfColumnWidth.Fixed(300))
        }
    };

    private static PdfParagraphStyle Paragraph() => new() {
        LineSpacing = PdfLineSpacing.Exactly(20), SpacingBefore = 0, SpacingAfter = 0, WidowControl = false
    };

    private static double X(UglyToad.PdfPig.PdfDocument read, int page, string marker) {
        var letters = read.GetPage(page).Letters;
        string text = string.Concat(letters.Select(letter => letter.Value));
        int index = text.IndexOf(marker, StringComparison.Ordinal);
        Assert.True(index >= 0, "Missing marker " + marker + " in " + text);
        return letters[index].StartBaseLine.X;
    }

    private static double Y(UglyToad.PdfPig.PdfDocument read, int page, string marker) {
        var letters = read.GetPage(page).Letters;
        string text = string.Concat(letters.Select(letter => letter.Value));
        int index = text.IndexOf(marker, StringComparison.Ordinal);
        Assert.True(index >= 0, "Missing marker " + marker);
        return letters[index].StartBaseLine.Y;
    }
}
