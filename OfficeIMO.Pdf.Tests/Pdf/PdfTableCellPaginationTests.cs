using OfficeIMO.Pdf;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfTableCellPaginationTests {
    [Theory]
    [InlineData("flow", false, true)]
    [InlineData("flow", true, false)]
    [InlineData("flow", null, false)]
    [InlineData("columns", false, true)]
    [InlineData("columns", true, false)]
    [InlineData("row", false, true)]
    [InlineData("row", true, false)]
    public void ResolvedCellWidowPolicyControlsOneLineRowFragments(string mode, bool? widowControl, bool firstLineFits) {
        var cells = Enumerable.Range(1, 8).Select(row => new[] {
            Cell(string.Join("\n", Enumerable.Range(1, 3).Select(line => $"Row{row:D2}Line{line:D3}")), widowControl)
        }).ToArray();
        using var pdf = PdfPigDocument.Open(Render(mode, cells));
        if (mode == "columns") {
            Assert.Equal(1, pdf.NumberOfPages);
            Assert.InRange(MarkerX(pdf, 1, "Row06Line001"), firstLineFits ? 39.9 : 259.9, firstLineFits ? 40.1 : 260.1);
        } else {
            Assert.Equal(2, pdf.NumberOfPages);
            Assert.Equal(firstLineFits, pdf.GetPage(1).Text.Contains("Row06Line001"));
        }
        Assert.Contains("Row08Line003", pdf.GetPage(pdf.NumberOfPages).Text);
    }

    [Theory]
    [InlineData("flow")]
    [InlineData("columns")]
    [InlineData("row")]
    public void CellKeepTogetherMovesAFittingParagraphToTheNextFrame(string mode) {
        var cells = new[] {
            new[] { Cell(string.Join("\n", Enumerable.Range(1, 14).Select(line => "Prelude" + line)), false) },
            new[] { Cell("KeptFirst\nKeptSecond\nKeptThird\nKeptLast", false, keepTogether: true) }
        };
        using var pdf = PdfPigDocument.Open(Render(mode, cells));
        if (mode == "columns") Assert.InRange(MarkerX(pdf, 1, "KeptFirst"), 259.9, 260.1);
        else Assert.DoesNotContain("KeptFirst", pdf.GetPage(1).Text);
        Assert.Contains("KeptLast", pdf.GetPage(pdf.NumberOfPages).Text);
    }

    private static PdfTableCell Cell(string text, bool? widowControl, bool keepTogether = false) {
        var runs = new[] { PdfTextRun.Normal(text) };
        return new PdfTableCell(runs, new[] { new PdfTableCellParagraph(runs, fontSize: 12,
            lineSpacing: PdfLineSpacing.Exactly(20), widowControl: widowControl, keepTogether: keepTogether) });
    }

    [Theory]
    [InlineData("flow")]
    [InlineData("columns")]
    [InlineData("row")]
    public void CellWidowControlDoesNotLeaveOneFinalLineInTheNextFrame(string mode) {
        var cells = new[] {
            new[] { Cell(string.Join("\n", Enumerable.Range(1, 12).Select(line => "Prelude" + line)), false) },
            new[] { Cell("BodyOne\nBodyTwo\nBodyThree\nBodyFour\nBodyFive", true) }
        };
        using var pdf = PdfPigDocument.Open(Render(mode, cells));
        if (mode == "columns") Assert.InRange(MarkerX(pdf, 1, "BodyFour"), 259.9, 260.1);
        else Assert.DoesNotContain("BodyFour", pdf.GetPage(1).Text);
        Assert.Contains("BodyFive", pdf.GetPage(pdf.NumberOfPages).Text);
    }

    [Theory]
    [InlineData("flow", true)]
    [InlineData("flow", false)]
    [InlineData("columns", true)]
    [InlineData("columns", false)]
    [InlineData("row", true)]
    [InlineData("row", false)]
    public void CellKeepWithNextReservesTheNextCellParagraph(string mode, bool keepWithNext) {
        var first = new[] { PdfTextRun.Normal("JoinedFirst") };
        var next = new[] { PdfTextRun.Normal("NextOne\nNextTwo\nNextThree") };
        var cell = new PdfTableCell(first.Concat(next), new[] {
            new PdfTableCellParagraph(first, fontSize: 12, lineSpacing: PdfLineSpacing.Exactly(20), widowControl: false, keepWithNext: keepWithNext),
            new PdfTableCellParagraph(next, fontSize: 12, lineSpacing: PdfLineSpacing.Exactly(20), widowControl: false)
        });
        var cells = new[] {
            new[] { Cell(string.Join("\n", Enumerable.Range(1, 15).Select(line => "Prelude" + line)), false) }, new[] { cell }
        };
        using var pdf = PdfPigDocument.Open(Render(mode, cells));
        if (mode == "columns") Assert.InRange(MarkerX(pdf, 1, "JoinedFirst"), keepWithNext ? 259.9 : 39.9, keepWithNext ? 260.1 : 40.1);
        else Assert.Equal(!keepWithNext, pdf.GetPage(1).Text.Contains("JoinedFirst"));
        Assert.Contains("NextThree", pdf.GetPage(pdf.NumberOfPages).Text);
    }

    private static byte[] Render(string mode, PdfTableCell[][] cells) {
        var options = new PdfOptions { PageWidth = 500, PageHeight = 400, MarginLeft = 40, MarginRight = 40,
            MarginTop = 40, MarginBottom = 40, DefaultFontSize = 12 };
        var style = new PdfTableStyle { HeaderRowCount = 0, CellPaddingX = 0, CellPaddingY = 0,
            FontSize = 12, LineHeight = 20D / 12D, SpacingBefore = 0, SpacingAfter = 0,
            BorderWidth = 0, RowSeparatorWidth = 0 };
        return mode switch {
            "flow" => PdfDocument.Create(options).Table(cells, style: style).ToBytes(),
            "columns" => PdfDocument.Create(options).Columns(content => content.Table(cells, style: style),
                new PdfMultiColumnOptions { Gap = 20, BalanceLastPage = false }).ToBytes(),
            "row" => PdfDocument.Create(options).Row(row => row.PercentColumn(100, column => column.Table(cells, style: style))).ToBytes(),
            _ => throw new ArgumentException(nameof(mode))
        };
    }

    private static double MarkerX(PdfPigDocument pdf, int page, string marker) {
        var letters = pdf.GetPage(page).Letters;
        int index = string.Concat(letters.Select(letter => letter.Value)).IndexOf(marker, StringComparison.Ordinal);
        Assert.True(index >= 0, "Missing marker " + marker);
        return letters[index].StartBaseLine.X;
    }
}
