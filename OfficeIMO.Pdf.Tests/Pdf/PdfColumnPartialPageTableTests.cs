using OfficeIMO.Pdf;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfColumnPartialPageTableTests {
    [Theory]
    [InlineData(13, false)]
    [InlineData(13, true)]
    [InlineData(21, false)]
    [InlineData(43, false)]
    public void Columns_ContinuingTableUsesTheAvailableFullPhysicalPageForCellKeepRules(int lines, bool keepTogether) {
        using var pdf = PdfPigDocument.Open(Render(lines, keepTogether, TableStyle()));
        int finalPage = lines == 43 ? 3 : 2;
        int leftLast = keepTogether ? 13 : lines == 43 ? 42 : lines == 21 ? 15 : 11;
        Assert.Equal(finalPage, pdf.NumberOfPages);
        if (keepTogether) Assert.DoesNotContain(pdf.GetPage(1).GetWords(), word => word.Text.StartsWith("Cell", StringComparison.Ordinal));
        else Assert.InRange(X(pdf, 1, "Cell008"), 259.9, 260.1);
        Assert.InRange(X(pdf, finalPage, $"Cell{leftLast:D3}"), 39.9, 40.1);
        if (!keepTogether) Assert.InRange(X(pdf, finalPage, $"Cell{leftLast + 1:D3}"), 259.9, 260.1);
        Assert.InRange(X(pdf, finalPage, "AfterColumns"), 39.9, 40.1);
        AssertMarkers(pdf, lines);
    }

    [Theory]
    [InlineData("minimum-row")]
    [InlineData("fixed-row")]
    [InlineData("caption")]
    [InlineData("kept-table")]
    [InlineData("unsplittable-row")]
    public void Columns_TableThatFitsAFullPageMovesPastPartialFramesInsteadOfThrowing(string constraint) {
        var style = TableStyle();
        if (constraint == "minimum-row") style.RowMinHeights = new List<double?> { 260 };
        if (constraint == "fixed-row") style.FixedRowHeights = new List<double?> { 260 };
        if (constraint == "caption") {
            style.RowMinHeights = new List<double?> { 260 };
            style.Caption = "TableCaption"; style.CaptionFontSize = 12; style.CaptionSpacingAfter = 0;
        }
        if (constraint == "kept-table") style.KeepTogether = true;
        if (constraint == "unsplittable-row") style.RowAllowBreakAcrossPages = new List<bool?> { false };
        int lines = constraint is "minimum-row" or "fixed-row" or "caption" ? 1 : 13;
        using var pdf = PdfPigDocument.Open(Render(lines, false, style));
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.DoesNotContain(pdf.GetPage(1).GetWords(), word => word.Text.StartsWith("Cell", StringComparison.Ordinal));
        Assert.InRange(X(pdf, 2, "Cell001"), 39.9, 40.1);
        Assert.InRange(X(pdf, 2, $"Cell{lines:D3}"), 39.9, 40.1);
        Assert.InRange(X(pdf, 2, "AfterColumns"), 39.9, 40.1);
        AssertMarkers(pdf, lines);
        if (constraint == "caption") Assert.Single(pdf.GetPage(2).GetWords(), word => word.Text == "TableCaption");
    }

    [Fact]
    public void Columns_VisibleMergedCellSpanMovesTogetherToTheFullPhysicalPage() {
        var style = TableStyle();
        style.FixedRowHeights = new List<double?> { 130, 130 };
        var viewport = new PdfTableCell("MergedCell", rowSpan: 2)
            .WithViewport(new PdfTableCellViewport(100, 260, 100, 260));
        using var pdf = PdfPigDocument.Open(RenderTable(new[] {
            new[] { viewport, new PdfTableCell("FirstRow") }, new[] { new PdfTableCell("SecondRow") }
        }, style));
        Assert.Equal(2, pdf.NumberOfPages);
        foreach (string marker in new[] { "MergedCell", "FirstRow", "SecondRow" }) {
            Assert.DoesNotContain(pdf.GetPage(1).GetWords(), word => word.Text == marker);
            Assert.Single(pdf.GetPage(2).GetWords(), word => word.Text == marker);
        }
        Assert.InRange(X(pdf, 2, "MergedCell"), 39.9, 40.1);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Columns_TableCapacityIncludesPaddingOutsideTheColumnsAndRejectsAnImpossibleRetry(bool keepTable) {
        var style = TableStyle();
        style.RowMinHeights = new List<double?> { 310 };
        style.KeepTogether = keepTable;
        var options = new PdfOptions { PageWidth = 500, PageHeight = 400, MarginLeft = 40, MarginRight = 40,
            MarginTop = 40, MarginBottom = 40, DefaultFontSize = 12, MaxGeneratedPages = 3 };
        var document = PdfDocument.Create(options).Container(content => {
            content.Paragraph(p => p.Text("Prelude"), style: ParagraphStyle());
            content.Columns(columns => columns.Table(new[] { new[] { new PdfTableCell("TooTall") } }, style: style),
                new PdfMultiColumnOptions { Gap = 20 });
        }, new PdfPanelStyle { PaddingY = 20, PaddingX = 0 });
        ArgumentException error = Assert.Throws<ArgumentException>(() => document.ToBytes());
        Assert.Contains("height", error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Columns_FittingHeaderAndMinimumBodyGroupMovesBeforePaintingTheHeader() {
        var style = TableStyle();
        style.HeaderRowCount = 1; style.MinimumBodyRowsOnFirstPage = 1;
        style.RowMinHeights = new List<double?> { 20, 260 };
        using var pdf = PdfPigDocument.Open(RenderTable(new[] {
            new[] { new PdfTableCell("Header") }, new[] { new PdfTableCell("Body") }
        }, style));
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.DoesNotContain(pdf.GetPage(1).GetWords(), word => word.Text == "Header");
        Assert.DoesNotContain(pdf.GetPage(1).GetWords(), word => word.Text == "Body");
        Assert.Single(pdf.GetPage(2).GetWords(), word => word.Text == "Header");
        Assert.Single(pdf.GetPage(2).GetWords(), word => word.Text == "Body");
    }

    [Fact]
    public void Columns_TableThatFitsThePaddedPhysicalPageStillAdvancesAndRendersOnce() {
        var style = TableStyle(); style.RowMinHeights = new List<double?> { 290 }; style.KeepTogether = true;
        using var pdf = PdfPigDocument.Open(PdfDocument.Create(new PdfOptions { PageWidth = 500, PageHeight = 400,
            MarginLeft = 40, MarginRight = 40, MarginTop = 40, MarginBottom = 40, DefaultFontSize = 12, MaxGeneratedPages = 3 })
            .Container(outer => outer.Panel(inner => {
                inner.Paragraph(p => p.Text("Prelude"), style: ParagraphStyle());
                inner.Columns(columns => columns.Table(new[] { new[] { new PdfTableCell("FittingCell") } }, style: style),
                    new PdfMultiColumnOptions { Gap = 20 });
            }, new PdfPanelStyle { PaddingY = 10, PaddingX = 0 }), new PdfPanelStyle { PaddingY = 20, PaddingX = 0 }).ToBytes());
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.DoesNotContain(pdf.GetPage(1).GetWords(), word => word.Text == "FittingCell");
        Assert.Single(pdf.GetPage(2).GetWords(), word => word.Text == "FittingCell");
        Assert.InRange(X(pdf, 2, "FittingCell"), 39.9, 40.1);
        Assert.All(pdf.GetPage(2).Letters, letter => Assert.True(letter.GlyphRectangle.Bottom >= 40));
    }

    private static byte[] Render(int lines, bool keepTogether, PdfTableStyle style) {
        var runs = new[] { PdfTextRun.Normal(string.Join("\n", Enumerable.Range(1, lines).Select(index => $"Cell{index:D3}"))) };
        var cell = new PdfTableCell(runs, new[] { new PdfTableCellParagraph(runs, fontSize: 12,
            lineSpacing: PdfLineSpacing.Exactly(20), widowControl: false, keepTogether: keepTogether) });
        return RenderTable(new[] { new[] { cell } }, style);
    }
    private static byte[] RenderTable(PdfTableCell[][] rows, PdfTableStyle style) =>
        PdfDocument.Create(new PdfOptions { PageWidth=500, PageHeight=400, MarginLeft=40, MarginRight=40,
            MarginTop=40, MarginBottom=40, DefaultFontSize=12 })
            .Paragraph(p => p.Text(string.Join("\n", Enumerable.Range(1, 12).Select(index => $"Before{index:D3}"))), style: ParagraphStyle())
            .Columns(content => content.Table(rows, style: style),
                new PdfMultiColumnOptions { Gap=20, BalanceTableRowLines=true })
            .Paragraph(p => p.Text("AfterColumns"), style: ParagraphStyle()).ToBytes();
    private static double X(PdfPigDocument pdf, int page, string marker) =>
        pdf.GetPage(page).GetWords().Single(word => word.Text == marker).BoundingBox.Left;
    private static void AssertMarkers(PdfPigDocument pdf, int count) {
        var words = Enumerable.Range(1, pdf.NumberOfPages).SelectMany(page => pdf.GetPage(page).GetWords()).ToArray();
        foreach (int index in Enumerable.Range(1, count)) Assert.Single(words, word => word.Text == $"Cell{index:D3}");
        foreach (int index in Enumerable.Range(1, 12)) Assert.Single(words, word => word.Text == $"Before{index:D3}");
        Assert.Single(words, word => word.Text == "AfterColumns");
    }
    private static PdfParagraphStyle ParagraphStyle() => new() { LineSpacing=PdfLineSpacing.Exactly(20),
        SpacingBefore=0, SpacingAfter=0, WidowControl=false };
    private static PdfTableStyle TableStyle() => new() { HeaderRowCount=0, CellPaddingX=0, CellPaddingY=0,
        FontSize=12, LineHeight=20D/12D, BorderWidth=0, RowSeparatorWidth=0, SpacingBefore=0, SpacingAfter=0,
        MinimumBodyRowsOnFirstPage=0, MinimumBodyRowsOnLastPage=0 };
}
