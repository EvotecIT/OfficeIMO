using OfficeIMO.Pdf;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfColumnTerminalSpacingTests {
    [Theory]
    [InlineData(0, 40)]
    [InlineData(20, 60)]
    public void Columns_TerminalSpacingRetainsBalancedBreakpointsAndIsSnapshotted(double spacing, double followingOffset) {
        var options = new PdfMultiColumnOptions { Gap = global::OfficeIMO.OfficeLength.Points(20), FinalColumnSpacingAfter = global::OfficeIMO.OfficeLength.Points(spacing) };
        var document = Render(options);
        options.FinalColumnSpacingAfter = global::OfficeIMO.OfficeLength.Points(100);
        using var pdf = PdfPigDocument.Open(document.ToBytes());
        Assert.Equal(1, pdf.NumberOfPages);
        Assert.InRange(Word(pdf, 1, "Column003").BoundingBox.Left, 259.9, 260.1);
        Assert.Equal(followingOffset, Word(pdf, 1, "Column001").BoundingBox.Top - Word(pdf, 1, "AfterColumns").BoundingBox.Top, 2);
    }

    [Fact]
    public void Columns_TerminalSpacingStopsAtThePhysicalPageMargin() {
        using var pdf = PdfPigDocument.Open(Render(new PdfMultiColumnOptions { Gap = global::OfficeIMO.OfficeLength.Points(20), FinalColumnSpacingAfter = global::OfficeIMO.OfficeLength.Points(1000) }).ToBytes());
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.InRange(Word(pdf, 1, "Column003").BoundingBox.Left, 259.9, 260.1);
        Assert.Equal(Word(pdf, 1, "Column001").BoundingBox.Top, Word(pdf, 2, "AfterColumns").BoundingBox.Top, 2);
    }

    [Fact]
    public void Columns_TerminalSpacingFollowsTheLastOccupiedColumnBeforeAnExplicitEmptyColumn() {
        var style = new PdfParagraphStyle { LineSpacing = PdfLineSpacing.Exact(global::OfficeIMO.OfficeLength.Points(20)), SpacingBefore = global::OfficeIMO.OfficeLength.Points(0), SpacingAfter = global::OfficeIMO.OfficeLength.Points(0), WidowControl = false };
        using var pdf = PdfPigDocument.Open(PdfDocument.Create(new PdfCreateOptions { PageSetup = {
            Size = new PdfPageSize(global::OfficeIMO.OfficeLength.Points(500), global::OfficeIMO.OfficeLength.Points(400)),
            Margins = { Left = global::OfficeIMO.OfficeLength.Points(40), Right = global::OfficeIMO.OfficeLength.Points(40), Top = global::OfficeIMO.OfficeLength.Points(40), Bottom = global::OfficeIMO.OfficeLength.Points(40) }
        }, DefaultFontSize = global::OfficeIMO.OfficeLength.Points(12) })
            .Columns(content => {
                content.AddParagraph(p => p.AddText("Column001\nColumn002"), style: style);
                content.ColumnBreak();
            }, new PdfMultiColumnOptions { Gap = global::OfficeIMO.OfficeLength.Points(20), BalanceLastPage = false, FinalColumnSpacingAfter = global::OfficeIMO.OfficeLength.Points(20) })
            .Paragraph(p => p.AddText("AfterColumns"), style: style).ToBytes());
        Assert.Equal(1, pdf.NumberOfPages);
        Assert.Equal(60D, Word(pdf, 1, "Column001").BoundingBox.Top - Word(pdf, 1, "AfterColumns").BoundingBox.Top, 2);
    }

    private static PdfDocument Render(PdfMultiColumnOptions options) {
        var style = new PdfParagraphStyle { LineSpacing = PdfLineSpacing.Exact(global::OfficeIMO.OfficeLength.Points(20)), SpacingBefore = global::OfficeIMO.OfficeLength.Points(0), SpacingAfter = global::OfficeIMO.OfficeLength.Points(0), WidowControl = false };
        return PdfDocument.Create(new PdfCreateOptions { PageSetup = {
            Size = new PdfPageSize(global::OfficeIMO.OfficeLength.Points(500), global::OfficeIMO.OfficeLength.Points(400)),
            Margins = { Left = global::OfficeIMO.OfficeLength.Points(40), Right = global::OfficeIMO.OfficeLength.Points(40), Top = global::OfficeIMO.OfficeLength.Points(40), Bottom = global::OfficeIMO.OfficeLength.Points(40) }
        }, DefaultFontSize = global::OfficeIMO.OfficeLength.Points(12) })
            .Columns(content => content.AddParagraph(p => p.AddText("Column001\nColumn002\nColumn003\nColumn004"), style: style), options)
            .Paragraph(p => p.AddText("AfterColumns"), style: style);
    }
    private static UglyToad.PdfPig.Content.Word Word(PdfPigDocument pdf, int page, string marker) =>
        pdf.GetPage(page).GetWords().Single(word => word.Text == marker);
}
