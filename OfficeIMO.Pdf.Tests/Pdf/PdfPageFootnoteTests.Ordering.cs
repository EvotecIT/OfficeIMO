using System;
using System.Linq;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfPageFootnoteTests {
    [Fact]
    public void StaggeredSpanReferencesUseSourceOrderAndOnePageSeparatorAcrossRepeatedSerialization() {
        var options = Options();
        options.PageHeight = 260;
        var document = PdfDocument.Create(options);
        // Deliberately create the right-cell definition first. Source order comes from
        // the body inventory; the earlier left-cell reference is admitted in row two.
        var laterSource = document.Content.DefineFootnote(body => body.Text("LaterSourceNote", style: Style()),
            new PdfFootnoteOptions { Separator = body => body.Text("PageSeparator", style: Style()) });
        var earlierSource = document.Content.DefineFootnote(body => body.Text("EarlierSourceNote", style: Style()),
            new PdfFootnoteOptions { Separator = body => body.Text("UnusedLaterSeparator", style: Style()) });
        var style = FootnoteTableStyle();
        style.MinimumBodyRowsOnFirstPage = 0;
        style.MinimumBodyRowsOnLastPage = 0;
        document.Content.Table(new[] {
            new[] {
                new PdfTableCell(TableReferenceLines(12, 8, "LRef", earlierSource), rowSpan: 2),
                new PdfTableCell(new[] { PdfTextRun.Normal("RRef").WithFootnoteReference(laterSource), PdfTextRun.Normal("\nRightFirstTail") })
            },
            new[] { new PdfTableCell(string.Join("\n", Enumerable.Range(1, 10).Select(index => "RightTail" + index))) }
        }, style: style);

        foreach (byte[] bytes in new[] { document.ToBytes(), document.ToBytes() }) {
            using var pdf = PdfPigDocument.Open(bytes);
            Assert.Equal(1, pdf.NumberOfPages);
            Assert.Contains("LRef", pdf.GetPage(1).Text);
            Assert.Contains("RRef", pdf.GetPage(1).Text);
            Assert.Single(pdf.GetPage(1).GetWords(), word => word.Text == "EarlierSourceNote");
            Assert.Single(pdf.GetPage(1).GetWords(), word => word.Text == "LaterSourceNote");
            Assert.Single(pdf.GetPage(1).GetWords(), word => word.Text == "PageSeparator");
            Assert.DoesNotContain("UnusedLaterSeparator", pdf.GetPage(1).Text);
            double firstNote = TableMarkerBaseline(pdf, 1, "EarlierSourceNote");
            double laterNote = TableMarkerBaseline(pdf, 1, "LaterSourceNote");
            double separator = TableMarkerBaseline(pdf, 1, "PageSeparator");
            Assert.True(separator > firstNote && firstNote > laterNote,
                "The page-owned separator must precede notes ordered by authored reference position, independently of creation and admission order.");
        }
    }

    [Theory]
    [InlineData(2, false)]
    [InlineData(25, true)]
    public void PageWidthChangeRejectsOnlyAnActiveFootnoteContinuation(int noteLines, bool activeContinuation) {
        var document = PdfDocument.Create(Options());
        var note = document.Content.DefineFootnote(body => {
            for (int index = 0; index < noteLines; index++) body.Text("WidthNote" + index, style: Style());
        });
        document.Compose(pages => {
            pages.Page(page => page.Size(260, 140).Content(body =>
                body.Paragraph(p => p.Runs(new[] { PdfTextRun.Normal("Reference").WithFootnoteReference(note) }), style: Style())));
            pages.Page(page => page.Size(240, 140).Content(body => body.Text("NarrowPageBody", style: Style())));
        });
        if (activeContinuation) {
            var exception = Assert.Throws<NotSupportedException>(() => document.ToBytes());
            Assert.Contains("usable page width", exception.Message);
        } else {
            using var pdf = PdfPigDocument.Open(document.ToBytes());
            Assert.Equal(2, pdf.NumberOfPages);
            Assert.Contains("WidthNote1", pdf.GetPage(1).Text);
            Assert.Contains("NarrowPageBody", pdf.GetPage(2).Text);
        }
    }
}
