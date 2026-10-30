using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfEmptyColumnPageBreakTests {
    [Theory]
    [InlineData(2, 1, false)]
    [InlineData(2, 1, true)]
    [InlineData(3, 2, false)]
    [InlineData(3, 2, true)]
    public void PageBreakAfterAnEmptyLaterColumnReturnsToTheFirstColumn(int columns, int breaks, bool header) {
        using var pdf = PdfPigDocument.Open(Render(columns, breaks, header, preserve: false));
        int expectedPages = header ? 2 : 1;
        Assert.Equal(expectedPages, pdf.NumberOfPages);
        var marker = Assert.Single(pdf.GetPages().SelectMany(page => page.GetWords()).Where(word => word.Text == "AFTER"));
        Assert.InRange(marker.BoundingBox.Left, 39.9, 40.1);
        Assert.Contains("AFTER", pdf.GetPage(expectedPages).Text);
        if (header) {
            Assert.Contains("HEADER", pdf.GetPage(1).Text);
            Assert.DoesNotContain("AFTER", pdf.GetPage(1).Text);
            Assert.Contains("HEADER", pdf.GetPage(2).Text);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PageBreakOnAFreshFirstColumnDoesNotEmitAnotherHeaderOnlyPage(bool rollover) {
        using var pdf = PdfPigDocument.Open(Render(2, rollover ? 2 : 0, header: true, preserve: false));
        int expectedPages = rollover ? 2 : 1;
        Assert.Equal(expectedPages, pdf.NumberOfPages);
        var marker = Assert.Single(pdf.GetPage(expectedPages).GetWords().Where(word => word.Text == "AFTER"));
        Assert.InRange(marker.BoundingBox.Left, 39.9, 40.1);
        if (rollover) Assert.DoesNotContain("AFTER", pdf.GetPage(1).Text);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    public void PreservedPageBreakRetainsTheEmptyPhysicalPage(int breaks) {
        using var pdf = PdfPigDocument.Open(Render(2, breaks, header: false, preserve: true));
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Empty(pdf.GetPage(1).Letters);
        var marker = Assert.Single(pdf.GetPage(2).GetWords());
        Assert.Equal("AFTER", marker.Text);
        Assert.InRange(marker.BoundingBox.Left, 39.9, 40.1);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void ExplicitBreakPreservesASpacingOnlyBodyPage(bool columns, bool nextColumn) {
        var document = PdfDocument.Create(new PdfCreateOptions { CompressContentStreams = false });
        document.Compose(builder => builder.Page(page => page.Content(body => {
            void Content(PdfContentBuilder content) {
                content.Spacer(global::OfficeIMO.OfficeLength.Points(12));
                if (nextColumn) content.ColumnBreak();
                content.PageBreak(preserveEmptyPage: false);
                content.Paragraph(paragraph => paragraph.Text("AFTER"));
            }
            if (columns) body.Columns(Content, new PdfMultiColumnOptions { ColumnCount = 2, BalanceLastPage = false });
            else Content(body);
        })));
        using var pdf = PdfPigDocument.Open(document.ToBytes());
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Empty(pdf.GetPage(1).Letters);
        Assert.Contains("AFTER", pdf.GetPage(2).Text);
    }

    private static byte[] Render(int columns, int breaks, bool header, bool preserve) {
        var document = PdfDocument.Create(new PdfCreateOptions {
            PageWidth = global::OfficeIMO.OfficeLength.Points(500), PageHeight = global::OfficeIMO.OfficeLength.Points(400), MarginLeft = global::OfficeIMO.OfficeLength.Points(40), MarginRight = global::OfficeIMO.OfficeLength.Points(40),
            MarginTop = global::OfficeIMO.OfficeLength.Points(40), MarginBottom = global::OfficeIMO.OfficeLength.Points(40), DefaultFontSize = global::OfficeIMO.OfficeLength.Points(12), CompressContentStreams = false
        });
        document.Compose(builder => builder.Page(page => {
            if (header) page.Header(running => running.Content(content => content.Paragraph(paragraph => paragraph.Text("HEADER")), 18));
            page.Content(body => body.Columns(content => {
                for (int index = 0; index < breaks; index++) content.ColumnBreak();
                content.PageBreak(preserve);
                content.Paragraph(paragraph => paragraph.Text("AFTER"), style: new PdfParagraphStyle {
                    LineSpacing = PdfLineSpacing.Exact(global::OfficeIMO.OfficeLength.Points(20)), SpacingBefore = global::OfficeIMO.OfficeLength.Points(0), SpacingAfter = global::OfficeIMO.OfficeLength.Points(0), WidowControl = false
                });
            }, new PdfMultiColumnOptions { ColumnCount = columns, Gap = 20, BalanceLastPage = false }));
        }));
        return document.ToBytes();
    }
}
