using System;
using System.Collections.Generic;
using System.Linq;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfDeferredTableRowGroupingTests {
    [Theory]
    [InlineData("first")]
    [InlineData("last")]
    [InlineData("roles")]
    public void TableDeferred_MinimumBodyRowsPreservePaginationAcrossBatches(string mode) {
        using var expected = PdfPigDocument.Open(RenderMinimumBodyRows(mode, 20));
        using var actual = PdfPigDocument.Open(RenderMinimumBodyRows(mode, 1));
        Assert.Equal(expected.NumberOfPages, actual.NumberOfPages);
        for (int page = 1; page <= expected.NumberOfPages; page++) {
            var expectedWords = expected.GetPage(page).GetWords().ToArray();
            var actualWords = actual.GetPage(page).GetWords().ToArray();
            Assert.Equal(expectedWords.Length, actualWords.Length);
            for (int word = 0; word < expectedWords.Length; word++) {
                Assert.Equal(expectedWords[word].Text, actualWords[word].Text);
                Assert.Equal(expectedWords[word].BoundingBox.Left, actualWords[word].BoundingBox.Left, 3);
                Assert.Equal(expectedWords[word].BoundingBox.Top, actualWords[word].BoundingBox.Top, 3);
            }
        }
    }

    private static byte[] RenderMinimumBodyRows(string mode, int batchSize) {
        bool roles = mode == "roles";
        var rows = new List<string[]>();
        if (roles) rows.Add(new[] { "Header" });
        rows.AddRange(Enumerable.Range(0, 7).Select(row => new[] { $"Body{row}" }));
        if (roles) rows.Add(new[] { "Footer" });
        var style = new PdfTableStyle {
            HeaderRowCount = roles ? 1 : 0, FooterRowCount = roles ? 1 : 0,
            HeaderFill = null, FooterFill = null, BorderColor = null, RowStripeFill = null,
            CellPaddingX = 0, CellPaddingY = 0, FontSize = 12, LineHeight = 1,
            FixedRowHeights = Enumerable.Repeat<double?>(20, rows.Count).ToList(),
            ColumnWidthPoints = new List<double?> { 100 }, PreferredWidth = 100,
            MinimumBodyRowsOnFirstPage = mode == "last" ? 0 : 2,
            MinimumBodyRowsOnLastPage = mode == "first" ? 0 : 2
        };
        var document = PdfDocument.Create(new PdfOptions {
            PageWidth = 320, PageHeight = roles ? 160 : 120,
            MarginLeft = 30, MarginRight = 30, MarginTop = 30, MarginBottom = 30
        });
        document.Compose(root => root.Page(page => page.Content(content => {
            if (mode != "last") content.Paragraph(paragraph => paragraph.Text("Preamble\nContinuation"),
                style: new PdfParagraphStyle { FontSize = 12, LineHeight = 1, SpacingAfter = 0 });
            content.TableDeferred(() => rows, batchSize, style: style);
        })));
        return document.ToBytes();
    }
}
