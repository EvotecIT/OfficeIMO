using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfDocumentVisualQualityTests {
    [Theory]
    [InlineData(false, 0)]
    [InlineData(true, 0)]
    [InlineData(false, 2)]
    [InlineData(true, 2)]
    public void Table_MergedTailSurvivesTheLastNeighbourRowSplitting(bool inRowColumn, double cellSpacing) {
        PdfTableStyle style = TableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.CellPaddingX = style.CellPaddingY = 0;
        style.MinRowHeight = 0;
        style.CellSpacing = cellSpacing;
        const int tokenCount = 40;
        PdfTableCell[][] rows = {
            new[] { PdfTableCell.Merge(string.Join("\n", Enumerable.Range(1, tokenCount).Select(n => $"Span{n}")), rowSpan: 3), PdfTableCell.TextCell("Neighbour0") },
            new[] { PdfTableCell.TextCell("Neighbour1") },
            new[] { PdfTableCell.TextCell(string.Join("\n", Enumerable.Range(1, 8).Select(n => $"LongNeighbour{n}"))) },
            new[] { PdfTableCell.TextCell("AfterSpan"), PdfTableCell.TextCell("AfterNeighbour") }
        };
        PdfDocument document = PdfDocument.Create(new PdfOptions {
            PageWidth = 300, PageHeight = 120, MarginLeft = 20, MarginRight = 20,
            MarginTop = 20, MarginBottom = 20, DefaultFontSize = 12
        });
        if (inRowColumn)
            document.Compose(builder => builder.Page(page => page.Content(content =>
                content.Row(row => row.PercentColumn(100, column => column.Table(rows, style: style))))));
        else document.Table(rows, style: style);
        using PdfPigDocument pdf = PdfPigDocument.Open(document.ToBytes());
        var words = pdf.GetPages().SelectMany(page => page.GetWords().Select(word => (page.Number, Word: word))).ToArray();
        foreach (int n in Enumerable.Range(1, tokenCount)) {
            var token = Assert.Single(words, item => item.Word.Text == $"Span{n}");
            Assert.True(token.Word.BoundingBox.Bottom >= 19D, $"Span{n} crosses page {token.Number}'s bottom margin.");
        }
        foreach (int n in Enumerable.Range(1, 8)) Assert.Single(words, item => item.Word.Text == $"LongNeighbour{n}");
        var last = Assert.Single(words, item => item.Word.Text == $"Span{tokenCount}");
        var following = Assert.Single(words, item => item.Word.Text == "AfterSpan");
        Assert.True(following.Number > last.Number || following.Number == last.Number && following.Word.BoundingBox.Top < last.Word.BoundingBox.Bottom,
            "The following physical row must follow the complete merged-cell tail.");
    }

    [Theory]
    [InlineData(false, 6)]
    [InlineData(false, 12)]
    [InlineData(false, 18)]
    [InlineData(true, 6)]
    [InlineData(true, 12)]
    [InlineData(true, 18)]
    public void Table_MergedTextAndPaintContinueWithinEachPage(bool inRowColumn, int tokenCount) {
        PdfTableStyle style = TableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.CellPaddingX = style.CellPaddingY = 0;
        style.MinRowHeight = 0;
        style.BorderColor = style.RowSeparatorColor = style.RowStripeFill = null;
        style.CellFills = new() { [(0, 0)] = new PdfColor(.31, .41, .51) };
        style.CellBorders = new() { [(0, 0)] = new PdfCellBorder { Color = new PdfColor(.61, .21, .11), Width = .5 } };
        PdfTableCell[][] rows = {
            new[] { PdfTableCell.Merge(string.Join("\n", Enumerable.Range(1, tokenCount).Select(n => $"Token{n}")), rowSpan: 3), PdfTableCell.TextCell("Neighbour0") },
            new[] { PdfTableCell.TextCell("Neighbour1") },
            new[] { PdfTableCell.TextCell("Neighbour2") }
        };
        PdfDocument document = PdfDocument.Create(new PdfOptions {
            PageWidth = 300, PageHeight = 120, MarginLeft = 20, MarginRight = 20,
            MarginTop = 20, MarginBottom = 20, DefaultFontSize = 12
        });
        if (inRowColumn)
            document.Compose(builder => builder.Page(page => page.Content(content =>
                content.Row(row => row.PercentColumn(100, column => column.Table(rows, style: style))))));
        else document.Table(rows, style: style);
        byte[] bytes = document.ToBytes();
        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        Assert.True(pdf.NumberOfPages > 1);
        var words = pdf.GetPages().SelectMany(page => page.GetWords().Select(word => (page.Number, Word: word))).ToArray();
        foreach (int n in Enumerable.Range(1, tokenCount)) {
            var token = Assert.Single(words, item => item.Word.Text == $"Token{n}");
            Assert.True(token.Word.BoundingBox.Bottom >= 19D, $"Token{n} crosses page {token.Number}'s bottom margin.");
        }
        Assert.True(Assert.Single(words, item => item.Word.Text == $"Token{tokenCount}").Number > 1);
        foreach (int n in Enumerable.Range(0, 3)) Assert.Single(words, item => item.Word.Text == $"Neighbour{n}");
        foreach (int n in Enumerable.Range(1, pdf.NumberOfPages)) {
            string content = string.Join("\n", GetPageContentStreams(bytes, n));
            Assert.All(ExtractPaintedRectangles(content, "0.31 0.41 0.51 rg", "f"), fill => {
                Assert.True(fill.Y >= 19.9D && fill.Y + fill.H <= 100.1D,
                    $"Merged fill crosses page {n}'s body: bottom {fill.Y}, height {fill.H}.");
            });
        }
    }
}
