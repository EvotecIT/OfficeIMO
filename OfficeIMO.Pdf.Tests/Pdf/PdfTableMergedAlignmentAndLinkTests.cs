using System.Text;
using System.Text.RegularExpressions;
using OfficeIMO.Pdf;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfDocumentVisualQualityTests {
    [Theory]
    [InlineData(false, PdfCellVerticalAlign.Middle)]
    [InlineData(true, PdfCellVerticalAlign.Middle)]
    [InlineData(false, PdfCellVerticalAlign.Bottom)]
    [InlineData(true, PdfCellVerticalAlign.Bottom)]
    public void MergedAlignmentUsesTheCompleteLogicalCell(bool inRow, PdfCellVerticalAlign alignment) {
        double Position(double height, out int markerPage) {
            PdfTableStyle style = Style();
            style.RowMinHeights = new() { 80, 80, 80 };
            style.CellVerticalAlignments = new() { [(0, 0)] = alignment };
            PdfDocument document = PdfDocument.Create(Options(height));
            AddTable(document, new[] {
                new[] { PdfTableCell.Merge("ANCHOR", rowSpan: 3), PdfTableCell.TextCell("First") },
                new[] { PdfTableCell.TextCell("Second") },
                new[] { PdfTableCell.TextCell("Third") }
            }, style, inRow);
            using PdfPigDocument pdf = PdfPigDocument.Open(document.ToBytes());
            var marker = Assert.Single(pdf.GetPages().SelectMany(page => page.GetWords()
                .Where(word => word.Text == "ANCHOR").Select(word => (page, word))));
            markerPage = marker.page.Number;
            return (marker.page.Number - 1) * 80 + marker.page.Height - marker.word.BoundingBox.Top - 20;
        }
        double whole = Position(300, out int wholePage);
        double split = Position(140, out int splitPage);
        Assert.Equal(1, wholePage);
        Assert.Equal(alignment == PdfCellVerticalAlign.Middle ? 2 : 3, splitPage);
        Assert.InRange(Math.Abs(whole - split), 0, .1);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void EmptyMergedContinuationLinksRemainUnderTheirCell(bool inRow) {
        PdfTableStyle style = Style();
        style.RowMinHeights = new() { 80, 80, 80 };
        PdfDocument document = PdfDocument.Create(Options(140)).TaggedPdfCatalogMarkers();
        AddTable(document, new[] {
            new[] { new PdfTableCell("ANCHOR", rowSpan: 3, linkUri: "https://example.com"), PdfTableCell.TextCell("First") },
            new[] { PdfTableCell.TextCell("Second") },
            new[] { PdfTableCell.TextCell("Third") }
        }, style, inRow);
        byte[] bytes = document.ToBytes();
        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        Assert.Equal(3, pdf.NumberOfPages);
        string raw = Encoding.ASCII.GetString(bytes);
        Dictionary<string, string> objects = Regex.Matches(raw, @"(\d+) 0 obj\s*(.*?)\s*endobj", RegexOptions.Singleline)
            .Cast<Match>().ToDictionary(match => match.Groups[1].Value, match => match.Groups[2].Value);
        string[] links = objects.Values.Where(value => Regex.IsMatch(value, @"/S\s*/Link\b")).ToArray();
        Assert.Equal(3, links.Length);
        foreach (string link in links) {
            string parent = Regex.Match(link, @"/P\s+(\d+) 0 R").Groups[1].Value;
            Assert.True(objects.ContainsKey(parent));
            Assert.Matches(@"/S\s*/TD\b", objects[parent]);
        }
    }
}
