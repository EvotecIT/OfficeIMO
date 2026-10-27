using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfPeripheralTableRecoveryTests {
    [Theory]
    [InlineData(0D)]
    [InlineData(200D)]
    public void HangingListWithEmphasizedFirstItemAndSideNoteRemainsAList(double annotationX) {
        var bands = new List<List<TextLayoutEngine.TextLine>> {
            new() { Line(520D, ("•", 40D, 5D, "Helvetica"), ("Alpha", 80D, 30D, "Helvetica-Bold")) },
            new() { Line(500D, ("•", 40D, 5D, "Helvetica"), ("Bravo", 80D, 30D, "Helvetica"),
                ("Note", annotationX, 20D, "Helvetica")) },
            new() { Line(480D, ("•", 40D, 5D, "Helvetica"), ("Charlie", 80D, 30D, "Helvetica")) }
        };

        Assert.Empty(TableDetector.DetectTablesFromBands(bands));
    }

    [Fact]
    public void NumberedRowsWithARealColumnHeaderRemainATable() {
        var bands = new List<List<TextLayoutEngine.TextLine>> {
            new() { Line(520D, ("No.", 40D, 15D, "Helvetica-Bold"), ("Name", 80D, 45D, "Helvetica-Bold")) },
            new() { Line(500D, ("1.", 40D, 10D, "Helvetica"), ("Alpha", 80D, 45D, "Helvetica")) },
            new() { Line(480D, ("2.", 40D, 10D, "Helvetica"), ("Bravo", 80D, 45D, "Helvetica")) }
        };

        PdfStructuredTable table = Assert.Single(TableDetector.DetectTablesFromBands(bands));

        Assert.Equal(new[] { "No.", "Name" }, table.Rows[0]);
        Assert.Equal(new[] { "1.", "Alpha" }, table.Rows[1]);
        Assert.Equal(new[] { "2.", "Bravo" }, table.Rows[2]);
    }

    private static TextLayoutEngine.TextLine Line(double y, params (string Text, double X, double Advance, string Font)[] content) {
        List<PdfTextSpan> spans = content.OrderBy(item => item.X)
            .Select(item => new PdfTextSpan(item.Text, item.Font, 10D, item.X, y, item.Advance, baseFont: item.Font)).ToList();
        return new TextLayoutEngine.TextLine(y, spans.Min(span => span.X), spans.Max(span => span.X + span.Advance),
            string.Join(" ", spans.Select(span => span.Text)), spans);
    }
}
