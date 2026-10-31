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

    [Theory]
    [InlineData("•", 4)]
    [InlineData("1.", 3)]
    [InlineData("1)", 3)]
    [InlineData("(1)", 3)]
    public void CompactListsRemainListsAcrossTableRecoveryPaths(string marker, int count) {
        var bands = new List<List<TextLayoutEngine.TextLine>>();
        string[] labels = { "Alpha", "Bravo", "Charlie", "Delta" };
        for (int index = 0; index < count; index++) {
            string font = index == 0 ? "Helvetica-Bold" : "Helvetica";
            bands.Add(new() { Line(520D - index * 20D,
                (marker, 40D, 10D, font), (labels[index], 80D, 35D, font)) });
        }

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

    [Theory]
    [InlineData(0D, 25D, false)]
    [InlineData(0D, 35D, false)]
    [InlineData(360D, 25D, false)]
    [InlineData(360D, 35D, false)]
    [InlineData(0D, 35D, true)]
    [InlineData(360D, 35D, true)]
    public void WiderAlignedLastCellKeepsSideAnnotationsOutsideTheTable(double annotationX, double lastCellWidth, bool splitLastCell) {
        var middle = new List<(string Text, double X, double Advance, string Font)> {
            ("HTML", 42D, 28D, "Helvetica"), ("OfficeIMO.Html", 112D, 65D, "Helvetica"),
            ("Ready", 282D, lastCellWidth, "Helvetica"), ("Ready", annotationX, 25D, "Helvetica")
        };
        if (splitLastCell) middle.Add(("+tail", 282D + lastCellWidth, 20D, "Helvetica"));
        var bands = new List<List<TextLayoutEngine.TextLine>> {
            new() { Line(220D, ("Area", 42D, 20D, "Helvetica-Bold"), ("Owner", 112D, 30D, "Helvetica-Bold"),
                ("Status", 282D, 30D, "Helvetica-Bold")) },
            new() { Line(200D, middle.ToArray()) },
            new() { Line(180D, ("PDF", 42D, 20D, "Helvetica"), ("OfficeIMO.Html.Pdf", 112D, 80D, "Helvetica"),
                ("Stable", 282D, 26D, "Helvetica")) }
        };
        PdfStructuredTable table = Assert.Single(TableDetector.DetectTablesFromBands(bands));
        Assert.Equal(3, table.Rows.Count);
        Assert.Equal(new[] { "Area", "Owner", "Status" }, table.Rows[0]);
        Assert.Equal(new[] { "HTML", "OfficeIMO.Html", splitLastCell ? "Ready+tail" : "Ready" }, table.Rows[1]);
        Assert.Equal(new[] { "PDF", "OfficeIMO.Html.Pdf", "Stable" }, table.Rows[2]);
        PdfTextSpan annotation = Assert.Single(bands[1][0].Spans, span => span.X == annotationX);
        Assert.DoesNotContain(annotation, table.SourceRuns);
    }

    private static TextLayoutEngine.TextLine Line(double y, params (string Text, double X, double Advance, string Font)[] content) {
        List<PdfTextSpan> spans = content.OrderBy(item => item.X)
            .Select(item => new PdfTextSpan(item.Text, item.Font, 10D, item.X, y, item.Advance, baseFont: item.Font)).ToList();
        return new TextLayoutEngine.TextLine(y, spans.Min(span => span.X), spans.Max(span => span.X + span.Advance),
            string.Join(" ", spans.Select(span => span.Text)), spans);
    }
}
