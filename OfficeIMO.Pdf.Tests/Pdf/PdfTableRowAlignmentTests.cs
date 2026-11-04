using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfTableRowAlignmentTests {
    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    public void PositionedRecoveryKeepsLaterCenteredCellsSeparateFromASingleLineFirstRow(int wrappedColumn) {
        double wrappedX = wrappedColumn == 0 ? 50D : 210D;
        TextLayoutEngine.TextLine[] lines = {
            CreateLine(552D, ("Product", 50D, 48D, "Helvetica-Bold"),
                ("Details", 210D, 44D, "Helvetica-Bold"), ("Qty", 330D, 24D, "Helvetica-Bold"),
                ("Unit price", 370D, 56D, "Helvetica-Bold"), ("Net amount", 440D, 58D, "Helvetica-Bold")),
            CreateLine(512D, ("Service A", 50D, 58D, "Helvetica"), ("Monthly", 210D, 48D, "Helvetica"),
                ("2", 330D, 8D, "Helvetica"), ("25.00 PLN", 370D, 58D, "Helvetica"), ("50.00 PLN", 440D, 58D, "Helvetica")),
            CreateLine(486D, ("Priority", wrappedX, 46D, "Helvetica")),
            CreateLine(472D, ("Service B", 50D, 58D, "Helvetica"), ("Annual", 210D, 42D, "Helvetica"),
                ("1", 330D, 8D, "Helvetica"), ("75.00 PLN", 370D, 58D, "Helvetica"), ("75.00 PLN", 440D, 58D, "Helvetica")),
            CreateLine(458D, ("response", wrappedX, 52D, "Helvetica"))
        };
        PdfStructuredTable table = Assert.Single(TableDetector.DetectPositionedCellTables(lines),
            candidate => candidate.Kind == "wrapped-positioned-cells-bounded");
        Assert.Equal(new[] { "Service A", "Monthly", "2", "25.00 PLN", "50.00 PLN" }, table.Rows[1]);
        Assert.Equal(wrappedColumn == 0 ? "Priority Service B response" : "Priority Annual response", table.Rows[2][wrappedColumn]);
    }

    private static TextLayoutEngine.TextLine CreateLine(
        double y,
        params (string Text, double X, double Advance, string Font)[] values) {
        var spans = values
            .Select(value => new PdfTextSpan(value.Text, value.Font, 11D, value.X, y, value.Advance, baseFont: value.Font))
            .ToList();
        return new TextLayoutEngine.TextLine(
            y,
            spans.Min(static span => span.X),
            spans.Max(static span => span.X + span.Advance),
            string.Join(" ", values.Select(static value => value.Text)),
            spans);
    }
}
