using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfTableCellFitTextTests {
    [Theory]
    [InlineData(" AB")]
    [InlineData("A    B C")]
    [InlineData("AB​")]
    public void FitTextEndsTheLastPaintedLetterAtTheCellEdge(string text) {
        PdfTableCell Cell(string value, bool fitText, OfficeIMO.OfficeAlignment? align = null) {
            var runs = new[] { PdfTextRun.Normal(value) };
            return new PdfTableCell(runs, new[] { new PdfTableCellParagraph(runs, align: align, fitText: fitText) });
        }

        var table = new PdfTableStyle {
            HeaderRowCount = 0, CellPaddingX = global::OfficeIMO.OfficeLength.Points(0), CellPaddingY = global::OfficeIMO.OfficeLength.Points(0),
            BorderWidth = global::OfficeIMO.OfficeLength.Points(0), FontSize = global::OfficeIMO.OfficeLength.Points(12)
        };
        // A right-aligned reference in the same column marks the inner right edge.
        PdfTableCell[][] rows = { new[] { Cell(text, fitText: true) }, new[] { Cell("Z", fitText: false, OfficeIMO.OfficeAlignment.Right) } };
        var options = new PdfCreateOptions {
            PageSetup = {
                Size = new PdfPageSize(global::OfficeIMO.OfficeLength.Points(360), global::OfficeIMO.OfficeLength.Points(240)),
                Margins = { Left = global::OfficeIMO.OfficeLength.Points(30), Right = global::OfficeIMO.OfficeLength.Points(30), Top = global::OfficeIMO.OfficeLength.Points(30), Bottom = global::OfficeIMO.OfficeLength.Points(30) }
            },
            DefaultFont = PdfStandardFont.Helvetica, TextWhitespaceMode = PdfTextWhitespaceMode.Collapse
        };
        byte[] bytes = PdfDocument.Create(builder => builder.Content(content =>
            content.Canvas(canvas => canvas.Table(rows, global::OfficeIMO.OfficeLength.Points(30), global::OfficeIMO.OfficeLength.Points(30),
                global::OfficeIMO.OfficeLength.Points(200), global::OfficeIMO.OfficeLength.Points(120), table))), options).ToBytes();

        using var pdf = PdfPigDocument.Open(bytes);
        var letters = pdf.GetPage(1).Letters.Where(letter => !string.IsNullOrWhiteSpace(letter.Value)).ToArray();
        double edge = letters.Single(letter => letter.Value == "Z").EndBaseLine.X;
        var fitted = letters.TakeWhile(letter => letter.Value != "Z").ToArray();

        Assert.Equal(text.Replace(" ", string.Empty).Replace("​", string.Empty), string.Concat(fitted.Select(letter => letter.Value)));
        Assert.InRange(fitted[fitted.Length - 1].EndBaseLine.X, edge - 1D, edge + 0.05D);
    }

    [Fact]
    public void FitTextRefitsAuthoredLinesThatContinueIntoAWiderColumn() {
        string text = string.Join("\n", Enumerable.Range(1, 24).Select(index => "Line" + index.ToString("D2")));
        var runs = new[] { PdfTextRun.Normal(text) };
        var cell = new PdfTableCell(runs, new[] {
            new PdfTableCellParagraph(runs, fontSize: global::OfficeIMO.OfficeLength.Points(12),
                lineSpacing: PdfLineSpacing.Exact(global::OfficeIMO.OfficeLength.Points(20)), fitText: true)
        });
        var options = new PdfCreateOptions {
            PageSetup = {
                Size = new PdfPageSize(global::OfficeIMO.OfficeLength.Points(500), global::OfficeIMO.OfficeLength.Points(400)),
                Margins = { Left = global::OfficeIMO.OfficeLength.Points(40), Right = global::OfficeIMO.OfficeLength.Points(40), Top = global::OfficeIMO.OfficeLength.Points(40), Bottom = global::OfficeIMO.OfficeLength.Points(40) }
            },
            DefaultFontSize = global::OfficeIMO.OfficeLength.Points(12)
        };
        // Columns span x 40-140 and 160-460.
        var columns = new PdfMultiColumnOptions {
            BalanceLastPage = false,
            ColumnDefinitions = new[] {
                new PdfFlowColumn(PdfColumnWidth.Fixed(global::OfficeIMO.OfficeLength.Points(100)), global::OfficeIMO.OfficeLength.Points(20)),
                new PdfFlowColumn(PdfColumnWidth.Fixed(global::OfficeIMO.OfficeLength.Points(300)))
            }
        };
        byte[] bytes = PdfDocument.Create(options).Columns(content => content.AddTable(new[] { new[] { cell } },
            style: new PdfTableStyle {
                HeaderRowCount = 0, CellPaddingX = global::OfficeIMO.OfficeLength.Points(0), CellPaddingY = global::OfficeIMO.OfficeLength.Points(0),
                BorderWidth = global::OfficeIMO.OfficeLength.Points(0), FontSize = global::OfficeIMO.OfficeLength.Points(12),
                SpacingBefore = global::OfficeIMO.OfficeLength.Points(0), SpacingAfter = global::OfficeIMO.OfficeLength.Points(0)
            }), columns).ToBytes();

        using var pdf = PdfPigDocument.Open(bytes);
        var lineEnds = pdf.GetPage(1).Letters.Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
            .GroupBy(letter => (Column: letter.StartBaseLine.X > 150D, Y: Math.Round(letter.StartBaseLine.Y, 1)))
            .Select(line => (line.Key.Column, End: line.Max(letter => letter.EndBaseLine.X))).ToArray();

        Assert.Equal(24, lineEnds.Length);
        Assert.Contains(lineEnds, line => line.Column);
        Assert.All(lineEnds, line => Assert.InRange(line.End, (line.Column ? 460D : 140D) - 1D, (line.Column ? 460D : 140D) + 0.05D));
    }
}
