using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfTableCellFitTextTests {
    [Theory]
    [InlineData(" AB")]
    [InlineData("A    B C")]
    public void FitTextSpreadsOnlyTheSpacesThatCollapsedWhitespaceKeeps(string text) {
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

        Assert.Equal(text.Replace(" ", string.Empty), string.Concat(fitted.Select(letter => letter.Value)));
        Assert.InRange(fitted[fitted.Length - 1].EndBaseLine.X, edge - 1D, edge + 0.05D);
    }
}
