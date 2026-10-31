using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(".", "1,234.56", "1234.56", false)]
    [InlineData(",", "1.234,56", "1234,56", false)]
    [InlineData(".", "1,234.56", "1234.56", true)]
    [InlineData(",", "1.234,56", "1234,56", true)]
    [InlineData(".", "1,234", "1234", false)]
    [InlineData(",", "1.234", "1234", false)]
    [InlineData(".", "1,234", "1234", true)]
    [InlineData(",", "1.234", "1234", true)]
    public void SaveAsPdf_DecimalTabsHonorDocumentSeparatorAndIgnoreGrouping(string symbol, string grouped, string control, bool inTable) {
        using var document = WordDocument.Create();
        document.Settings.DecimalSymbol = symbol;
        WordTableCell? cell = inTable ? document.AddTable(1, 1).Rows[0].Cells[0] : null;
        foreach (string text in new[] { "Grouped\t" + grouped, "Control\t" + control }) {
            WordParagraph paragraph = cell == null ? document.AddParagraph(text) : cell.AddParagraph(text);
            paragraph.FontFamily = "Arial"; paragraph.FontSize = global::OfficeIMO.OfficeLength.Points(12);
            paragraph.AddTabStop(OfficeLength.Points(200), WordTabAlignment.Decimal);
        }
        using var output = new MemoryStream();
        document.SaveAsPdf(output, new WordPdfExportOptions { IncludePageNumbers = false, ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic() });
        using var pdf = PdfPigDocument.Open(output.ToArray());
        var digits = pdf.GetPage(1).Letters.Where(letter => letter.Value.Length == 1 && "0123456789.,".Contains(letter.Value, StringComparison.Ordinal)).ToArray();
        Assert.Equal(grouped + control, string.Concat(digits.Select(letter => letter.Value)));
        double groupedAnchor, controlAnchor;
        if (grouped.Contains(symbol, StringComparison.Ordinal)) {
            groupedAnchor = digits.Take(grouped.Length).Single(letter => letter.Value == symbol).StartBaseLine.X;
            controlAnchor = digits.Skip(grouped.Length).Single(letter => letter.Value == symbol).StartBaseLine.X;
        } else {
            // Without the configured separator, decimal tabs right-align the complete value.
            groupedAnchor = digits[grouped.Length - 1].EndBaseLine.X;
            controlAnchor = digits[digits.Length - 1].EndBaseLine.X;
        }
        Assert.InRange(Math.Abs(groupedAnchor - controlAnchor), 0, .02);
    }
}
