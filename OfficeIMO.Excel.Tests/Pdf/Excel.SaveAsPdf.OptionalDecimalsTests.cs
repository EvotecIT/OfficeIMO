using OfficeIMO.Excel;
using OfficeIMO.Excel.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Excel {
    [Theory]
    [InlineData(-0.5d, "0.###%;[Red]0.###%", "50%", true)]
    [InlineData(0.5d, "0.###%;[Red]0.###%", "50%", false)]
    [InlineData(-0.5d, "0.###%;[Red](0.###%)", "(50%)", true)]
    [InlineData(0.5d, "0.###%\"[Red]\"", "50%[Red]", false)]
    public void Numeric_format_color_preserves_negative_display_semantics_in_pdf(double value, string format, string expected, bool red) {
        using var workbook = ExcelDocument.Create();
        workbook.AddWorksheet("Formats").CellAt(1, 1).SetValue(value).SetNumberFormat(format).SetFontColor("0000FF");
        using var saved = new MemoryStream(); workbook.Save(saved); saved.Position = 0;
        using var reopened = ExcelDocument.Load(saved);
        byte[] bytes = reopened.ToPdfBytes(new ExcelToPdfOptions { IncludeSheetHeadings = false, HeaderRowCount = 0 });
        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        Assert.Equal(expected, Assert.Single(pdf.GetPages()).Text.Trim());
        string operators = PdfOperatorSearchText.From(bytes);
        Assert.Equal(red, operators.Contains("1 0 0 rg", StringComparison.Ordinal));
        Assert.Equal(!red, operators.Contains("0 0 1 rg", StringComparison.Ordinal));
    }

    [Theory]
    [InlineData(0.5d, "0.###############%", "50%")]
    [InlineData(0d, "0.###%", "0%")]
    [InlineData(0.125d, "0.###%", "12.5%")]
    [InlineData(0.5d, "0.00%", "50.00%")]
    [InlineData(1234.5d, "#,##0.###", "1,234.5")]
    [InlineData(1d, "0.0##", "1.0")]
    [InlineData(-0.5d, "0.###%;[Red](0.###%)", "(50%)")]
    [InlineData(-0.5d, "0.###%;[Red]0.###%", "50%")]
    [InlineData(5d, "0\"%\"", "5%")]
    [InlineData(1234d, "#,##0;0;0;@;[>0]0\"extra\"", "1,234")]
    public void Saved_numeric_formats_use_shared_display_semantics_in_pdf(double value, string format, string expected) {
        using var workbook = ExcelDocument.Create();
        workbook.AddWorksheet("Formats").CellAt(1, 1).SetValue(value).SetNumberFormat(format);
        using var saved = new MemoryStream(); workbook.Save(saved); saved.Position = 0;
        using var reopened = ExcelDocument.Load(saved);
        byte[] bytes = reopened.ToPdfBytes(new ExcelToPdfOptions { IncludeSheetHeadings = false, HeaderRowCount = 0 });
        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        Assert.Equal(expected, Assert.Single(pdf.GetPages()).Text.Trim());
    }
}
