using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Html;
using Xunit;

namespace OfficeIMO.Html.Tests;

public sealed class ExcelImageOptionsTests {
    [Fact]
    public void WorksheetVisualReviewPreservesCellAndEncodedByteLimits() {
        using var document = ExcelDocument.Create();
        var sheet = document.AddSheet("Data");
        sheet.Cell(1, 1).SetValue("One");
        sheet.Cell(1, 2).SetValue("Two");
        var options = new ExcelHtmlSaveOptions {
            ExportProfile = ExcelHtmlExportProfile.VisualReview,
            VisualOptions = new ExcelImageExportOptions { MaximumRenderedCells = 1 }
        };
        Assert.Throws<InvalidOperationException>(() => sheet.ToHtml(options));

        options.VisualOptions.MaximumRenderedCells = 2;
        options.VisualOptions.MaximumTotalEncodedBytes = 32;
        Assert.Throws<OfficeImageExportBatchLimitException>(() => sheet.ToHtml(options));

        options.VisualOptions.MaximumTotalEncodedBytes = 1_000_000;
        string html = sheet.ToHtml(options);
        Assert.Contains("<svg", html, StringComparison.Ordinal);
        Assert.Equal(2, options.VisualOptions.MaximumRenderedCells);
    }
}
