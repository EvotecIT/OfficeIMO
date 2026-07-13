using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ExcelPdfAsyncContractTests {
    [Fact]
    public async Task PdfAsyncSavesPerformIoAndDoNotTurnCancellationIntoFailureResults() {
        using ExcelDocument workbook = ExcelDocument.Create();
        workbook.AddWorksheet("Data").CellValue(1, 1, "Async Excel PDF");
        using var output = new MemoryStream();

        await workbook.SaveAsPdfAsync(output);

        Assert.Equal("%PDF-", System.Text.Encoding.ASCII.GetString(output.ToArray(), 0, 5));
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            workbook.SaveAsPdfAsync(new MemoryStream(), cancellationToken: cancellation.Token));
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            workbook.TrySaveAsPdfAsync(new MemoryStream(), cancellationToken: cancellation.Token));
    }
}
