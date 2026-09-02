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

        using var optionCancellation = new CancellationTokenSource();
        optionCancellation.Cancel();
        var options = new ExcelPdfSaveOptions { CancellationToken = optionCancellation.Token };
        Assert.ThrowsAny<OperationCanceledException>(() => workbook.TrySaveAsPdf(new MemoryStream(), options));
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            workbook.TrySaveAsPdfAsync(new MemoryStream(), options));
    }

    [Fact]
    public async Task PdfAsyncSaveMethodTokenStopsSynchronousConversionAtTheNextCheckpoint() {
        using ExcelDocument workbook = ExcelDocument.Create();
        ExcelSheet first = workbook.AddWorksheet("First");
        first.CellValue(1, 1, "First sheet");
        first.SetHeaderFooter(headerCenter: "&D");
        ExcelSheet second = workbook.AddWorksheet("Second");
        second.CellValue(1, 1, "Second sheet");
        second.SetHeaderFooter(headerCenter: "&D");
        using var cancellation = new CancellationTokenSource();
        int providerCalls = 0;
        var options = new ExcelPdfSaveOptions {
            HeaderFooterDateTimeProvider = () => {
                providerCalls++;
                cancellation.Cancel();
                return new DateTime(2026, 9, 1);
            }
        };

        await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            workbook.SaveAsPdfAsync(new MemoryStream(), options, cancellation.Token));

        Assert.Equal(1, providerCalls);
    }
}
