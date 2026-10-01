using OfficeIMO.Pdf;
using OfficeIMO.Tool.Commands.Workflow;
using OfficeIMO.Workflows;
using Xunit;

namespace OfficeIMO.Tool.Tests;

public sealed class PdfPrintCommandTests {
    [Fact]
    public async Task PrintUsesThePreparedSheetsAndExplicitDeliverySettings() {
        string path = Path.Combine(Path.GetTempPath(), "officeimo-print-cli-" + Guid.NewGuid().ToString("N") + ".pdf");
        PdfPlainTextConverter.ToPdfDocumentResult("Print command content").SaveResult(path).RequireSuccess();
        try {
            var service = new PrinterBoundary();
            using var output = new StringWriter();
            using var error = new StringWriter();
            int code = await WorkflowPrintCommand.RunAsync(["print", path, "--printer", "Test queue", "--copies", "2", "--duplex", "long", "--dpi", "72"],
                output, error, CancellationToken.None, service);
            Assert.Equal((int)OfficeImoToolExitCode.Success, code);
            Assert.Single(service.Document!.Sheets);
            Assert.Equal(2, service.Options!.Copies);
            Assert.Equal(PdfPrintDuplex.LongEdge, service.Options.Duplex);
            Assert.Contains("Physical delivery is unconfirmed", output.ToString());
        } finally { File.Delete(path); }
    }

    private sealed class PrinterBoundary : IPdfPrinterService {
        public PdfPreparedPrintDocument? Document { get; private set; }
        public PdfPrintDeliveryOptions? Options { get; private set; }
        public Task<IReadOnlyList<PdfPrinterInfo>> GetPrintersAsync(CancellationToken cancellationToken = default) =>
            Task.FromResult<IReadOnlyList<PdfPrinterInfo>>([new("Test queue", false, false)]);
        public Task<IReadOnlyList<PdfPaperSourceInfo>> GetPaperSourcesAsync(string printerName, CancellationToken cancellationToken = default) =>
            Task.FromResult<IReadOnlyList<PdfPaperSourceInfo>>([]);
        public Task<PdfPrintSubmission> SubmitAsync(PdfPreparedPrintDocument document, PdfPrintDeliveryOptions options, CancellationToken cancellationToken = default) {
            Document = document; Options = options;
            return Task.FromResult(new PdfPrintSubmission(options.PrinterName, "test-42", document.Sheets.Count, options.Copies, null));
        }
    }
}
