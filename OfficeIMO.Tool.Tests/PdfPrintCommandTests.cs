using OfficeIMO.Pdf;
using OfficeIMO.Tool.Commands.Workflow;
using OfficeIMO.Workflows;
using Xunit;

namespace OfficeIMO.Tool.Tests;

public sealed class PdfPrintCommandTests {
    [Fact]
    public async Task PrintSendsTheComposedJobAndExplicitDeliverySettings() {
        string path = Path.Combine(Path.GetTempPath(), "officeimo-print-cli-" + Guid.NewGuid().ToString("N") + ".pdf");
        PdfPlainTextConverter.ToPdfDocumentResult("Print command content").TrySaveWithEvidence(path).RequireSuccess();
        try {
            var service = new PrinterBoundary();
            using var output = new StringWriter();
            using var error = new StringWriter();
            int code = await WorkflowPrintCommand.RunAsync(["print", path, "--printer", "Test queue", "--copies", "2", "--duplex", "long", "--dpi", "72",
                "--pages-per-sheet", "9", "--scale", "custom", "--custom-scale", "25", "--alignment", "bottom-right",
                "--margin-left", "24", "--color", "grayscale"],
                output, error, CancellationToken.None, service);
            Assert.Equal((int)OfficeImoToolExitCode.Success, code);
            Assert.Single(service.Document!.Plan.Sheets);
            Assert.Equal(PdfPrintColorMode.Grayscale, service.Document.Plan.ColorMode);
            PdfPrintPlacement placement = Assert.Single(service.Document.Plan.Sheets[0].Placements);
            Assert.Equal(0.25, placement.Scale);
            Assert.Equal(24, placement.SlotX);
            Assert.Equal(placement.SlotX + placement.SlotWidth - placement.Width, placement.X, 6);
            Assert.Equal(2, service.Options!.Copies);
            Assert.Equal(PdfPrintDuplex.LongEdge, service.Options.Duplex);
            Assert.Equal(72, service.Options.MaximumRasterDpi);
            Assert.Contains("Physical delivery is unconfirmed", output.ToString());
        } finally { File.Delete(path); }
    }

    [Theory]
    [InlineData(null, 612, 792)]
    [InlineData("A3", 842, 1190)]
    public async Task PaperDefaultsToThePrinterMediaUnlessChosen(string? paper, double width, double height) {
        string path = Path.Combine(Path.GetTempPath(), "officeimo-print-cli-" + Guid.NewGuid().ToString("N") + ".pdf");
        PdfPlainTextConverter.ToPdfDocumentResult("Print command content").TrySaveWithEvidence(path).RequireSuccess();
        try {
            var service = new PrinterBoundary { DefaultMedia = new("Letter", "US Letter", PdfPageSizes.Letter) };
            string[] arguments = paper is null ? ["print", path, "--printer", "Test queue"] : ["print", path, "--printer", "Test queue", "--paper", paper];
            int code = await WorkflowPrintCommand.RunAsync(arguments, TextWriter.Null, TextWriter.Null, CancellationToken.None, service);
            Assert.Equal((int)OfficeImoToolExitCode.Success, code);
            PdfPageSize sheet = service.Document!.Plan.Sheets[0].PaperSize;
            Assert.Equal(width, sheet.Width.ToPoints(), 0);
            Assert.Equal(height, sheet.Height.ToPoints(), 0);
        } finally { File.Delete(path); }
    }

    private sealed class PrinterBoundary : IPdfPrinterService {
        public Task<PdfPrintJobStatus> GetJobAsync(string printerName, string jobId, CancellationToken cancellationToken = default) =>
            Task.FromResult(new PdfPrintJobStatus(printerName, jobId, PdfPrintJobState.Completed));
        public Task<PdfPrintJobStatus> CancelJobAsync(string printerName, string jobId, CancellationToken cancellationToken = default) =>
            Task.FromResult(new PdfPrintJobStatus(printerName, jobId, PdfPrintJobState.Canceled));
        public PdfPrinterMedia? DefaultMedia { get; init; }
        public PdfPrintJob? Document { get; private set; }
        public PdfPrintDeliveryOptions? Options { get; private set; }
        public Task<IReadOnlyList<PdfPrinterInfo>> GetPrintersAsync(CancellationToken cancellationToken = default) =>
            Task.FromResult<IReadOnlyList<PdfPrinterInfo>>([new("Test queue", false, false)]);
        public Task<PdfPrinterCapabilities> GetCapabilitiesAsync(string printerName, CancellationToken cancellationToken = default) =>
            Task.FromResult(new PdfPrinterCapabilities { PrinterName = printerName, DefaultMedia = DefaultMedia });
        public Task<PdfPrintSubmission> SubmitAsync(PdfPrintJob job, PdfPrintDeliveryOptions options, CancellationToken cancellationToken = default) {
            Document = job; Options = options;
            return Task.FromResult(new PdfPrintSubmission(options.PrinterName, "test-42", job.Plan.Sheets.Count, options.Copies, null));
        }
    }
}
