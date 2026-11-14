using System.Text.Json;
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
                "--margin-left", "24", "--color", "grayscale", "--raster-grayscale"],
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
            Assert.True(service.Options.RasterizeGrayscale);
            using (JsonDocument receipt = JsonDocument.Parse(output.ToString())) {
                Assert.Equal("officeimo.workflow.print.v1", receipt.RootElement.GetProperty("schema").GetString());
                Assert.Equal("test-42", receipt.RootElement.GetProperty("jobId").GetString());
                Assert.Equal(2, receipt.RootElement.GetProperty("copies").GetInt32());
            }

            output.GetStringBuilder().Clear();
            Assert.Equal((int)OfficeImoToolExitCode.Success, await WorkflowPrintCommand.RunAsync(["printers"], output, error, CancellationToken.None, service));
            using (JsonDocument queues = JsonDocument.Parse(output.ToString())) {
                Assert.Equal("officeimo.workflow.printers.v1", queues.RootElement.GetProperty("schema").GetString());
                Assert.Equal("Test queue", Assert.Single(queues.RootElement.GetProperty("printers").EnumerateArray().ToArray()).GetProperty("name").GetString());
            }

            output.GetStringBuilder().Clear();
            var described = new PrinterBoundary { DefaultMedia = new("Letter", "US Letter", PdfPageSizes.Letter) };
            Assert.Equal((int)OfficeImoToolExitCode.Success,
                await WorkflowPrintCommand.RunAsync(["printers", "--capabilities", "Test queue"], output, error, CancellationToken.None, described));
            using JsonDocument capabilities = JsonDocument.Parse(output.ToString());
            Assert.Equal("officeimo.workflow.printer-capabilities.v1", capabilities.RootElement.GetProperty("schema").GetString());
            JsonElement defaultMedia = capabilities.RootElement.GetProperty("defaultMedia");
            Assert.Equal("Letter", defaultMedia.GetProperty("id").GetString());
            Assert.Equal(792, defaultMedia.GetProperty("heightPoints").GetDouble());
            Assert.Equal(JsonValueKind.Null, capabilities.RootElement.GetProperty("supportsDuplex").ValueKind);
            await Assert.ThrowsAsync<WorkflowUsageException>(() => WorkflowPrintCommand.RunAsync(
                ["printers", "--capabilities", "Test queue", "--paper-sources", "Test queue"], TextWriter.Null, TextWriter.Null, CancellationToken.None, described));
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

    [Theory]
    [InlineData("reader", PdfStandardPermissions.Print | PdfStandardPermissions.HighQualityPrint, true)] // printing allowed, copying denied
    [InlineData("wrong", PdfStandardPermissions.Print | PdfStandardPermissions.HighQualityPrint, false)]
    [InlineData("reader", PdfStandardPermissions.CopyContents, false)] // printing denied
    public async Task PlanAndPrintAdmitSourcePasswordsOnlyFromTheEnvironment(string password, PdfStandardPermissions permissions, bool admitted) {
        string path = Path.Combine(Path.GetTempPath(), "officeimo-print-locked-" + Guid.NewGuid().ToString("N") + ".pdf");
        string variable = "OFFICEIMO_PRINT_TEST_" + Guid.NewGuid().ToString("N");
        byte[] plain = PdfDocument.Create().Paragraph(paragraph => paragraph.AddText("Protected content")).ToBytes();
        File.WriteAllBytes(path, PdfDocument.Load(plain).Security.Encrypt(new("reader") { OwnerPassword = "owner", AllowedPermissions = permissions }).Pdf);
        Environment.SetEnvironmentVariable(variable, password);
        try {
            using var planOutput = new StringWriter();
            int planned = await WorkflowCommand.RunAsync(["print-plan", path, "--password-env", variable], planOutput, TextWriter.Null, CancellationToken.None);
            Assert.Equal(admitted, planned == (int)OfficeImoToolExitCode.Success);

            var service = new PrinterBoundary();
            int printed = await WorkflowCommand.RunAsync(["print", path, "--printer", "Test queue", "--paper", "A4", "--password-env", variable],
                TextWriter.Null, TextWriter.Null, CancellationToken.None, printerService: service);
            Assert.Equal(admitted, printed == (int)OfficeImoToolExitCode.Success);
            Assert.Equal(admitted, service.Document is not null);
            Assert.DoesNotContain(password, planOutput.ToString());
        } finally {
            Environment.SetEnvironmentVariable(variable, null);
            File.Delete(path);
        }
        Assert.Equal((int)OfficeImoToolExitCode.Usage, await WorkflowCommand.RunAsync(["print-plan", path, "--password-env", variable],
            TextWriter.Null, TextWriter.Null, CancellationToken.None));
    }

    [Fact]
    public async Task PrintConvertsOtherFormatsAndTracksJobs() {
        string path = Path.Combine(Path.GetTempPath(), "officeimo-print-cli-" + Guid.NewGuid().ToString("N") + ".md");
        File.WriteAllText(path, "# Heading\n\nPrinted from Markdown.");
        try {
            var service = new PrinterBoundary();
            using var output = new StringWriter();
            int code = await WorkflowPrintCommand.RunAsync(["print", path, "--printer", "ipp://localhost:8631/ipp/print", "--paper", "A4"],
                output, TextWriter.Null, CancellationToken.None, service);
            Assert.Equal((int)OfficeImoToolExitCode.Success, code);
            Assert.Equal("ipp://localhost:8631/ipp/print", service.Options!.PrinterName);
            Assert.NotEmpty(service.Document!.Plan.Sheets);

            using var status = new StringWriter();
            Assert.Equal((int)OfficeImoToolExitCode.Success, await WorkflowPrintCommand.RunAsync(
                ["print-job", "cancel", "--printer", "Test queue", "--job", "test-42"], status, TextWriter.Null, CancellationToken.None, service));
            using (JsonDocument canceled = JsonDocument.Parse(status.ToString())) {
                Assert.Equal("officeimo.workflow.print-job.v1", canceled.RootElement.GetProperty("schema").GetString());
                Assert.Equal("Canceled", canceled.RootElement.GetProperty("state").GetString());
                Assert.True(canceled.RootElement.GetProperty("isTerminal").GetBoolean());
            }
            status.GetStringBuilder().Clear();
            Assert.Equal((int)OfficeImoToolExitCode.Success, await WorkflowPrintCommand.RunAsync(
                ["print-job", "cancel", "--printer", "Test queue", "--job", "test-42", "--format", "text"], status, TextWriter.Null, CancellationToken.None, service));
            Assert.Equal("Job test-42 on Test queue: Canceled", status.ToString().Trim());
            await Assert.ThrowsAsync<WorkflowUsageException>(() => WorkflowPrintCommand.RunAsync(
                ["print-job", "status", "--printer", "Test queue"], TextWriter.Null, TextWriter.Null, CancellationToken.None, service));
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
