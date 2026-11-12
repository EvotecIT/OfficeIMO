using OfficeIMO.Pdf;

namespace OfficeIMO.Workflows.Tests;

public sealed class PdfCupsSubmissionTests {
    [Fact]
    public async Task AcceptedAcknowledgementReturnsReceiptAndCleansStaging() {
        string? path = null;
        PdfPrintSubmission receipt = await Submit(new PdfPrintDeliveryOptions { PrinterName = "test-queue", Copies = 2 }, [1], default,
            (_, arguments, _, started) => {
                path = arguments[^1];
                started!();
                return Task.FromResult((0, "request id is test-queue-42 (1 file(s))", string.Empty));
            });
        Assert.Equal("test-queue-42", receipt.JobId);
        Assert.Equal(2, receipt.Copies);
        Assert.Null(receipt.CleanupWarning);
        Assert.False(Directory.Exists(Path.GetDirectoryName(path!)));
    }

    [Fact]
    public async Task ProcessStartFailureIsDefinitiveAndCleansStaging() {
        string? path = null;
        await Assert.ThrowsAsync<IOException>(() => Submit(new PdfPrintDeliveryOptions { PrinterName = "test-queue" }, [1], default,
            (_, arguments, _, _) => {
                path = arguments[^1];
                throw new IOException("Executable unavailable");
            }));
        Assert.False(Directory.Exists(Path.GetDirectoryName(path!)));
    }

    [Theory]
    [InlineData(1, "/usr/bin/lp: Error - The printer or class does not exist.", false)]
    [InlineData(1, "lp: Error - The printer or class does not exist.", false)]
    [InlineData(1, "lp: The printer or class does not exist.", true)]
    [InlineData(1, "lp: Unable to finish document: Connection reset by peer", true)]
    [InlineData(1, "Unsupported media", true)]
    [InlineData(137, "/usr/bin/lp: Error - The printer or class does not exist.", true)]
    public async Task CommandExitDistinguishesProvenPreSubmissionRejection(int exitCode, string diagnostic, bool uncertain) {
        string? path = null;
        Exception? error = await Record.ExceptionAsync(() => Submit(new PdfPrintDeliveryOptions { PrinterName = "test-queue" }, [1, 2, 3], default,
            (command, arguments, token, started) => {
                Assert.Equal("lp", command);
                path = arguments[^1];
                Assert.Equal(new byte[] { 1, 2, 3 }, File.ReadAllBytes(path));
                started!();
                return Task.FromResult((exitCode, string.Empty, diagnostic));
            }));
        if (uncertain) Assert.IsType<PdfPrintDeliveryException>(error);
        else Assert.IsType<IOException>(error);
        Assert.False(Directory.Exists(Path.GetDirectoryName(path!)));
    }

    [Fact]
    public async Task LostAcknowledgementRemainsUncertainAndCleansStaging() {
        string? path = null;
        await Assert.ThrowsAsync<PdfPrintDeliveryException>(() => Submit(new PdfPrintDeliveryOptions { PrinterName = "test-queue" }, [1], default,
            (_, arguments, _, started) => {
                path = arguments[^1];
                started!();
                return Task.FromResult((0, string.Empty, string.Empty));
            }));
        Assert.False(Directory.Exists(Path.GetDirectoryName(path!)));
    }

    [Fact]
    public async Task CancellationAfterProcessStartRemainsUncertainAndCleansStaging() {
        using var cancellation = new CancellationTokenSource();
        string? path = null;
        PdfPrintDeliveryException error = await Assert.ThrowsAsync<PdfPrintDeliveryException>(() =>
            Submit(new PdfPrintDeliveryOptions { PrinterName = "test-queue" }, [1],
                cancellation.Token, (_, arguments, token, started) => {
                    path = arguments[^1];
                    started!();
                    cancellation.Cancel();
                    return Task.FromCanceled<(int, string, string)>(token);
                }));
        Assert.IsAssignableFrom<OperationCanceledException>(error.InnerException);
        Assert.False(Directory.Exists(Path.GetDirectoryName(path!)));
    }

    [Theory]
    [InlineData("request id is test-queue-42 (1 file(s))", "test-queue-42")]
    [InlineData("ID żądania to test-queue-42 (1 plików)", "test-queue-42")] // macOS localizes lp regardless of LC_ALL.
    [InlineData("request id is other-queue-42 (1 file(s))", null)]
    [InlineData("request id is test-queue-42x", null)]
    public void JobAcknowledgementIsRecognizedByQueueNameInAnyLanguage(string output, string? expected) =>
        Assert.Equal(expected, CupsPdfPrinter.ParseJobId(output, "test-queue"));

    [Theory]
    [InlineData("system default destination: Office", "Office")]
    [InlineData("domyślny cel systemowy: Office", "Office")]
    [InlineData("brak domyślnego celu systemowego", null)]
    public void DefaultQueueIsRecognizedInLocalizedOutput(string output, string? expected) {
        IReadOnlyList<PdfPrinterInfo> printers = CupsPdfPrinter.ParsePrinters("Archive\nOffice\n", output);
        Assert.Equal(expected, printers.SingleOrDefault(printer => printer.IsDefault)?.Name);
        Assert.Equal(2, printers.Count);
    }

    [Fact]
    public void MixedOrientationKeepsOneMediaAndMixedSizesLetEachPageChooseItsMedia() {
        PdfPrintPlan mixedOrientation = Plan(PdfPrintOrientation.Automatic, page => page.Size(OfficeIMO.OfficeLength.Points(200), OfficeIMO.OfficeLength.Points(300)),
            page => page.Size(OfficeIMO.OfficeLength.Points(300), OfficeIMO.OfficeLength.Points(200)), PdfPageSizes.A4);
        List<string> arguments = CupsPdfPrinter.CreateArguments(mixedOrientation, new PdfPrintDeliveryOptions { PrinterName = "queue" }, [], null);
        Assert.Contains("media=A4", arguments);
        Assert.DoesNotContain(arguments, argument => argument.StartsWith("orientation-requested=", StringComparison.Ordinal));

        var mixedSizes = new PdfPrintPlan(2, [1, 2], [new PdfPrintSheet(1, PdfPageSizes.A4, []), new PdfPrintSheet(2, PdfPageSizes.A3, [])]);
        arguments = CupsPdfPrinter.CreateArguments(mixedSizes, new PdfPrintDeliveryOptions { PrinterName = "queue" }, [], null);
        Assert.DoesNotContain(arguments, argument => argument.StartsWith("media=", StringComparison.Ordinal));
        Assert.Contains("print-scaling=none", arguments);
    }

    [Fact]
    public void UniformJobsUseTheQueueMediaKeywordForTheirPaper() {
        PdfPrintPlan plan = Plan(PdfPrintOrientation.Landscape, page => page.Size(OfficeIMO.OfficeLength.Points(200), OfficeIMO.OfficeLength.Points(300)),
            page => page.Size(OfficeIMO.OfficeLength.Points(300), OfficeIMO.OfficeLength.Points(200)), PdfPageSizes.Letter);
        PdfPrinterCapabilities capabilities = CupsPdfPrinter.ParseCapabilities("queue",
            "media/Media Size: iso_a4_210x297mm *na_letter_8.5x11in na_legal_8.5x14in\n", string.Empty);
        List<string> arguments = CupsPdfPrinter.CreateArguments(plan, new PdfPrintDeliveryOptions { PrinterName = "queue" }, capabilities.Media, capabilities.DefaultMedia);
        Assert.Contains("media=na_letter_8.5x11in", arguments);
        Assert.Contains("orientation-requested=4", arguments);
        PdfPrintPlan custom = Plan(PdfPrintOrientation.Portrait, page => page.Size(OfficeIMO.OfficeLength.Points(200), OfficeIMO.OfficeLength.Points(300)),
            page => page.Size(OfficeIMO.OfficeLength.Points(200), OfficeIMO.OfficeLength.Points(300)), new PdfPageSize(OfficeIMO.OfficeLength.Points(250), OfficeIMO.OfficeLength.Points(400)));
        Assert.Contains("media=Custom.250x400pt", CupsPdfPrinter.CreateArguments(custom, new PdfPrintDeliveryOptions { PrinterName = "queue" }, capabilities.Media, capabilities.DefaultMedia));
    }

    private static PdfPrintPlan Plan(PdfPrintOrientation orientation, Action<PdfPageBuilder> first, Action<PdfPageBuilder> second, PdfPageSize paper) =>
        PdfPrintPlanner.Create(PdfDocument.Create(document => { document.Page(first); document.Page(second); }),
            new PdfPrintPlanRequest { InputPath = "snapshot.pdf", PaperSize = paper, Orientation = orientation });

    private static Task<PdfPrintSubmission> Submit(PdfPrintDeliveryOptions options, byte[] content, CancellationToken token, CupsPdfPrinter.CupsCommandRunner run) {
        PdfPrintPlan plan = PdfPrintPlanner.Create(PdfDocument.Create(document => document.Page(page => page.Size(OfficeIMO.OfficeLength.Points(100), OfficeIMO.OfficeLength.Points(100)).Margin(OfficeIMO.OfficeLength.Points(0)))),
            new PdfPrintPlanRequest { InputPath = "snapshot.pdf", PaperSize = new PdfPageSize(OfficeIMO.OfficeLength.Points(100), OfficeIMO.OfficeLength.Points(100)) });
        return CupsPdfPrinter.SubmitSpoolAsync(plan, options, CupsPdfPrinter.CreateArguments(plan, options, [], null), new MemoryStream(content), token, run);
    }
}
