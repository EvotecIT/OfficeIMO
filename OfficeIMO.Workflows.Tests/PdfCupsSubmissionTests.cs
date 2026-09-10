using OfficeIMO.Pdf;

namespace OfficeIMO.Workflows.Tests;

public sealed class PdfCupsSubmissionTests {
    [Fact]
    public async Task AcceptedAcknowledgementReturnsReceiptAndCleansStaging() {
        string? path = null;
        PdfPrintSubmission receipt = await CupsPdfPrinter.SubmitSpoolAsync(Prepare(),
            new PdfPrintDeliveryOptions { PrinterName = "test-queue", Copies = 2 }, [1], default,
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
        await Assert.ThrowsAsync<IOException>(() => CupsPdfPrinter.SubmitSpoolAsync(Prepare(),
            new PdfPrintDeliveryOptions { PrinterName = "test-queue" }, [1], default,
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
        PdfPreparedPrintDocument document = Prepare();
        string? path = null;
        Exception? error = await Record.ExceptionAsync(() => CupsPdfPrinter.SubmitSpoolAsync(document,
            new PdfPrintDeliveryOptions { PrinterName = "test-queue" }, [1, 2, 3], default,
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
        await Assert.ThrowsAsync<PdfPrintDeliveryException>(() => CupsPdfPrinter.SubmitSpoolAsync(Prepare(),
            new PdfPrintDeliveryOptions { PrinterName = "test-queue" }, [1], default,
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
            CupsPdfPrinter.SubmitSpoolAsync(Prepare(), new PdfPrintDeliveryOptions { PrinterName = "test-queue" }, [1],
                cancellation.Token, (_, arguments, token, started) => {
                    path = arguments[^1];
                    started!();
                    cancellation.Cancel();
                    return Task.FromCanceled<(int, string, string)>(token);
                }));
        Assert.IsAssignableFrom<OperationCanceledException>(error.InnerException);
        Assert.False(Directory.Exists(Path.GetDirectoryName(path!)));
    }

    [Fact]
    public void CancellationDuringSpoolSerializationStopsComposition() {
        using var cancellation = new CancellationTokenSource();
        int callbacks = 0;
        PdfDocument document = PdfDocument.Create(new PdfOptions {
            TextLineBreakCallback = text => {
                callbacks++;
                cancellation.Cancel();
                return [text.Length / 2];
            }
        });
        for (int index = 0; index < 30; index++) document.Paragraph(p => p.Text(new string('W', 600)));
        Assert.ThrowsAny<OperationCanceledException>(() => CupsPdfPrinter.SerializeSpool([document], cancellation.Token));
        Assert.InRange(callbacks, 1, 2);
    }

    private static PdfPreparedPrintDocument Prepare() => PdfPrintRenderer.Prepare(
        PdfDocument.Create(document => document.Page(page => page.Size(100, 100).Margin(0))),
        new PdfPrintPlanRequest { InputPath = "snapshot.pdf", PaperSize = new PageSize(100, 100) },
        new PdfPrintRenderOptions { Dpi = 72 });
}
