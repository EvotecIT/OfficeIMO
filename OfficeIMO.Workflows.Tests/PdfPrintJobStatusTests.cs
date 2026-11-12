namespace OfficeIMO.Workflows.Tests;

/// <summary>Job states read from each transport map onto the same shared states.</summary>
public sealed class PdfPrintJobStatusTests {
    [Theory]
    [InlineData(0u, PdfPrintJobState.Pending)]
    [InlineData(0x8u, PdfPrintJobState.Processing)]                 // spooling
    [InlineData(0x10u | 0x4000u, PdfPrintJobState.Processing)]      // printing, rendering locally
    [InlineData(0x1u, PdfPrintJobState.Held)]                       // paused
    [InlineData(0x10u | 0x40u, PdfPrintJobState.Stopped)]           // printing but out of paper
    [InlineData(0x2u | 0x400u, PdfPrintJobState.Stopped)]           // error needing intervention
    [InlineData(0x80u, PdfPrintJobState.Completed)]                 // printed
    [InlineData(0x1000u | 0x2000u, PdfPrintJobState.Completed)]     // sent to the printer and retained
    [InlineData(0x4u | 0x10u, PdfPrintJobState.Canceled)]           // deleting while printing
    public void WindowsSpoolerStatusBitsMapToSharedStates(uint status, PdfPrintJobState expected) {
        (PdfPrintJobState state, IReadOnlyList<string> reasons) = WindowsJobStatus.Map(status);
        Assert.Equal(expected, state);
        Assert.Equal(System.Numerics.BitOperations.PopCount(status), reasons.Count);
    }

    [Theory]
    [InlineData(3, PdfPrintJobState.Pending, false)]
    [InlineData(4, PdfPrintJobState.Held, false)]
    [InlineData(5, PdfPrintJobState.Processing, false)]
    [InlineData(6, PdfPrintJobState.Stopped, false)]
    [InlineData(7, PdfPrintJobState.Canceled, true)]
    [InlineData(8, PdfPrintJobState.Aborted, true)]
    [InlineData(9, PdfPrintJobState.Completed, true)]
    public void IppJobStatesMapOneToOne(int ipp, PdfPrintJobState expected, bool terminal) {
        PdfPrintJobStatus status = IppPdfPrinter.MapJob("ipp://printer/ipp/print", "7",
            new IppJobInfo(7, (IppJobState)ipp) { StateReasons = ["job-printing"], StateMessage = "Printing" });
        Assert.Equal(expected, status.State);
        Assert.Equal(terminal, status.IsTerminal);
        Assert.Equal(["job-printing"], status.StateReasons);
        Assert.Equal("Printing", status.StateMessage);
    }

    [Fact]
    public async Task WaitingPassesTheDeadlineToEachQueryAndRejectsLateResults() {
        // A query that honours its token ends at the deadline.
        var stalled = new FakeService(async token => {
            await Task.Delay(Timeout.Infinite, token);
            throw new InvalidOperationException();
        });
        await Assert.ThrowsAsync<TimeoutException>(() => stalled.WaitForJobAsync("queue", "1", TimeSpan.FromMilliseconds(150)));

        // A terminal answer that arrives after the deadline is not accepted.
        var late = new FakeService(async _ => {
            await Task.Delay(300);
            return new PdfPrintJobStatus("queue", "1", PdfPrintJobState.Completed);
        });
        await Assert.ThrowsAsync<TimeoutException>(() => late.WaitForJobAsync("queue", "1", TimeSpan.FromMilliseconds(100)));

        // Caller cancellation stays distinct from the deadline.
        using var canceled = new CancellationTokenSource(TimeSpan.FromMilliseconds(100));
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => stalled.WaitForJobAsync("queue", "1", TimeSpan.FromMinutes(5), cancellationToken: canceled.Token));

        var pending = new FakeService(_ => Task.FromResult(new PdfPrintJobStatus("queue", "1", PdfPrintJobState.Processing)));
        TimeoutException timeout = await Assert.ThrowsAsync<TimeoutException>(() => pending.WaitForJobAsync("queue", "1", TimeSpan.FromMilliseconds(150), TimeSpan.FromMilliseconds(20)));
        Assert.Contains("Processing", timeout.Message);
    }

    private sealed class FakeService(Func<CancellationToken, Task<PdfPrintJobStatus>> query) : IPdfPrinterService {
        public Task<PdfPrintJobStatus> GetJobAsync(string printerName, string jobId, CancellationToken cancellationToken = default) => query(cancellationToken);
        public Task<IReadOnlyList<PdfPrinterInfo>> GetPrintersAsync(CancellationToken cancellationToken = default) => throw new NotSupportedException();
        public Task<PdfPrinterCapabilities> GetCapabilitiesAsync(string printerName, CancellationToken cancellationToken = default) => throw new NotSupportedException();
        public Task<PdfPrintSubmission> SubmitAsync(PdfPrintJob job, PdfPrintDeliveryOptions options, CancellationToken cancellationToken = default) => throw new NotSupportedException();
        public Task<PdfPrintJobStatus> CancelJobAsync(string printerName, string jobId, CancellationToken cancellationToken = default) => throw new NotSupportedException();
    }
}
