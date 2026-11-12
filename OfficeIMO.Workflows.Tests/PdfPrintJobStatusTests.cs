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
}
