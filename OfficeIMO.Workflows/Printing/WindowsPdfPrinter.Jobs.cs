using System.ComponentModel;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;

namespace OfficeIMO.Workflows;

#if NET8_0_OR_GREATER
[SupportedOSPlatform("windows")]
#endif
internal static partial class WindowsPdfPrinter {
    private const int ErrorInvalidParameter = 87, ErrorInsufficientBuffer = 122;
    private const uint JobControlDelete = 5;

    internal static Task<PdfPrintJobStatus> GetJobAsync(string printerName, string jobId, CancellationToken token) =>
        Task.Run(() => ReadJob(printerName, ParseJobId(jobId), token), token);

    /// <summary>Deletes the job from the spooler (JOB_CONTROL_DELETE; JOB_CONTROL_CANCEL is documented as not to be used).</summary>
    internal static Task<PdfPrintJobStatus> CancelJobAsync(string printerName, string jobId, CancellationToken token) => Task.Run(() => {
        uint id = ParseJobId(jobId);
        token.ThrowIfCancellationRequested();
        IntPtr printer = Open(printerName);
        try {
            if (!Native.SetJob(printer, id, 0, IntPtr.Zero, JobControlDelete)) {
                int error = Marshal.GetLastWin32Error();
                if (error != ErrorInvalidParameter) throw new Win32Exception(error);
            }
        } finally { _ = Native.ClosePrinter(printer); }
        return ReadJob(printerName, id, token);
    }, token);

    private static PdfPrintJobStatus ReadJob(string printerName, uint id, CancellationToken token) {
        token.ThrowIfCancellationRequested();
        string jobId = id.ToString(CultureInfo.InvariantCulture);
        IntPtr printer = Open(printerName);
        IntPtr buffer = IntPtr.Zero;
        try {
            if (!Native.GetJob(printer, id, 1, IntPtr.Zero, 0, out uint needed)) {
                int error = Marshal.GetLastWin32Error();
                // The spooler forgets a job once it has printed (unless the queue keeps printed documents) or was deleted.
                if (error == ErrorInvalidParameter) return new PdfPrintJobStatus(printerName, jobId, PdfPrintJobState.Removed);
                if (error != ErrorInsufficientBuffer) throw new Win32Exception(error);
            }
            if (needed < Marshal.SizeOf<Native.JobInfo1>() || needed > 1024 * 1024) throw new InvalidOperationException("The spooler returned invalid job information.");
            buffer = Marshal.AllocHGlobal((int)needed);
            if (!Native.GetJob(printer, id, 1, buffer, needed, out _)) {
                int error = Marshal.GetLastWin32Error();
                if (error == ErrorInvalidParameter) return new PdfPrintJobStatus(printerName, jobId, PdfPrintJobState.Removed);
                throw new Win32Exception(error);
            }
            var info = Marshal.PtrToStructure<Native.JobInfo1>(buffer);
            (PdfPrintJobState state, IReadOnlyList<string> reasons) = WindowsJobStatus.Map(info.Status);
            return new PdfPrintJobStatus(printerName, jobId, state) { StateReasons = reasons, StateMessage = Marshal.PtrToStringUni(info.StatusText) };
        } finally {
            if (buffer != IntPtr.Zero) Marshal.FreeHGlobal(buffer);
            _ = Native.ClosePrinter(printer);
        }
    }

    private static IntPtr Open(string printerName) {
        if (!Native.OpenPrinter(printerName, out IntPtr printer, IntPtr.Zero)) throw new Win32Exception();
        return printer;
    }

    private static uint ParseJobId(string jobId) =>
        uint.TryParse(jobId, NumberStyles.None, CultureInfo.InvariantCulture, out uint id) && id > 0
            ? id : throw new ArgumentException($"'{jobId}' is not a Windows spooler job identifier.", nameof(jobId));
}

/// <summary>Spooler job status mapping, kept free of Windows calls so it can be checked on every platform.</summary>
internal static class WindowsJobStatus {
    // JOB_INFO_1 status bits (winspool.h).
    private const uint Paused = 0x1, Error = 0x2, Deleting = 0x4, Spooling = 0x8, Printing = 0x10, Offline = 0x20, PaperOut = 0x40,
        Printed = 0x80, Deleted = 0x100, BlockedDevQueue = 0x200, UserIntervention = 0x400, Restart = 0x800, Complete = 0x1000,
        Retained = 0x2000, RenderingLocally = 0x4000;

    /// <summary>Maps JOB_INFO_1 status bits to the shared job states, keeping each set bit as a reason.</summary>
    internal static (PdfPrintJobState State, IReadOnlyList<string> Reasons) Map(uint status) {
        var reasons = new List<string>();
        void Reason(uint bit, string name) { if ((status & bit) != 0) reasons.Add(name); }
        Reason(Paused, "paused"); Reason(Error, "error"); Reason(Deleting, "deleting"); Reason(Spooling, "spooling");
        Reason(Printing, "printing"); Reason(Offline, "offline"); Reason(PaperOut, "paper-out"); Reason(Printed, "printed");
        Reason(Deleted, "deleted"); Reason(BlockedDevQueue, "blocked"); Reason(UserIntervention, "user-intervention");
        Reason(Restart, "restart"); Reason(Complete, "sent-to-printer"); Reason(Retained, "retained"); Reason(RenderingLocally, "rendering-locally");
        PdfPrintJobState state =
            (status & (Deleted | Deleting)) != 0 ? PdfPrintJobState.Canceled
            : (status & (Printed | Complete)) != 0 ? PdfPrintJobState.Completed
            : (status & (Error | Offline | PaperOut | UserIntervention | BlockedDevQueue)) != 0 ? PdfPrintJobState.Stopped
            : (status & Paused) != 0 ? PdfPrintJobState.Held
            : (status & (Printing | Spooling | RenderingLocally | Restart)) != 0 ? PdfPrintJobState.Processing
            : PdfPrintJobState.Pending;
        return (state, reasons);
    }
}
