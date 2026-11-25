using System.Diagnostics;
using System.Runtime.InteropServices;
using OfficeIMO.Processes.Internal;

namespace OfficeIMO.Ocr.Process;

/// <summary>Owns the operating-system boundary used to terminate an OCR command and any descendants.</summary>
internal sealed class OcrProcessLifetime : IDisposable {
    private readonly OfficeUnixProcessGroup? _unixGroup;
    private OfficeWindowsProcessJob? _job;
    private bool _disposed;

    private OcrProcessLifetime(OfficeUnixProcessGroup? unixGroup) {
        _unixGroup = unixGroup;
    }

    internal static OcrProcessLifetime Configure(ProcessStartInfo startInfo) {
        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) {
            return new OcrProcessLifetime(null);
        }

        OfficeUnixProcessGroup? unixGroup = OfficeUnixProcessGroup.TryConfigure(startInfo);
        if (unixGroup != null) return new OcrProcessLifetime(unixGroup);

        throw new PlatformNotSupportedException(
            "Bounded OCR process execution requires a Windows Job Object, /usr/bin/setsid or /bin/setsid on Unix, or /usr/bin/perl with POSIX::setsid on macOS.");
    }

    internal OcrStartedProcess Start(ProcessStartInfo startInfo) {
        if (_unixGroup == null) {
            OfficeWindowsSuspendedProcessResult started = OfficeWindowsSuspendedProcess.Start(
                startInfo,
                PrepareWindowsJob,
                AssignSuspendedWindowsProcess);
            return new OcrStartedProcess(started.Process, started.StandardOutput, started.StandardError);
        }

        var process = new System.Diagnostics.Process { StartInfo = startInfo, EnableRaisingEvents = true };
        try {
            if (!process.Start()) throw new InvalidOperationException("Failed to start OCR process '" + startInfo.FileName + "'.");
            _unixGroup.TrackLeader(process.Id);
            return new OcrStartedProcess(process, process.StandardOutput, process.StandardError);
        } catch {
            process.Dispose();
            throw;
        }
    }

    internal void PrepareWindowsJob() {
        if (_unixGroup != null) throw new InvalidOperationException("A Windows Job Object is not configured for this process.");
        _job = OfficeWindowsProcessJob.TryCreateKillOnClose()
            ?? throw new InvalidOperationException("Unable to create the Windows Job Object required for bounded OCR process execution.");
    }

    internal bool AssignSuspendedWindowsProcess(IntPtr processHandle) => _job?.TryAssign(processHandle) == true;

    internal void Terminate(System.Diagnostics.Process process) {
        _job?.Terminate();
        _unixGroup?.TryTerminate();
        TryKillDirectProcess(process);
    }

    public void Dispose() {
        if (_disposed) return;
        _disposed = true;
        _unixGroup?.TryTerminate();
        _job?.Dispose();
        _job = null;
    }

    private static void TryKillDirectProcess(System.Diagnostics.Process process) {
        try {
            if (!process.HasExited) {
#if NET8_0_OR_GREATER
                process.Kill(entireProcessTree: true);
#else
                process.Kill();
#endif
            }
        } catch (InvalidOperationException) {
        } catch (System.ComponentModel.Win32Exception) {
        } catch (NotSupportedException) {
        }
    }
}
