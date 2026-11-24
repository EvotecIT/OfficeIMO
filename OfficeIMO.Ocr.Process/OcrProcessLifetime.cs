using System.Diagnostics;
using System.Runtime.InteropServices;
using OfficeIMO.Processes.Internal;

namespace OfficeIMO.Ocr.Process;

/// <summary>Owns the operating-system boundary used to terminate an OCR command and any descendants.</summary>
internal sealed class OcrProcessLifetime : IDisposable {
    private const int SigKill = 9;

    private readonly LifetimeMode _mode;
    private OfficeWindowsProcessJob? _job;
    private int _processGroupId;
    private bool _disposed;

    private OcrProcessLifetime(LifetimeMode mode) {
        _mode = mode;
    }

    internal static OcrProcessLifetime Configure(
        ProcessStartInfo startInfo,
        string fileName,
        IReadOnlyList<string> arguments) {
        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) {
            return new OcrProcessLifetime(LifetimeMode.WindowsJob);
        }

        string? setSid = ResolveSetSid();
        if (setSid != null) {
            startInfo.FileName = setSid;
            startInfo.Arguments = JoinArguments(new[] { fileName }.Concat(arguments));
            return new OcrProcessLifetime(LifetimeMode.UnixProcessGroup);
        }

        const string macPerl = "/usr/bin/perl";
        if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX) && File.Exists(macPerl)) {
            startInfo.FileName = macPerl;
            startInfo.Arguments = JoinArguments(new[] {
                "-MPOSIX",
                "-e",
                "defined(POSIX::setsid()) or die qq(setsid failed: $!\\n); exec { $ARGV[0] } @ARGV; die qq(exec failed: $!\\n);",
                fileName
            }.Concat(arguments));
            return new OcrProcessLifetime(LifetimeMode.UnixProcessGroup);
        }

        throw new PlatformNotSupportedException(
            "Bounded OCR process execution requires a Windows Job Object, /usr/bin/setsid or /bin/setsid on Unix, or /usr/bin/perl with POSIX::setsid on macOS.");
    }

    internal OcrStartedProcess Start(ProcessStartInfo startInfo) {
        if (_mode == LifetimeMode.WindowsJob) {
            OfficeWindowsSuspendedProcessResult started = OfficeWindowsSuspendedProcess.Start(
                startInfo,
                PrepareWindowsJob,
                AssignSuspendedWindowsProcess);
            return new OcrStartedProcess(started.Process, started.StandardOutput, started.StandardError);
        }

        var process = new System.Diagnostics.Process { StartInfo = startInfo, EnableRaisingEvents = true };
        try {
            if (!process.Start()) throw new InvalidOperationException("Failed to start OCR process '" + startInfo.FileName + "'.");
            _processGroupId = process.Id;
            return new OcrStartedProcess(process, process.StandardOutput, process.StandardError);
        } catch {
            process.Dispose();
            throw;
        }
    }

    internal void PrepareWindowsJob() {
        if (_mode != LifetimeMode.WindowsJob) throw new InvalidOperationException("A Windows Job Object is not configured for this process.");
        _job = OfficeWindowsProcessJob.TryCreateKillOnClose()
            ?? throw new InvalidOperationException("Unable to create the Windows Job Object required for bounded OCR process execution.");
    }

    internal bool AssignSuspendedWindowsProcess(IntPtr processHandle) => _job?.TryAssign(processHandle) == true;

    internal void Terminate(System.Diagnostics.Process process) {
        _job?.Terminate();
        KillUnixProcessGroup();
        TryKillDirectProcess(process);
    }

    public void Dispose() {
        if (_disposed) return;
        _disposed = true;
        KillUnixProcessGroup();
        _job?.Dispose();
        _job = null;
    }

    private static string JoinArguments(IEnumerable<string> values) {
        return string.Join(" ", values.Select(OcrProcessRunner.QuoteArgument));
    }

    private static string? ResolveSetSid() {
        string[] candidates = { "/usr/bin/setsid", "/bin/setsid" };
        return candidates.FirstOrDefault(File.Exists);
    }

    private void KillUnixProcessGroup() {
        int processGroupId = Interlocked.Exchange(ref _processGroupId, 0);
        if (processGroupId <= 0) return;
        try {
            _ = Kill(-processGroupId, SigKill);
        } catch (DllNotFoundException) {
        } catch (EntryPointNotFoundException) {
        }
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

    [DllImport("libc", EntryPoint = "kill", SetLastError = true)]
    private static extern int Kill(int pid, int signal);

    private enum LifetimeMode {
        UnixProcessGroup,
        WindowsJob
    }
}
