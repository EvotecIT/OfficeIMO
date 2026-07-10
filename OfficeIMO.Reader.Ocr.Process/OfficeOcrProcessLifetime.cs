using Microsoft.Win32.SafeHandles;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace OfficeIMO.Reader.Ocr.Process;

/// <summary>Owns the operating-system boundary used to terminate an OCR command and any descendants.</summary>
internal sealed class OfficeOcrProcessLifetime : IDisposable {
    private const int JobObjectExtendedLimitInformationClass = 9;
    private const uint JobObjectLimitKillOnJobClose = 0x00002000;
    private const int SigKill = 9;

    private readonly LifetimeMode _mode;
    private SafeFileHandle? _jobHandle;
    private int _processGroupId;
    private bool _disposed;

    private OfficeOcrProcessLifetime(LifetimeMode mode) {
        _mode = mode;
    }

    internal static OfficeOcrProcessLifetime Configure(
        ProcessStartInfo startInfo,
        string fileName,
        IReadOnlyList<string> arguments) {
        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) {
            return new OfficeOcrProcessLifetime(LifetimeMode.WindowsJob);
        }

        string? setSid = ResolveSetSid();
        if (setSid != null) {
            startInfo.FileName = setSid;
            startInfo.Arguments = JoinArguments(new[] { fileName }.Concat(arguments));
            return new OfficeOcrProcessLifetime(LifetimeMode.UnixProcessGroup);
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
            return new OfficeOcrProcessLifetime(LifetimeMode.UnixProcessGroup);
        }

        throw new PlatformNotSupportedException(
            "Bounded OCR process execution requires a Windows Job Object, /usr/bin/setsid or /bin/setsid on Unix, or /usr/bin/perl with POSIX::setsid on macOS.");
    }

    internal OfficeOcrStartedProcess Start(ProcessStartInfo startInfo) {
        if (_mode == LifetimeMode.WindowsJob) {
            return OfficeOcrWindowsSuspendedProcess.Start(startInfo, this);
        }

        var process = new System.Diagnostics.Process { StartInfo = startInfo, EnableRaisingEvents = true };
        try {
            if (!process.Start()) throw new InvalidOperationException("Failed to start OCR process '" + startInfo.FileName + "'.");
            _processGroupId = process.Id;
            return new OfficeOcrStartedProcess(process, process.StandardOutput, process.StandardError);
        } catch {
            process.Dispose();
            throw;
        }
    }

    internal void PrepareWindowsJob() {
        if (_mode != LifetimeMode.WindowsJob) throw new InvalidOperationException("A Windows Job Object is not configured for this process.");
        _jobHandle = TryCreateWindowsJob()
            ?? throw new InvalidOperationException("Unable to create the Windows Job Object required for bounded OCR process execution.");
    }

    internal bool AssignSuspendedWindowsProcess(IntPtr processHandle) {
        return _jobHandle != null
            && !_jobHandle.IsInvalid
            && !_jobHandle.IsClosed
            && AssignProcessToJobObject(_jobHandle.DangerousGetHandle(), processHandle);
    }

    internal void Terminate(System.Diagnostics.Process process) {
        if (_jobHandle != null && !_jobHandle.IsInvalid && !_jobHandle.IsClosed) {
            _ = TerminateJobObject(_jobHandle.DangerousGetHandle(), 1);
        }
        KillUnixProcessGroup();
        TryKillDirectProcess(process);
    }

    public void Dispose() {
        if (_disposed) return;
        _disposed = true;
        KillUnixProcessGroup();
        _jobHandle?.Dispose();
        _jobHandle = null;
    }

    private static string JoinArguments(IEnumerable<string> values) {
        return string.Join(" ", values.Select(OfficeOcrProcessRunner.QuoteArgument));
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

    private static SafeFileHandle? TryCreateWindowsJob() {
        IntPtr rawHandle = CreateJobObject(IntPtr.Zero, null);
        if (rawHandle == IntPtr.Zero || rawHandle == new IntPtr(-1)) return null;
        var handle = new SafeFileHandle(rawHandle, ownsHandle: true);
        var information = new JobObjectExtendedLimitInformation {
            BasicLimitInformation = new JobObjectBasicLimitInformation {
                LimitFlags = JobObjectLimitKillOnJobClose
            }
        };
        int size = Marshal.SizeOf<JobObjectExtendedLimitInformation>();
        IntPtr pointer = Marshal.AllocHGlobal(size);
        try {
            Marshal.StructureToPtr(information, pointer, fDeleteOld: false);
            if (!SetInformationJobObject(handle.DangerousGetHandle(), JobObjectExtendedLimitInformationClass, pointer, (uint)size)) {
                handle.Dispose();
                return null;
            }
            return handle;
        } finally {
            Marshal.FreeHGlobal(pointer);
        }
    }

    [DllImport("libc", EntryPoint = "kill", SetLastError = true)]
    private static extern int Kill(int pid, int signal);

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern IntPtr CreateJobObject(IntPtr jobAttributes, string? name);

    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool SetInformationJobObject(IntPtr job, int informationClass, IntPtr information, uint informationLength);

    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool AssignProcessToJobObject(IntPtr job, IntPtr process);

    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool TerminateJobObject(IntPtr job, uint exitCode);

    private enum LifetimeMode {
        UnixProcessGroup,
        WindowsJob
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct JobObjectBasicLimitInformation {
        internal long PerProcessUserTimeLimit;
        internal long PerJobUserTimeLimit;
        internal uint LimitFlags;
        internal UIntPtr MinimumWorkingSetSize;
        internal UIntPtr MaximumWorkingSetSize;
        internal uint ActiveProcessLimit;
        internal UIntPtr Affinity;
        internal uint PriorityClass;
        internal uint SchedulingClass;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct IoCounters {
        internal ulong ReadOperationCount;
        internal ulong WriteOperationCount;
        internal ulong OtherOperationCount;
        internal ulong ReadTransferCount;
        internal ulong WriteTransferCount;
        internal ulong OtherTransferCount;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct JobObjectExtendedLimitInformation {
        internal JobObjectBasicLimitInformation BasicLimitInformation;
        internal IoCounters IoInfo;
        internal UIntPtr ProcessMemoryLimit;
        internal UIntPtr JobMemoryLimit;
        internal UIntPtr PeakProcessMemoryUsed;
        internal UIntPtr PeakJobMemoryUsed;
    }
}
