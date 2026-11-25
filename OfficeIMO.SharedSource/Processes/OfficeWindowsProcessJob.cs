using Microsoft.Win32.SafeHandles;
using System.Runtime.InteropServices;

namespace OfficeIMO.Processes.Internal;

/// <summary>Owns a Windows Job Object that terminates its member processes when its last handle closes.</summary>
internal sealed class OfficeWindowsProcessJob : IDisposable {
    private const int ExtendedLimitInformationClass = 9;
    private const uint LimitKillOnJobClose = 0x00002000;
    private readonly SafeFileHandle _handle;

    private OfficeWindowsProcessJob(SafeFileHandle handle) {
        _handle = handle;
    }

    /// <summary>Creates an unnamed job and enables process-tree cleanup before any process joins it.</summary>
    internal static OfficeWindowsProcessJob? TryCreateKillOnClose() {
        if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) return null;

        SafeFileHandle? handle = CreateJobObject(IntPtr.Zero, null);
        try {
            if (handle.IsInvalid) return null;
            var information = new JobObjectExtendedLimitInformation {
                BasicLimitInformation = new JobObjectBasicLimitInformation {
                    LimitFlags = LimitKillOnJobClose
                }
            };
            uint size = (uint)Marshal.SizeOf<JobObjectExtendedLimitInformation>();
            if (!SetInformationJobObject(handle, ExtendedLimitInformationClass, ref information, size)) return null;

            var job = new OfficeWindowsProcessJob(handle);
            handle = null;
            return job;
        } finally {
            handle?.Dispose();
        }
    }

    /// <summary>Assigns a process to the job. Assign it before the process can launch descendants.</summary>
    internal bool TryAssign(IntPtr processHandle) =>
        !_handle.IsClosed && AssignProcessToJobObject(_handle, processHandle);

    /// <summary>Requests immediate termination of the job's member processes.</summary>
    internal void Terminate(uint exitCode = 1) {
        if (!_handle.IsClosed) {
            _ = TerminateJobObject(_handle, exitCode);
        }
    }

    /// <summary>Closes the job handle and terminates any remaining member processes.</summary>
    public void Dispose() => _handle.Dispose();

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern SafeFileHandle CreateJobObject(IntPtr jobAttributes, string? name);

    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool SetInformationJobObject(SafeFileHandle job, int informationClass,
        ref JobObjectExtendedLimitInformation information, uint informationLength);

    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool AssignProcessToJobObject(SafeFileHandle job, IntPtr process);

    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool TerminateJobObject(SafeFileHandle job, uint exitCode);

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
