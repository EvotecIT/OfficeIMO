using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using Microsoft.Win32.SafeHandles;

namespace OfficeIMO.Word {
    /// <summary>Contains a native signing process and its descendants on supported Windows runtimes.</summary>
    internal sealed class WordMacroProjectProcessTree : IDisposable {
        private const uint JobObjectLimitKillOnJobClose = 0x00002000;
        private const int JobObjectBasicAccountingInformationClass = 1;
        private const int JobObjectExtendedLimitInformationClass = 9;
        private readonly SafeFileHandle _job;

        private WordMacroProjectProcessTree(SafeFileHandle job) => _job = job;

        internal static bool TryAttach(
            Process process,
            out WordMacroProjectProcessTree? processTree,
            out string detail) {
            processTree = null;
            detail = string.Empty;
            if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) return true;

            SafeFileHandle job = CreateJobObject(IntPtr.Zero, null);
            if (job.IsInvalid) {
                detail = new Win32Exception(Marshal.GetLastWin32Error()).Message;
                job.Dispose();
                return false;
            }
            try {
                var limits = new JobObjectExtendedLimitInformation {
                    BasicLimitInformation = new JobObjectBasicLimitInformation {
                        LimitFlags = JobObjectLimitKillOnJobClose
                    }
                };
                int size = Marshal.SizeOf(typeof(JobObjectExtendedLimitInformation));
                if (!SetInformationJobObject(job, JobObjectExtendedLimitInformationClass, ref limits, size) ||
                    !AssignProcessToJobObject(job, process.Handle)) {
                    detail = new Win32Exception(Marshal.GetLastWin32Error()).Message;
                    return false;
                }
                processTree = new WordMacroProjectProcessTree(job);
                job = null!;
                return true;
            } finally {
                job?.Dispose();
            }
        }

        internal bool TerminateAndWait(TimeSpan timeout) {
            if (!TerminateJobObject(_job, 1)) return false;
            var stopwatch = Stopwatch.StartNew();
            do {
                if (TryGetActiveProcessCount(out uint activeProcesses) && activeProcesses == 0) return true;
                Thread.Sleep(10);
            } while (stopwatch.Elapsed < timeout);
            return TryGetActiveProcessCount(out uint remainingProcesses) && remainingProcesses == 0;
        }

        private bool TryGetActiveProcessCount(out uint activeProcesses) {
            int size = Marshal.SizeOf(typeof(JobObjectBasicAccountingInformation));
            bool succeeded = QueryInformationJobObject(
                _job,
                JobObjectBasicAccountingInformationClass,
                out JobObjectBasicAccountingInformation information,
                size,
                IntPtr.Zero);
            activeProcesses = succeeded ? information.ActiveProcesses : uint.MaxValue;
            return succeeded;
        }

        public void Dispose() => _job.Dispose();

        [StructLayout(LayoutKind.Sequential)]
        private struct JobObjectBasicAccountingInformation {
            internal long TotalUserTime;
            internal long TotalKernelTime;
            internal long ThisPeriodTotalUserTime;
            internal long ThisPeriodTotalKernelTime;
            internal uint TotalPageFaultCount;
            internal uint TotalProcesses;
            internal uint ActiveProcesses;
            internal uint TotalTerminatedProcesses;
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

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern SafeFileHandle CreateJobObject(IntPtr securityAttributes, string? name);

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool SetInformationJobObject(
            SafeFileHandle job,
            int informationClass,
            ref JobObjectExtendedLimitInformation information,
            int informationLength);

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool AssignProcessToJobObject(SafeFileHandle job, IntPtr process);

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool TerminateJobObject(SafeFileHandle job, uint exitCode);

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool QueryInformationJobObject(
            SafeFileHandle job,
            int informationClass,
            out JobObjectBasicAccountingInformation information,
            int informationLength,
            IntPtr returnLength);
    }
}
