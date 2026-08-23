using System.Diagnostics;
using System.Runtime.InteropServices;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

internal sealed class ProcessTreeMemorySampler : IAsyncDisposable {
    private readonly int _rootProcessId;
    private readonly CancellationTokenSource _stop = new();
    private readonly HashSet<int> _knownProcessIds = new();
    private readonly Task _samplingTask;
    private int _stopped;
    private long _peakWorkingSetBytes;
    private int _sampleCount;
    private int _minimumProcessCount = int.MaxValue;
    private int _maximumProcessCount;

    internal ProcessTreeMemorySampler(Process rootProcess) {
        ArgumentNullException.ThrowIfNull(rootProcess);
        _rootProcessId = rootProcess.Id;
        Sample();
        _samplingTask = Task.Run(SampleAsync);
    }

    internal async Task StopAsync() {
        if (Interlocked.Exchange(ref _stopped, 1) != 0) return;
        _stop.Cancel();
        try {
            await _samplingTask.ConfigureAwait(false);
        } catch (OperationCanceledException) {
            // Normal sampler shutdown.
        }
        Sample();
    }

    internal HtmlPdfProcessTreeMemoryEvidence CreateEvidence() {
        if (Volatile.Read(ref _stopped) == 0) {
            throw new InvalidOperationException("Process-tree memory evidence cannot be read before the sampler is stopped.");
        }
        return new HtmlPdfProcessTreeMemoryEvidence(
            PeakWorkingSetBytes: _peakWorkingSetBytes,
            SampleCount: _sampleCount,
            MinimumObservedProcessCount: _minimumProcessCount == int.MaxValue ? 0 : _minimumProcessCount,
            MaximumObservedProcessCount: _maximumProcessCount,
            Sampler: ProcessTreeSnapshot.SamplerIdentity);
    }

    public async ValueTask DisposeAsync() {
        await StopAsync().ConfigureAwait(false);
        _stop.Dispose();
    }

    private async Task SampleAsync() {
        while (!_stop.IsCancellationRequested) {
            await Task.Delay(TimeSpan.FromMilliseconds(10), _stop.Token).ConfigureAwait(false);
            Sample();
        }
    }

    private void Sample() {
        ProcessTreeSample sample = ProcessTreeSnapshot.Capture(_rootProcessId, _knownProcessIds);
        if (sample.ProcessCount <= 0) return;
        _sampleCount++;
        _peakWorkingSetBytes = Math.Max(_peakWorkingSetBytes, sample.WorkingSetBytes);
        _minimumProcessCount = Math.Min(_minimumProcessCount, sample.ProcessCount);
        _maximumProcessCount = Math.Max(_maximumProcessCount, sample.ProcessCount);
    }
}

internal readonly record struct ProcessTreeSample(long WorkingSetBytes, int ProcessCount);

internal static class ProcessTreeSnapshot {
    internal static string SamplerIdentity => OperatingSystem.IsWindows()
        ? "Toolhelp process tree + Process.WorkingSet64"
        : OperatingSystem.IsLinux()
            ? "/proc parent map + Process.WorkingSet64"
            : "ps pid/ppid/rss snapshot";

    internal static ProcessTreeSample Capture(int rootProcessId, HashSet<int> knownProcessIds) {
        IReadOnlyList<ProcessMemoryInfo> processes = OperatingSystem.IsWindows()
            ? CaptureWindows()
            : OperatingSystem.IsLinux()
                ? CaptureLinux()
                : CapturePs();
        if (processes.Count == 0) return default;

        var byParent = new Dictionary<int, List<int>>();
        var byProcessId = new Dictionary<int, ProcessMemoryInfo>();
        foreach (ProcessMemoryInfo process in processes) {
            byProcessId[process.ProcessId] = process;
            if (!byParent.TryGetValue(process.ParentProcessId, out List<int>? children)) {
                children = new List<int>();
                byParent.Add(process.ParentProcessId, children);
            }
            children.Add(process.ProcessId);
        }

        var currentProcessIds = new HashSet<int>();
        var pending = new Stack<int>();
        pending.Push(rootProcessId);
        while (pending.Count > 0) {
            int processId = pending.Pop();
            if (!currentProcessIds.Add(processId)) continue;
            if (!byParent.TryGetValue(processId, out List<int>? children)) continue;
            foreach (int child in children) pending.Push(child);
        }

        knownProcessIds.UnionWith(currentProcessIds);
        knownProcessIds.RemoveWhere(processId => !byProcessId.ContainsKey(processId));
        long total = 0L;
        int observedProcessCount = 0;
        foreach (int processId in knownProcessIds.ToArray()) {
            ProcessMemoryInfo process = byProcessId[processId];
            long? workingSet = process.WorkingSetBytes >= 0
                ? process.WorkingSetBytes
                : ReadWorkingSet(processId);
            if (workingSet == null) {
                knownProcessIds.Remove(processId);
                continue;
            }
            total += workingSet.Value;
            observedProcessCount++;
        }
        return new ProcessTreeSample(total, observedProcessCount);
    }

    private static IReadOnlyList<ProcessMemoryInfo> CaptureWindows() {
        var result = new List<ProcessMemoryInfo>();
        IntPtr snapshot = CreateToolhelp32Snapshot(0x00000002, 0);
        if (snapshot == new IntPtr(-1)) return result;
        try {
            var entry = new ProcessEntry32 { Size = (uint)Marshal.SizeOf<ProcessEntry32>() };
            if (!Process32First(snapshot, ref entry)) return result;
            do {
                result.Add(new ProcessMemoryInfo(
                    checked((int)entry.ProcessId),
                    checked((int)entry.ParentProcessId),
                    -1L));
                entry.Size = (uint)Marshal.SizeOf<ProcessEntry32>();
            } while (Process32Next(snapshot, ref entry));
        } finally {
            _ = CloseHandle(snapshot);
        }
        return result;
    }

    private static long? ReadWorkingSet(int processId) {
        try {
            using Process process = Process.GetProcessById(processId);
            process.Refresh();
            return process.WorkingSet64;
        } catch (Exception exception) when (exception is ArgumentException or InvalidOperationException or System.ComponentModel.Win32Exception or NotSupportedException) {
            return null;
        }
    }

    private static IReadOnlyList<ProcessMemoryInfo> CaptureLinux() {
        var result = new List<ProcessMemoryInfo>();
        foreach (string directory in Directory.EnumerateDirectories("/proc")) {
            string name = Path.GetFileName(directory);
            if (!int.TryParse(name, out int processId)) continue;
            try {
                string stat = File.ReadAllText(Path.Combine(directory, "stat"));
                int commandEnd = stat.LastIndexOf(')');
                if (commandEnd < 0 || commandEnd + 2 >= stat.Length) continue;
                string[] fields = stat[(commandEnd + 2)..].Split(' ', StringSplitOptions.RemoveEmptyEntries);
                if (fields.Length < 2 || !int.TryParse(fields[1], out int parentProcessId)) continue;
                result.Add(new ProcessMemoryInfo(processId, parentProcessId, -1L));
            } catch (Exception exception) when (exception is IOException or UnauthorizedAccessException) {
                // A process can exit or deny inspection between enumeration and sampling.
            }
        }
        return result;
    }

    private static IReadOnlyList<ProcessMemoryInfo> CapturePs() {
        var startInfo = new ProcessStartInfo {
            FileName = "ps",
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true
        };
        startInfo.ArgumentList.Add("-axo");
        startInfo.ArgumentList.Add("pid=,ppid=,rss=");
        using Process process = Process.Start(startInfo)
            ?? throw new InvalidOperationException("Could not start ps for process-tree memory sampling.");
        string output = process.StandardOutput.ReadToEnd();
        process.WaitForExit();
        if (process.ExitCode != 0) return Array.Empty<ProcessMemoryInfo>();

        var result = new List<ProcessMemoryInfo>();
        foreach (string line in output.Split('\n', StringSplitOptions.RemoveEmptyEntries)) {
            string[] fields = line.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
            if (fields.Length != 3 ||
                !int.TryParse(fields[0], out int processId) ||
                !int.TryParse(fields[1], out int parentProcessId) ||
                !long.TryParse(fields[2], out long residentKilobytes)) continue;
            result.Add(new ProcessMemoryInfo(processId, parentProcessId, checked(residentKilobytes * 1024L)));
        }
        return result;
    }

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern IntPtr CreateToolhelp32Snapshot(uint flags, uint processId);

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool Process32First(IntPtr snapshot, ref ProcessEntry32 entry);

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool Process32Next(IntPtr snapshot, ref ProcessEntry32 entry);

    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool CloseHandle(IntPtr handle);

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct ProcessEntry32 {
        internal uint Size;
        internal uint Usage;
        internal uint ProcessId;
        internal IntPtr DefaultHeapId;
        internal uint ModuleId;
        internal uint Threads;
        internal uint ParentProcessId;
        internal int BasePriority;
        internal uint Flags;

        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
        internal string ExecutableFile;
    }

    private readonly record struct ProcessMemoryInfo(
        int ProcessId,
        int ParentProcessId,
        long WorkingSetBytes);
}
