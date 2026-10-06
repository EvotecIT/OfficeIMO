using System.ComponentModel;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Provenance;

namespace OfficeIMO.Provenance.C2pa;

internal interface IC2paToolProcessRunner {
    C2paToolProcessResult Run(C2paToolProcessRequest request, CancellationToken cancellationToken = default);
}

internal sealed class C2paToolProcessRequest {
    internal C2paToolProcessRequest(string executablePath, IReadOnlyList<string> arguments, string workingDirectory, TimeSpan timeout, long maximumOutputBytes) {
        ExecutablePath = executablePath;
        Arguments = arguments;
        WorkingDirectory = workingDirectory;
        Timeout = timeout;
        MaximumOutputBytes = maximumOutputBytes;
    }
    internal string ExecutablePath { get; }
    internal IReadOnlyList<string> Arguments { get; }
    internal string WorkingDirectory { get; }
    internal TimeSpan Timeout { get; }
    internal long MaximumOutputBytes { get; }
}

internal sealed class C2paToolProcessResult {
    internal C2paToolProcessResult(int exitCode, string standardOutput, string standardError) {
        ExitCode = exitCode;
        StandardOutput = standardOutput;
        StandardError = standardError;
    }
    internal int ExitCode { get; }
    internal string StandardOutput { get; }
    internal string StandardError { get; }
}

internal sealed class C2paToolProcessRunner : IC2paToolProcessRunner {
    private static readonly char[] ProcessSnapshotLineSeparators = { '\r', '\n' };
    private readonly bool _useExternalUnixSessionLauncher;

    internal C2paToolProcessRunner(bool useExternalUnixSessionLauncher = true) {
        _useExternalUnixSessionLauncher = useExternalUnixSessionLauncher;
    }

    public C2paToolProcessResult Run(
        C2paToolProcessRequest request,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        string targetExecutable = ResolveUnixExecutable(request.ExecutablePath, request.WorkingDirectory);
        string executable = targetExecutable;
        string arguments = string.Join(" ", request.Arguments.Select(QuoteArgument));
        string? sessionLauncher = _useExternalUnixSessionLauncher ? FindUnixSessionLauncher() : null;
        bool ownsUnixProcessGroup = sessionLauncher != null;
        if (Environment.OSVersion.Platform != PlatformID.Win32NT && sessionLauncher == null) {
            throw new Win32Exception(2,
                "c2patool execution on Unix requires a setsid executable so child processes can be contained safely.");
        }
        if (sessionLauncher != null) {
            executable = sessionLauncher;
            arguments = QuoteArgument(targetExecutable) + (arguments.Length == 0 ? string.Empty : " " + arguments);
        }
        var startInfo = new ProcessStartInfo {
            FileName = executable,
            Arguments = arguments,
            WorkingDirectory = request.WorkingDirectory,
            UseShellExecute = false,
            CreateNoWindow = true,
            RedirectStandardOutput = true,
            RedirectStandardError = true
        };
        using var process = new Process { StartInfo = startInfo };
        process.Start();
        using C2paToolProcessContainment containment = C2paToolProcessContainment.Create(process, ownsUnixProcessGroup);
        Task<string> stdout = ReadBoundedAsync(process.StandardOutput.BaseStream, request.MaximumOutputBytes, "standard output");
        Task<string> stderr = ReadBoundedAsync(process.StandardError.BaseStream, request.MaximumOutputBytes, "standard error");
        Stopwatch timer = Stopwatch.StartNew();
        while (true) {
            if (process.WaitForExit(50)) break;
            ThrowIfCancellationRequested(process, containment, cancellationToken);
            if (stdout.IsFaulted || stderr.IsFaulted) {
                Terminate(process, containment);
                throw stdout.Exception?.GetBaseException() ?? stderr.Exception?.GetBaseException() ?? new InvalidDataException("c2patool output failed.");
            }
            if (timer.Elapsed > request.Timeout) {
                Terminate(process, containment);
                throw new TimeoutException($"c2patool exceeded the configured timeout of {request.Timeout}.");
            }
        }
        try {
            var outputTasks = new Task[] { stdout, stderr };
            while (true) {
                TimeSpan remaining = request.Timeout - timer.Elapsed;
                if (remaining <= TimeSpan.Zero) {
                    Terminate(process, containment);
                    throw new TimeoutException($"c2patool exceeded the configured timeout of {request.Timeout}.");
                }
                int waitMilliseconds = (int)Math.Min(50D, Math.Ceiling(remaining.TotalMilliseconds));
                if (Task.WaitAll(outputTasks, waitMilliseconds, CancellationToken.None)) break;
                ThrowIfCancellationRequested(process, containment, cancellationToken);
            }
        } catch (AggregateException exception) {
            throw exception.GetBaseException();
        }
        cancellationToken.ThrowIfCancellationRequested();
        return new C2paToolProcessResult(process.ExitCode, stdout.Result, stderr.Result);
    }

    private static void ThrowIfCancellationRequested(
        Process process,
        C2paToolProcessContainment containment,
        CancellationToken cancellationToken) {
        if (!cancellationToken.IsCancellationRequested) return;
        Terminate(process, containment);
        cancellationToken.ThrowIfCancellationRequested();
    }

    internal static Task<string> ReadBoundedAsync(Stream stream, long maximumBytes, string streamName) => Task.Run(() => {
        try {
            using var reader = new StreamReader(
                stream,
                new UTF8Encoding(encoderShouldEmitUTF8Identifier: false, throwOnInvalidBytes: true),
                detectEncodingFromByteOrderMarks: true,
                bufferSize: 4096,
                leaveOpen: true);
            var builder = new StringBuilder();
            char[] buffer = new char[4096];
            long bytes = 0;
            while (true) {
                int read = reader.Read(buffer, 0, buffer.Length);
                if (read <= 0) break;
                bytes += Encoding.UTF8.GetByteCount(buffer, 0, read);
                if (bytes > maximumBytes) throw new InvalidDataException($"c2patool {streamName} exceeds the configured limit of {maximumBytes} bytes.");
                builder.Append(buffer, 0, read);
            }
            return builder.ToString();
        } catch (DecoderFallbackException exception) {
            throw new InvalidDataException($"c2patool {streamName} is not valid UTF-8.", exception);
        }
    });

    private static void Terminate(Process process, C2paToolProcessContainment containment) {
        try {
            containment.Terminate();
            if (!TryKillEntireProcessTree(process) && !process.HasExited) {
                process.Kill();
                process.WaitForExit(1000);
            }
        } catch (InvalidOperationException) { }
        catch (Win32Exception) { }
        finally {
            try { process.StandardOutput.Dispose(); } catch (InvalidOperationException) { }
            try { process.StandardError.Dispose(); } catch (InvalidOperationException) { }
        }
    }

    private sealed class C2paToolProcessContainment : IDisposable {
        private const uint JobObjectLimitKillOnJobClose = 0x00002000;
        private IntPtr _job;
        private readonly int _unixProcessGroupId;

        private C2paToolProcessContainment(IntPtr job, int unixProcessGroupId = 0) {
            _job = job;
            _unixProcessGroupId = unixProcessGroupId;
        }

        internal static C2paToolProcessContainment Create(Process process, bool ownsUnixProcessGroup) {
            if (Environment.OSVersion.Platform != PlatformID.Win32NT) {
                return new C2paToolProcessContainment(IntPtr.Zero, ownsUnixProcessGroup ? process.Id : 0);
            }
            IntPtr job = CreateJobObject(IntPtr.Zero, null);
            if (job == IntPtr.Zero) return new C2paToolProcessContainment(IntPtr.Zero);
            var information = new JobObjectExtendedLimitInformation {
                BasicLimitInformation = new JobObjectBasicLimitInformation { LimitFlags = JobObjectLimitKillOnJobClose }
            };
            int length = Marshal.SizeOf<JobObjectExtendedLimitInformation>();
            if (!SetInformationJobObject(job, 9, ref information, length) || !AssignProcessToJobObject(job, process.Handle)) {
                CloseHandle(job);
                return new C2paToolProcessContainment(IntPtr.Zero);
            }
            return new C2paToolProcessContainment(job);
        }

        internal void Terminate() => Dispose();

        public void Dispose() {
            if (_unixProcessGroupId > 0) _ = KillUnixProcessGroup(-_unixProcessGroupId, 9);
            IntPtr job = _job;
            if (job == IntPtr.Zero) return;
            _job = IntPtr.Zero;
            CloseHandle(job);
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
        private struct JobObjectExtendedLimitInformation {
            internal JobObjectBasicLimitInformation BasicLimitInformation;
            internal IoCounters IoInfo;
            internal UIntPtr ProcessMemoryLimit;
            internal UIntPtr JobMemoryLimit;
            internal UIntPtr PeakProcessMemoryUsed;
            internal UIntPtr PeakJobMemoryUsed;
        }

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
        private static extern IntPtr CreateJobObject(IntPtr securityAttributes, string? name);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool SetInformationJobObject(IntPtr job, int informationClass, ref JobObjectExtendedLimitInformation information, int informationLength);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool AssignProcessToJobObject(IntPtr job, IntPtr process);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool CloseHandle(IntPtr handle);

        [DllImport("libc", EntryPoint = "kill", SetLastError = true)]
        private static extern int KillUnixProcessGroup(int processId, int signal);
    }

    private static string? FindUnixSessionLauncher() {
        if (Environment.OSVersion.Platform == PlatformID.Win32NT) return null;
        foreach (string path in new[] {
            "/usr/bin/setsid",
            "/bin/setsid",
            "/usr/local/bin/setsid",
            "/opt/homebrew/opt/util-linux/bin/setsid",
            "/usr/local/opt/util-linux/bin/setsid"
        }) {
            if (File.Exists(path) && IsUnixExecutable(path)) return path;
        }
        string searchPath = Environment.GetEnvironmentVariable("PATH") ?? string.Empty;
        foreach (string directory in searchPath.Split(Path.PathSeparator)) {
            if (string.IsNullOrWhiteSpace(directory)) continue;
            string candidate;
            try { candidate = Path.GetFullPath(Path.Combine(directory.Trim(), "setsid")); }
            catch (Exception exception) when (exception is ArgumentException || exception is NotSupportedException || exception is PathTooLongException) { continue; }
            if (File.Exists(candidate) && IsUnixExecutable(candidate)) return candidate;
        }
        return null;
    }

    private static string ResolveUnixExecutable(string configuredPath, string workingDirectory) {
        if (Environment.OSVersion.Platform == PlatformID.Win32NT) return configuredPath;
        bool containsSeparator = configuredPath.Contains(Path.DirectorySeparatorChar.ToString()) ||
            configuredPath.Contains(Path.AltDirectorySeparatorChar.ToString());
        if (containsSeparator || Path.IsPathRooted(configuredPath)) {
            string candidate = Path.IsPathRooted(configuredPath)
                ? Path.GetFullPath(configuredPath)
                : Path.GetFullPath(Path.Combine(workingDirectory, configuredPath));
            if (File.Exists(candidate) && IsUnixExecutable(candidate)) return candidate;
            throw new Win32Exception(2, $"The configured c2patool executable '{configuredPath}' was not found or is not executable.");
        }

        string path = Environment.GetEnvironmentVariable("PATH") ?? string.Empty;
        foreach (string directory in path.Split(Path.PathSeparator)) {
            if (string.IsNullOrWhiteSpace(directory)) continue;
            string candidate;
            try { candidate = Path.GetFullPath(Path.Combine(directory.Trim(), configuredPath)); }
            catch (Exception exception) when (exception is ArgumentException || exception is NotSupportedException || exception is PathTooLongException) { continue; }
            if (File.Exists(candidate) && IsUnixExecutable(candidate)) return candidate;
        }
        throw new Win32Exception(2, $"The configured c2patool executable '{configuredPath}' was not found or is not executable.");
    }

    private static bool IsUnixExecutable(string path) => UnixAccess(path, 1) == 0;

    [DllImport("libc", EntryPoint = "access", SetLastError = true, CharSet = CharSet.Ansi,
        BestFitMapping = false, ThrowOnUnmappableChar = true)]
    private static extern int UnixAccess([MarshalAs(UnmanagedType.LPStr)] string path, int mode);

    private static bool TryKillEntireProcessTree(Process process) {
        System.Reflection.MethodInfo? method = typeof(Process).GetMethod("Kill", new[] { typeof(bool) });
        if (method != null) {
            try {
                method.Invoke(process, new object[] { true });
                return true;
            } catch (System.Reflection.TargetInvocationException exception) when (
                exception.InnerException is InvalidOperationException || exception.InnerException is Win32Exception) { }
        }
        if (Environment.OSVersion.Platform == PlatformID.Win32NT) return TryKillWithTaskKill(process);
        return TryKillUnixProcessTree(process);
    }

    private static bool TryKillWithTaskKill(Process process) {
        try {
            using Process? killer = Process.Start(new ProcessStartInfo {
                FileName = "taskkill.exe",
                Arguments = $"/PID {process.Id} /T /F",
                UseShellExecute = false,
                CreateNoWindow = true
            });
            if (killer == null || !killer.WaitForExit(2000)) return false;
            return killer.ExitCode == 0 || process.HasExited;
        } catch (InvalidOperationException) { return false; }
        catch (Win32Exception) { return false; }
    }

    private static bool TryKillUnixProcessTree(Process process) {
        try {
            var children = new Dictionary<int, List<int>>();
            using Process? snapshot = Process.Start(new ProcessStartInfo {
                FileName = "/bin/ps",
                Arguments = "-e -o pid= -o ppid=",
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            });
            if (snapshot == null) return false;
            string output = ReadProcessSnapshot(snapshot.StandardOutput, 4 * 1024 * 1024);
            if (!snapshot.WaitForExit(2000) || snapshot.ExitCode != 0) return false;
            foreach (string line in output.Split(ProcessSnapshotLineSeparators, StringSplitOptions.RemoveEmptyEntries)) {
                string[] fields = line.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
                if (fields.Length != 2 || !int.TryParse(fields[0], out int pid) || !int.TryParse(fields[1], out int parent)) continue;
                if (!children.TryGetValue(parent, out List<int>? list)) children.Add(parent, list = new List<int>());
                list.Add(pid);
            }
            var descendants = new List<int>();
            var pending = new Stack<int>();
            pending.Push(process.Id);
            while (pending.Count > 0) {
                int parent = pending.Pop();
                if (!children.TryGetValue(parent, out List<int>? direct)) continue;
                foreach (int child in direct) { descendants.Add(child); pending.Push(child); }
            }
            process.Kill();
            for (int index = descendants.Count - 1; index >= 0; index--) {
                try { using Process child = Process.GetProcessById(descendants[index]); child.Kill(); }
                catch (ArgumentException) { }
                catch (InvalidOperationException) { }
                catch (Win32Exception) { }
            }
            return true;
        } catch (InvalidOperationException) { return false; }
        catch (Win32Exception) { return false; }
        catch (InvalidDataException) { return false; }
    }

    private static string ReadProcessSnapshot(TextReader reader, int maximumCharacters) {
        var builder = new StringBuilder();
        char[] buffer = new char[4096];
        while (true) {
            int read = reader.Read(buffer, 0, buffer.Length);
            if (read <= 0) return builder.ToString();
            if (builder.Length > maximumCharacters - read) throw new InvalidDataException("The process-tree snapshot exceeds its safety limit.");
            builder.Append(buffer, 0, read);
        }
    }

    private static string QuoteArgument(string value) {
        if (value.Length > 0 && value.All(character => !char.IsWhiteSpace(character) && character != '"')) return value;
        var builder = new StringBuilder("\"");
        int backslashes = 0;
        foreach (char character in value) {
            if (character == '\\') { backslashes++; continue; }
            if (character == '"') {
                builder.Append('\\', backslashes * 2 + 1).Append('"');
                backslashes = 0;
                continue;
            }
            builder.Append('\\', backslashes).Append(character);
            backslashes = 0;
        }
        builder.Append('\\', backslashes * 2).Append('"');
        return builder.ToString();
    }
}
