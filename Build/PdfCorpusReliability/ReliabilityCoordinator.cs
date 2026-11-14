using System.Diagnostics;
using System.Globalization;
using System.Reflection;
using System.Text;
using System.Text.Json;
using System.Threading.Channels;

namespace OfficeIMO.PdfCorpusReliability;

internal static class ReliabilityCoordinator {
    public static async Task RunAsync(RunOptions options) {
        Directory.CreateDirectory(options.OutputDirectory);
        List<CorpusEntry> entries = CorpusInventory.Load(options, out int missing);
        var done = new HashSet<string>(StringComparer.Ordinal);
        if (options.Resume && File.Exists(options.ResultsLinesPath)) {
            // An interrupted run can leave a partial last record; drop it so the next result starts on its own line.
            ReliabilityReport.TrimIncompleteTail(options.ResultsLinesPath);
            foreach (FileResult result in ReliabilityReport.ReadResults(options.ResultsLinesPath)) done.Add(result.Path);
        } else if (File.Exists(options.ResultsLinesPath)) {
            File.Delete(options.ResultsLinesPath);
        }
        List<CorpusEntry> pending = entries.Where(entry => !done.Contains(entry.Path)).ToList();
        Console.WriteLine($"{entries.Count} files selected, {missing} manifest entries missing on disk, {pending.Count} to run with {options.Parallelism} workers.");

        var queue = Channel.CreateUnbounded<CorpusEntry>();
        foreach (CorpusEntry entry in pending) queue.Writer.TryWrite(entry);
        queue.Writer.Complete();

        using var output = new StreamWriter(new FileStream(options.ResultsLinesPath, FileMode.Append, FileAccess.Write, FileShare.Read), new UTF8Encoding(false));
        var gate = new object();
        int completed = 0;
        var stopwatch = Stopwatch.StartNew();
        string tempRoot = Path.Combine(options.OutputDirectory, "tmp");
        Directory.CreateDirectory(tempRoot);

        async Task Slot(int index) {
            using var host = new WorkerHost(options, Path.Combine(tempRoot, "w" + index.ToString(CultureInfo.InvariantCulture)));
            await foreach (CorpusEntry entry in queue.Reader.ReadAllAsync().ConfigureAwait(false)) {
                FileResult result = await host.ProcessAsync(entry).ConfigureAwait(false);
                lock (gate) {
                    output.WriteLine(JsonSerializer.Serialize(result, ReliabilityJson.Line));
                    output.Flush();
                    completed++;
                    if (completed % 250 == 0 || Outcome.Severity(result.Outcome) >= 4) {
                        Console.WriteLine($"[{completed}/{pending.Count} {stopwatch.Elapsed:hh\\:mm\\:ss}] {result.Outcome,-10} {result.Path}");
                    }
                }
            }
        }

        await Task.WhenAll(Enumerable.Range(0, options.Parallelism).Select(Slot)).ConfigureAwait(false);
        output.Close();
        try { Directory.Delete(tempRoot, recursive: true); } catch (IOException) { } catch (UnauthorizedAccessException) { }
        ReliabilityReport.Write(options, missing);
        Console.WriteLine($"Finished {completed} files in {stopwatch.Elapsed:hh\\:mm\\:ss}.");
        Console.WriteLine($"Summary: {options.SummaryPath}");
    }
}

/// <summary>One recycled worker process with timeout, crash and working-set supervision.</summary>
internal sealed class WorkerHost : IDisposable {
    private readonly RunOptions _options;
    private readonly string _tempDirectory;
    private Process? _process;
    private Channel<string>? _lines;
    private readonly StringBuilder _stderr = new();

    public WorkerHost(RunOptions options, string tempDirectory) {
        _options = options;
        _tempDirectory = tempDirectory;
        Directory.CreateDirectory(tempDirectory);
    }

    public async Task<FileResult> ProcessAsync(CorpusEntry entry) {
        var result = new FileResult {
            Path = entry.Path,
            Collection = entry.Collection,
            Source = entry.Source,
            Sha256 = entry.Sha256,
            Bytes = entry.Bytes,
            QpdfExitCode = entry.QpdfExitCode,
            EncryptionNote = entry.CertificateEncrypted ? "certificate-encrypted (key store not used)"
                : entry.PasswordsUnpublished ? "password not published"
                : entry.Passwords is not null ? "published test password" : null
        };
        EnsureStarted();
        Process process = _process!;
        var request = new WorkerRequest {
            FullPath = entry.FullPath,
            Passwords = entry.Passwords,
            RenderPages = _options.RenderPages,
            PngDirectory = _options.PngDirectory
        };
        lock (_stderr) _stderr.Clear();
        var stopwatch = Stopwatch.StartNew();
        long peak = SampleWorkingSet(process);
        string? inFlight = null;
        try {
            await process.StandardInput.WriteLineAsync(JsonSerializer.Serialize(request, ReliabilityJson.Line)).ConfigureAwait(false);
            await process.StandardInput.FlushAsync().ConfigureAwait(false);
        } catch (IOException) {
            return Lost(result, Outcome.Crash, "worker pipe closed before request", null, stopwatch, peak);
        }

        ChannelReader<string> reader = _lines!.Reader;
        Task<bool>? wait = null;
        while (true) {
            wait ??= reader.WaitToReadAsync().AsTask();
            await Task.WhenAny(wait, Task.Delay(100)).ConfigureAwait(false);
            if (wait.IsCompleted) {
                if (!wait.Result) {
                    process.WaitForExit(5000);
                    string code = process.HasExited ? "0x" + process.ExitCode.ToString("X8", CultureInfo.InvariantCulture) : "unknown";
                    return Lost(result, Outcome.Crash, $"worker exited with {code}: {StderrTail()}", inFlight, stopwatch, peak);
                }
                wait = null;
                while (reader.TryRead(out string? line)) {
                    if (line.StartsWith(ReliabilityWorker.ProgressPrefix, StringComparison.Ordinal)) {
                        inFlight = line[ReliabilityWorker.ProgressPrefix.Length..];
                    } else if (line.StartsWith(ReliabilityWorker.ResultPrefix, StringComparison.Ordinal)) {
                        WorkerResponse? response;
                        try {
                            response = JsonSerializer.Deserialize<WorkerResponse>(line[ReliabilityWorker.ResultPrefix.Length..], ReliabilityJson.Line);
                        } catch (JsonException) {
                            response = null;
                        }
                        // A truncated or malformed frame is no trustworthy result: the worker is lost for this file.
                        if (response is null) return Lost(result, Outcome.Crash, "worker wrote a malformed result frame", inFlight, stopwatch, peak);
                        peak = Math.Max(peak, SampleWorkingSet(process));
                        result.WallMilliseconds = stopwatch.ElapsedMilliseconds;
                        result.PeakWorkingSetBytes = peak;
                        result.Worker = response;
                        result.Sha256Verified = entry.Sha256 is null || string.Equals(entry.Sha256, response.Sha256, StringComparison.OrdinalIgnoreCase);
                        result.Sha256 ??= response.Sha256;
                        result.Outcome = Outcome.Worst(response.Operations);
                        // A worker that ran out of managed memory exits after answering; replace it now so the
                        // next file is not blamed for that exit.
                        if (result.Outcome == Outcome.Memory) Kill();
                        return result;
                    }
                }
            }
            peak = Math.Max(peak, SampleWorkingSet(process));
            if (peak > _options.MemoryLimitBytes) {
                return Lost(result, Outcome.Memory, $"working set exceeded {_options.MemoryLimitBytes / (1024 * 1024)} MB", inFlight, stopwatch, peak);
            }
            if (stopwatch.Elapsed > _options.Timeout) {
                return Lost(result, Outcome.Timeout, $"no result after {_options.Timeout.TotalSeconds:0} s", inFlight, stopwatch, peak);
            }
        }
    }

    private FileResult Lost(FileResult result, string outcome, string failure, string? inFlight, Stopwatch stopwatch, long peak) {
        Kill();
        result.Outcome = outcome;
        result.ProcessFailure = failure;
        result.InFlightOperation = inFlight;
        result.WallMilliseconds = stopwatch.ElapsedMilliseconds;
        result.PeakWorkingSetBytes = peak;
        return result;
    }

    private static long SampleWorkingSet(Process process) {
        try {
            if (process.HasExited) return 0;
            process.Refresh();
            return process.WorkingSet64;
        } catch (InvalidOperationException) {
            return 0;
        }
    }

    private void EnsureStarted() {
        if (_process is { HasExited: false }) return;
        Kill();
        string assembly = Assembly.GetExecutingAssembly().Location;
        string host = Environment.ProcessPath ?? "dotnet";
        var startInfo = new ProcessStartInfo {
            FileName = host,
            UseShellExecute = false,
            RedirectStandardInput = true,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true,
            StandardOutputEncoding = Encoding.UTF8,
            WorkingDirectory = _tempDirectory
        };
        if (string.Equals(Path.GetFileNameWithoutExtension(host), "dotnet", StringComparison.OrdinalIgnoreCase)) startInfo.ArgumentList.Add(assembly);
        startInfo.ArgumentList.Add("worker");
        // Managed allocations fail with OutOfMemoryException before the working-set watchdog has to kill the process.
        long heapLimit = Math.Max(256L * 1024 * 1024, _options.MemoryLimitBytes - 512L * 1024 * 1024);
        startInfo.Environment["DOTNET_GCHeapHardLimit"] = "0x" + heapLimit.ToString("X", CultureInfo.InvariantCulture);
        startInfo.Environment["DOTNET_TieredPGO"] = "0";
        startInfo.Environment["TEMP"] = _tempDirectory;
        startInfo.Environment["TMP"] = _tempDirectory;
        var process = new Process { StartInfo = startInfo };
        process.Start();
        try { process.PriorityClass = ProcessPriorityClass.BelowNormal; } catch (InvalidOperationException) { } catch (System.ComponentModel.Win32Exception) { }
        var lines = Channel.CreateUnbounded<string>();
        _ = Task.Run(async () => {
            try {
                string? line;
                while ((line = await process.StandardOutput.ReadLineAsync().ConfigureAwait(false)) != null) lines.Writer.TryWrite(line);
            } catch (IOException) {
            } catch (ObjectDisposedException) {
            } finally {
                lines.Writer.TryComplete();
            }
        });
        process.ErrorDataReceived += (_, args) => {
            if (args.Data is null) return;
            lock (_stderr) {
                if (_stderr.Length < 64 * 1024) _stderr.AppendLine(args.Data);
            }
        };
        process.BeginErrorReadLine();
        _process = process;
        _lines = lines;
    }

    /// <summary>First stderr line plus the distinct OfficeIMO frames, which name the recursion in a stack overflow.</summary>
    private string StderrTail() {
        string[] lines;
        lock (_stderr) lines = _stderr.ToString().Split('\n', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        if (lines.Length == 0) return "no stderr";
        IEnumerable<string> frames = lines
            .Where(static line => line.StartsWith("at OfficeIMO.", StringComparison.Ordinal))
            .Select(static line => {
                string frame = line[3..];
                int generic = frame.IndexOf("[[", StringComparison.Ordinal);
                if (generic > 0) frame = frame[..generic];
                int args = frame.IndexOf('(');
                return args > 0 ? frame[..args] : frame;
            })
            .Distinct(StringComparer.Ordinal)
            .Take(8);
        return lines[0] + " frames: " + string.Join(" <- ", frames);
    }

    private void Kill() {
        Process? process = _process;
        _process = null;
        _lines = null;
        if (process is null) return;
        try {
            if (!process.HasExited) {
                process.Kill(entireProcessTree: true);
                process.WaitForExit(10_000);
            }
        } catch (InvalidOperationException) {
        } catch (System.ComponentModel.Win32Exception) {
        }
        process.Dispose();
    }

    public void Dispose() {
        Process? process = _process;
        if (process is { HasExited: false }) {
            try {
                process.StandardInput.Close();
                if (!process.WaitForExit(5000)) Kill();
            } catch (IOException) {
                Kill();
            }
        }
        _process?.Dispose();
        _process = null;
    }
}
