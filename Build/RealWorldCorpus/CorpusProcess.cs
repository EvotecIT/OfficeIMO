using System.Diagnostics;
using System.Reflection;
using System.Text.Json;

namespace OfficeIMO.RealWorldCorpus;

internal static class CorpusProcess {
    public static async Task<CorpusProcessResult> RunAsync(
        string command,
        string stage,
        string inputPath,
        long maxFileBytes,
        TimeSpan timeout,
        CancellationToken cancellationToken) {
        ProcessStartInfo startInfo = CreateStartInfo(command, stage, inputPath, maxFileBytes);
        using var process = new Process { StartInfo = startInfo };
        var stopwatch = Stopwatch.StartNew();
        if (!process.Start()) throw new InvalidOperationException("Unable to start the corpus worker process.");

        Task<string> stdout = process.StandardOutput.ReadToEndAsync(cancellationToken);
        Task<string> stderr = process.StandardError.ReadToEndAsync(cancellationToken);
        using var timeoutSource = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        timeoutSource.CancelAfter(timeout);
        try {
            await process.WaitForExitAsync(timeoutSource.Token).ConfigureAwait(false);
            string output = await stdout.ConfigureAwait(false);
            _ = await stderr.ConfigureAwait(false);
            stopwatch.Stop();
            if (process.ExitCode != 0) {
                return CorpusProcessResult.Failed(stopwatch.ElapsedMilliseconds, "worker-exit-" + process.ExitCode);
            }
            CorpusWorkerResult? result = JsonSerializer.Deserialize<CorpusWorkerResult>(output, CorpusJson.Options);
            return result == null
                ? CorpusProcessResult.Failed(stopwatch.ElapsedMilliseconds, "worker-empty-result")
                : CorpusProcessResult.Completed(stopwatch.ElapsedMilliseconds, result);
        } catch (OperationCanceledException) when (!cancellationToken.IsCancellationRequested) {
            TryKill(process);
            await IgnoreFailure(stdout).ConfigureAwait(false);
            await IgnoreFailure(stderr).ConfigureAwait(false);
            stopwatch.Stop();
            return CorpusProcessResult.TimedOut(stopwatch.ElapsedMilliseconds);
        } catch {
            TryKill(process);
            throw;
        }
    }

    private static ProcessStartInfo CreateStartInfo(string command, string stage, string inputPath, long maxFileBytes) {
        string assemblyPath = Assembly.GetExecutingAssembly().Location;
        string processPath = Environment.ProcessPath ?? throw new InvalidOperationException("Current process path is unavailable.");
        bool hostedByDotNet = string.Equals(Path.GetFileNameWithoutExtension(processPath), "dotnet", StringComparison.OrdinalIgnoreCase);
        var startInfo = new ProcessStartInfo {
            FileName = hostedByDotNet ? processPath : processPath,
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true
        };
        if (hostedByDotNet) startInfo.ArgumentList.Add(assemblyPath);
        startInfo.ArgumentList.Add(command);
        startInfo.ArgumentList.Add("--stage");
        startInfo.ArgumentList.Add(stage);
        startInfo.ArgumentList.Add("--input");
        startInfo.ArgumentList.Add(inputPath);
        startInfo.ArgumentList.Add("--max-file-bytes");
        startInfo.ArgumentList.Add(maxFileBytes.ToString(System.Globalization.CultureInfo.InvariantCulture));
        return startInfo;
    }

    private static void TryKill(Process process) {
        try {
            if (!process.HasExited) process.Kill(entireProcessTree: true);
        } catch (InvalidOperationException) {
        } catch (System.ComponentModel.Win32Exception) {
        }
    }

    private static async Task IgnoreFailure(Task task) {
        try { await task.ConfigureAwait(false); } catch { }
    }
}

internal sealed class CorpusProcessResult {
    public bool IsTimedOut { get; private set; }
    public long DurationMilliseconds { get; private set; }
    public CorpusWorkerResult? Worker { get; private set; }
    public string? FailureCode { get; private set; }

    public static CorpusProcessResult Completed(long durationMilliseconds, CorpusWorkerResult worker) => new() {
        DurationMilliseconds = durationMilliseconds,
        Worker = worker
    };

    public static CorpusProcessResult Failed(long durationMilliseconds, string failureCode) => new() {
        DurationMilliseconds = durationMilliseconds,
        FailureCode = failureCode
    };

    public static CorpusProcessResult TimedOut(long durationMilliseconds) => new() {
        IsTimedOut = true,
        DurationMilliseconds = durationMilliseconds
    };
}
