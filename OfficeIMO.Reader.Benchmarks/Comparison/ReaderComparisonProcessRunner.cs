using System.Diagnostics;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Reader.Benchmarks.Comparison;

internal sealed class ReaderComparisonProcessOutput {
    public string Status { get; set; } = string.Empty;
    public string Markdown { get; set; } = string.Empty;
    public string? Error { get; set; }
    public double DurationMilliseconds { get; set; }
    public long? PeakWorkingSetBytes { get; set; }
    public bool Rejected { get; set; }
}

internal static class ReaderComparisonProcessRunner {
    public static async Task<ReaderComparisonProcessOutput> RunAsync(
        ReaderComparisonRunnerConfiguration configuration,
        string inputPath,
        string outputPath,
        CancellationToken cancellationToken) {
        Validate(configuration);
        var startInfo = new ProcessStartInfo {
            FileName = configuration.FileName,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true
        };
        foreach (string argument in configuration.Arguments) {
            startInfo.ArgumentList.Add(argument
                .Replace("{input}", inputPath, StringComparison.Ordinal)
                .Replace("{output}", outputPath, StringComparison.Ordinal));
        }

        using var process = new Process { StartInfo = startInfo };
        var stopwatch = Stopwatch.StartNew();
        try {
            if (!process.Start()) return Failure("failed", "The process did not start.", stopwatch.Elapsed.TotalMilliseconds);
        } catch (Exception ex) when (ex is System.ComponentModel.Win32Exception or InvalidOperationException) {
            return Failure("unavailable", ex.Message, stopwatch.Elapsed.TotalMilliseconds);
        }

        using var timeout = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        timeout.CancelAfter(TimeSpan.FromSeconds(configuration.TimeoutSeconds));
        Task<BoundedText> stdoutTask = ReadBoundedAsync(
            process.StandardOutput.BaseStream,
            configuration.MaxOutputBytes,
            timeout.Token);
        Task<BoundedText> stderrTask = ReadBoundedAsync(
            process.StandardError.BaseStream,
            Math.Min(configuration.MaxOutputBytes, 1024 * 1024),
            timeout.Token);

        try {
            await process.WaitForExitAsync(timeout.Token).ConfigureAwait(false);
        } catch (OperationCanceledException) when (!cancellationToken.IsCancellationRequested) {
            TryKill(process);
            await DrainAfterTerminationAsync(stdoutTask, stderrTask).ConfigureAwait(false);
            stopwatch.Stop();
            return Failure("timed-out", "The runner exceeded its configured timeout.", stopwatch.Elapsed.TotalMilliseconds);
        }

        BoundedText stdout = await stdoutTask.ConfigureAwait(false);
        BoundedText stderr = await stderrTask.ConfigureAwait(false);
        stopwatch.Stop();
        string markdown;
        if (string.Equals(configuration.OutputMode, "file", StringComparison.OrdinalIgnoreCase)) {
            markdown = File.Exists(outputPath)
                ? await ReadFileBoundedAsync(outputPath, configuration.MaxOutputBytes, cancellationToken).ConfigureAwait(false)
                : string.Empty;
        } else {
            markdown = stdout.Text;
        }

        string? error = process.ExitCode == 0 ? null : EmptyToNull(stderr.Text) ?? "Runner exited with code " + process.ExitCode + ".";
        if (stdout.Truncated || stderr.Truncated) {
            error = Append(error, "Runner output was truncated at the configured byte limit.");
        }

        return new ReaderComparisonProcessOutput {
            Status = process.ExitCode == 0 ? "success" : "failed",
            Markdown = markdown,
            Error = error,
            DurationMilliseconds = stopwatch.Elapsed.TotalMilliseconds,
            PeakWorkingSetBytes = SafePeakWorkingSet(process),
            Rejected = process.ExitCode != 0
        };
    }

    private static void Validate(ReaderComparisonRunnerConfiguration configuration) {
        if (string.IsNullOrWhiteSpace(configuration.Name)) throw new InvalidDataException("Runner name is required.");
        if (string.IsNullOrWhiteSpace(configuration.FileName)) throw new InvalidDataException("Runner fileName is required.");
        if (!string.Equals(configuration.OutputMode, "stdout", StringComparison.OrdinalIgnoreCase) &&
            !string.Equals(configuration.OutputMode, "file", StringComparison.OrdinalIgnoreCase)) {
            throw new InvalidDataException("Runner outputMode must be 'stdout' or 'file'.");
        }
        if (configuration.TimeoutSeconds < 1 || configuration.TimeoutSeconds > 3600) {
            throw new InvalidDataException("Runner timeoutSeconds must be between 1 and 3600.");
        }
        if (configuration.MaxOutputBytes < 1024 || configuration.MaxOutputBytes > 256 * 1024 * 1024) {
            throw new InvalidDataException("Runner maxOutputBytes must be between 1024 and 268435456.");
        }
    }

    private static async Task<BoundedText> ReadBoundedAsync(Stream stream, int maxBytes, CancellationToken cancellationToken) {
        byte[] buffer = new byte[8192];
        using var captured = new MemoryStream(Math.Min(maxBytes, 64 * 1024));
        bool truncated = false;
        while (true) {
            int read = await stream.ReadAsync(buffer.AsMemory(0, buffer.Length), cancellationToken).ConfigureAwait(false);
            if (read == 0) break;
            int remaining = maxBytes - (int)captured.Length;
            if (remaining > 0) captured.Write(buffer, 0, Math.Min(read, remaining));
            if (read > remaining) truncated = true;
        }
        return new BoundedText(Encoding.UTF8.GetString(captured.ToArray()), truncated);
    }

    private static async Task<string> ReadFileBoundedAsync(string path, int maxBytes, CancellationToken cancellationToken) {
        using FileStream stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
        BoundedText text = await ReadBoundedAsync(stream, maxBytes, cancellationToken).ConfigureAwait(false);
        return text.Text;
    }

    private static async Task DrainAfterTerminationAsync(Task<BoundedText> stdout, Task<BoundedText> stderr) {
        try {
            await Task.WhenAll(stdout, stderr).ConfigureAwait(false);
        } catch (OperationCanceledException) {
            // The timeout is already represented by the returned status.
        }
    }

    private static void TryKill(Process process) {
        try {
            if (!process.HasExited) process.Kill(entireProcessTree: true);
        } catch (InvalidOperationException) {
        } catch (System.ComponentModel.Win32Exception) {
        } catch (NotSupportedException) {
        }
    }

    private static long? SafePeakWorkingSet(Process process) {
        try {
            return process.PeakWorkingSet64;
        } catch (InvalidOperationException) {
            return null;
        }
    }

    private static ReaderComparisonProcessOutput Failure(string status, string error, double duration) =>
        new ReaderComparisonProcessOutput { Status = status, Error = error, DurationMilliseconds = duration, Rejected = true };

    private static string? EmptyToNull(string value) => string.IsNullOrWhiteSpace(value) ? null : value.Trim();

    private static string Append(string? current, string addition) =>
        string.IsNullOrWhiteSpace(current) ? addition : current + " " + addition;

    private readonly struct BoundedText {
        public BoundedText(string text, bool truncated) {
            Text = text;
            Truncated = truncated;
        }

        public string Text { get; }
        public bool Truncated { get; }
    }
}