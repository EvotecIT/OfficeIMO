using OfficeIMO.Reader.Ocr.Process;
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
        IReadOnlyList<string> arguments = configuration.Arguments
            .Select(argument => argument
                .Replace("{input}", inputPath, StringComparison.Ordinal)
                .Replace("{output}", outputPath, StringComparison.Ordinal))
            .ToArray();
        var startInfo = new ProcessStartInfo {
            FileName = configuration.FileName,
            Arguments = string.Join(" ", arguments.Select(OfficeOcrProcessRunner.QuoteArgument)),
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true
        };

        if (IsMissingUnixExecutable(configuration.FileName)) {
            return Failure(
                "unavailable",
                "Runner executable '" + configuration.FileName + "' is missing or is not executable.",
                0);
        }

        bool fileOutputMode = string.Equals(configuration.OutputMode, "file", StringComparison.OrdinalIgnoreCase);
        if (fileOutputMode) {
            try {
                File.Delete(outputPath);
            } catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) {
                return Failure("failed", "Could not remove the previous runner output: " + ex.Message, 0);
            }
        }

        var stopwatch = Stopwatch.StartNew();
        OfficeOcrProcessLifetime processLifetime;
        OfficeOcrStartedProcess startedProcess;
        try {
            processLifetime = OfficeOcrProcessLifetime.Configure(startInfo, configuration.FileName, arguments);
            try {
                startedProcess = processLifetime.Start(startInfo);
            } catch {
                processLifetime.Dispose();
                throw;
            }
        } catch (Exception ex) when (ex is System.ComponentModel.Win32Exception or InvalidOperationException or PlatformNotSupportedException) {
            return Failure("unavailable", ex.Message, stopwatch.Elapsed.TotalMilliseconds);
        }
        using (processLifetime)
        using (startedProcess) {
            Process process = startedProcess.Process;

            using var timeout = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
            timeout.CancelAfter(TimeSpan.FromSeconds(configuration.TimeoutSeconds));
            Task<BoundedText> stdoutTask = ReadBoundedAsync(
                GetBaseStream(startedProcess.StandardOutput),
                configuration.MaxOutputBytes,
                timeout.Token);
            Task<BoundedText> stderrTask = ReadBoundedAsync(
                GetBaseStream(startedProcess.StandardError),
                Math.Min(configuration.MaxOutputBytes, 1024 * 1024),
                timeout.Token);

            BoundedText stdout;
            BoundedText stderr;
            try {
                await process.WaitForExitAsync(timeout.Token).ConfigureAwait(false);
                BoundedText[] streams = await WaitWithCancellationAsync(
                    Task.WhenAll(stdoutTask, stderrTask),
                    timeout.Token).ConfigureAwait(false);
                stdout = streams[0];
                stderr = streams[1];
            } catch (OperationCanceledException) {
                processLifetime.Terminate(process);
                startedProcess.CloseRedirectedStreams();
                ObserveReadFailure(stdoutTask);
                ObserveReadFailure(stderrTask);
                stopwatch.Stop();
                cancellationToken.ThrowIfCancellationRequested();
                return Failure("timed-out", "The runner exceeded its configured timeout.", stopwatch.Elapsed.TotalMilliseconds);
            } catch {
                processLifetime.Terminate(process);
                startedProcess.CloseRedirectedStreams();
                ObserveReadFailure(stdoutTask);
                ObserveReadFailure(stderrTask);
                throw;
            }

            stopwatch.Stop();
            string markdown;
            bool missingFileOutput = false;
            bool fileOutputTruncated = false;
            if (fileOutputMode) {
                missingFileOutput = !File.Exists(outputPath);
                BoundedText fileOutput = missingFileOutput
                    ? new BoundedText(string.Empty, truncated: false)
                    : await ReadFileBoundedAsync(outputPath, configuration.MaxOutputBytes, cancellationToken).ConfigureAwait(false);
                markdown = fileOutput.Text;
                fileOutputTruncated = fileOutput.Truncated;
            } else {
                markdown = stdout.Text;
            }

            string? error = process.ExitCode == 0 ? null : EmptyToNull(stderr.Text) ?? "Runner exited with code " + process.ExitCode + ".";
            if (missingFileOutput) {
                error = Append(error, "Runner did not create the expected output file.");
            }
            bool truncated = stdout.Truncated || stderr.Truncated || fileOutputTruncated;
            if (truncated) {
                error = Append(error, "Runner output was truncated at the configured byte limit.");
            }

            bool succeeded = process.ExitCode == 0 && !missingFileOutput && !truncated;

            return new ReaderComparisonProcessOutput {
                Status = succeeded ? "success" : "failed",
                Markdown = markdown,
                Error = error,
                DurationMilliseconds = stopwatch.Elapsed.TotalMilliseconds,
                PeakWorkingSetBytes = SafePeakWorkingSet(process),
                Rejected = process.ExitCode != 0 && !missingFileOutput && !truncated
            };
        }
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

    private static async Task<BoundedText> ReadFileBoundedAsync(string path, int maxBytes, CancellationToken cancellationToken) {
        using FileStream stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
        byte[] buffer = new byte[8192];
        using var captured = new MemoryStream(Math.Min(maxBytes, 64 * 1024));
        while (captured.Length < maxBytes) {
            int remaining = maxBytes - (int)captured.Length;
            int read = await stream.ReadAsync(
                buffer.AsMemory(0, Math.Min(buffer.Length, remaining)),
                cancellationToken).ConfigureAwait(false);
            if (read == 0) break;
            captured.Write(buffer, 0, read);
        }

        bool truncated = stream.Position < stream.Length;
        return new BoundedText(Encoding.UTF8.GetString(captured.ToArray()), truncated);
    }

    private static bool IsMissingUnixExecutable(string fileName) {
        if (OperatingSystem.IsWindows()) return false;
        if (fileName.IndexOf(Path.DirectorySeparatorChar) >= 0 ||
            fileName.IndexOf(Path.AltDirectorySeparatorChar) >= 0) {
            return !IsExecutableUnixFile(fileName);
        }

        string? path = Environment.GetEnvironmentVariable("PATH");
        if (string.IsNullOrWhiteSpace(path)) return true;
        return !path.Split(Path.PathSeparator, StringSplitOptions.RemoveEmptyEntries)
            .Any(directory => IsExecutableUnixFile(Path.Combine(directory, fileName)));
    }

    private static bool IsExecutableUnixFile(string path) {
        if (OperatingSystem.IsWindows()) return false;
        if (!File.Exists(path)) return false;
        try {
            const UnixFileMode executeBits =
                UnixFileMode.UserExecute |
                UnixFileMode.GroupExecute |
                UnixFileMode.OtherExecute;
            return (File.GetUnixFileMode(path) & executeBits) != 0;
        } catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or PlatformNotSupportedException) {
            return false;
        }
    }

    private static Stream GetBaseStream(TextReader reader) {
        if (reader is StreamReader streamReader) return streamReader.BaseStream;
        throw new InvalidOperationException("The process reader is not backed by a stream.");
    }

    private static async Task<T> WaitWithCancellationAsync<T>(Task<T> operation, CancellationToken cancellationToken) {
        if (operation.IsCompleted) return await operation.ConfigureAwait(false);
        var cancellation = new TaskCompletionSource<object?>(TaskCreationOptions.RunContinuationsAsynchronously);
        using (cancellationToken.Register(() => cancellation.TrySetResult(null))) {
            Task completed = await Task.WhenAny(operation, cancellation.Task).ConfigureAwait(false);
            if (completed != operation) throw new OperationCanceledException(cancellationToken);
        }
        return await operation.ConfigureAwait(false);
    }

    private static void ObserveReadFailure(Task operation) {
        if (operation.IsCompleted) {
            _ = operation.Exception;
            return;
        }
        _ = operation.ContinueWith(
            static completed => { _ = completed.Exception; },
            CancellationToken.None,
            TaskContinuationOptions.ExecuteSynchronously | TaskContinuationOptions.OnlyOnFaulted,
            TaskScheduler.Default);
    }

    private static long? SafePeakWorkingSet(Process process) {
        try {
            return process.PeakWorkingSet64;
        } catch (InvalidOperationException) {
            return null;
        }
    }

    private static ReaderComparisonProcessOutput Failure(string status, string error, double duration) =>
        new ReaderComparisonProcessOutput { Status = status, Error = error, DurationMilliseconds = duration, Rejected = false };

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
