using System.Diagnostics;

namespace OfficeIMO.Reader.Ocr.Process;

/// <summary>One direct executable invocation used by optional OCR providers.</summary>
public sealed class OfficeOcrProcessCommand {
    /// <summary>Executable path or name resolved by the operating system.</summary>
    public string FileName { get; set; } = string.Empty;

    /// <summary>Argument values passed directly to the executable without invoking a shell.</summary>
    public IReadOnlyList<string> Arguments { get; set; } = Array.Empty<string>();

    /// <summary>Optional working directory.</summary>
    public string? WorkingDirectory { get; set; }

    /// <summary>Optional environment values applied to the child process.</summary>
    public IReadOnlyDictionary<string, string> EnvironmentVariables { get; set; } = new Dictionary<string, string>(StringComparer.Ordinal);

    /// <summary>Maximum process duration. Defaults to two minutes.</summary>
    public TimeSpan Timeout { get; set; } = TimeSpan.FromMinutes(2);

    /// <summary>Maximum retained standard-output characters. The stream is drained after the bound is reached.</summary>
    public int MaxStandardOutputCharacters { get; set; } = 64 * 1024;

    /// <summary>Maximum retained standard-error characters. The stream is drained after the bound is reached.</summary>
    public int MaxStandardErrorCharacters { get; set; } = 64 * 1024;
}

/// <summary>Bounded output from one direct executable invocation.</summary>
public sealed class OfficeOcrProcessResult {
    /// <summary>Child process exit code.</summary>
    public int ExitCode { get; set; }

    /// <summary>Retained standard output.</summary>
    public string StandardOutput { get; set; } = string.Empty;

    /// <summary>Retained standard error.</summary>
    public string StandardError { get; set; } = string.Empty;

    /// <summary>Whether standard output exceeded its retention bound.</summary>
    public bool StandardOutputTruncated { get; set; }

    /// <summary>Whether standard error exceeded its retention bound.</summary>
    public bool StandardErrorTruncated { get; set; }
}

/// <summary>Runs a configured executable directly with bounded output, timeout, and cancellation handling.</summary>
public static class OfficeOcrProcessRunner {
    /// <summary>Runs one process command without invoking a command shell.</summary>
    public static async Task<OfficeOcrProcessResult> RunAsync(OfficeOcrProcessCommand command, CancellationToken cancellationToken = default) {
        if (command == null) throw new ArgumentNullException(nameof(command));
        if (string.IsNullOrWhiteSpace(command.FileName)) throw new ArgumentException("Process filename cannot be empty.", nameof(command));
        if (command.Timeout <= TimeSpan.Zero) throw new ArgumentOutOfRangeException(nameof(command.Timeout));
        if (command.MaxStandardOutputCharacters < 1) throw new ArgumentOutOfRangeException(nameof(command.MaxStandardOutputCharacters));
        if (command.MaxStandardErrorCharacters < 1) throw new ArgumentOutOfRangeException(nameof(command.MaxStandardErrorCharacters));
        cancellationToken.ThrowIfCancellationRequested();
        using var timeout = new CancellationTokenSource(command.Timeout);
        using var linked = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken, timeout.Token);

        var startInfo = new ProcessStartInfo {
            FileName = command.FileName,
            Arguments = string.Join(" ", (command.Arguments ?? Array.Empty<string>()).Select(QuoteArgument)),
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true
        };
        if (!string.IsNullOrWhiteSpace(command.WorkingDirectory)) startInfo.WorkingDirectory = command.WorkingDirectory;
        foreach (KeyValuePair<string, string> variable in command.EnvironmentVariables ?? new Dictionary<string, string>()) {
            startInfo.EnvironmentVariables[variable.Key] = variable.Value;
        }

        using var process = new System.Diagnostics.Process { StartInfo = startInfo, EnableRaisingEvents = true };
        if (!process.Start()) throw new InvalidOperationException("Failed to start OCR process '" + command.FileName + "'.");
        Task<BoundedText> outputTask = ReadBoundedAsync(process.StandardOutput, command.MaxStandardOutputCharacters);
        Task<BoundedText> errorTask = ReadBoundedAsync(process.StandardError, command.MaxStandardErrorCharacters);
        try {
            await WaitForExitAsync(process, linked.Token).ConfigureAwait(false);
        } catch (OperationCanceledException) when (!cancellationToken.IsCancellationRequested && timeout.IsCancellationRequested) {
            TryKill(process);
            throw new TimeoutException("OCR process exceeded timeout " + command.Timeout + ".");
        } catch {
            TryKill(process);
            throw;
        }

        BoundedText[] streams = await Task.WhenAll(outputTask, errorTask).ConfigureAwait(false);
        return new OfficeOcrProcessResult {
            ExitCode = process.ExitCode,
            StandardOutput = streams[0].Text,
            StandardError = streams[1].Text,
            StandardOutputTruncated = streams[0].Truncated,
            StandardErrorTruncated = streams[1].Truncated
        };
    }

    private static async Task<BoundedText> ReadBoundedAsync(TextReader reader, int maxCharacters) {
        var builder = new StringBuilder(Math.Min(maxCharacters, 4096));
        var buffer = new char[4096];
        bool truncated = false;
        while (true) {
            int read = await reader.ReadAsync(buffer, 0, buffer.Length).ConfigureAwait(false);
            if (read == 0) break;
            int remaining = maxCharacters - builder.Length;
            if (remaining > 0) builder.Append(buffer, 0, Math.Min(remaining, read));
            if (read > remaining) truncated = true;
        }
        return new BoundedText(builder.ToString(), truncated);
    }

    private static Task WaitForExitAsync(System.Diagnostics.Process process, CancellationToken cancellationToken) {
        if (process.HasExited) return Task.CompletedTask;
        var completion = new TaskCompletionSource<object?>(TaskCreationOptions.RunContinuationsAsynchronously);
        EventHandler? handler = null;
        CancellationTokenRegistration registration = default;
        handler = (_, _) => completion.TrySetResult(null);
        process.Exited += handler;
        registration = cancellationToken.Register(() => {
            TryKill(process);
            completion.TrySetCanceled();
        });
        if (process.HasExited) completion.TrySetResult(null);
        return CompleteAndCleanupAsync(completion.Task, process, handler, registration);
    }

    private static async Task CompleteAndCleanupAsync(Task task, System.Diagnostics.Process process, EventHandler handler, CancellationTokenRegistration registration) {
        try {
            await task.ConfigureAwait(false);
        } finally {
            process.Exited -= handler;
            registration.Dispose();
        }
    }

    private static void TryKill(System.Diagnostics.Process process) {
        try {
            if (!process.HasExited) {
#if NET8_0_OR_GREATER
                process.Kill(entireProcessTree: true);
#else
                process.Kill();
#endif
            }
        } catch (InvalidOperationException) {
        }
    }

    internal static string QuoteArgument(string value) {
        if (value == null) return "\"\"";
        if (value.Length > 0 && value.All(static ch => !char.IsWhiteSpace(ch) && ch != '\"')) return value;
        var builder = new StringBuilder(value.Length + 2);
        builder.Append('\"');
        int slashes = 0;
        foreach (char character in value) {
            if (character == '\\') {
                slashes++;
                continue;
            }
            if (character == '\"') {
                builder.Append('\\', (slashes * 2) + 1);
                builder.Append('\"');
                slashes = 0;
                continue;
            }
            if (slashes > 0) builder.Append('\\', slashes);
            slashes = 0;
            builder.Append(character);
        }
        if (slashes > 0) builder.Append('\\', slashes * 2);
        builder.Append('\"');
        return builder.ToString();
    }

    private sealed class BoundedText {
        internal BoundedText(string text, bool truncated) {
            Text = text;
            Truncated = truncated;
        }

        internal string Text { get; }
        internal bool Truncated { get; }
    }
}
