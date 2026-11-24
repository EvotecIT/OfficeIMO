using System.Text.Json;
using WorkerConnectionException = OfficeIMO.Notebook.PowerShell.PowerShellNotebookProcess.WorkerConnectionException;

namespace OfficeIMO.Notebook.PowerShell;

/// <summary>Runs PowerShell notebook cells in a persistent, explicitly started PowerShell runspace.</summary>
/// <remarks>Only run notebooks whose code you trust. Cells execute with the current user's permissions in a child process.
/// Runs submitted to the same session are serialized. Interrupting an active cell closes its session.</remarks>
public sealed class PowerShellNotebookSession : IDisposable {
    private const int MaximumOutputCharacters = 16 * 1024 * 1024;
    private const int MaximumOutputRecords = 100_000;
    private readonly PowerShellNotebookProcess _worker;
    private readonly SemaphoreSlim _executionLock = new(1, 1);
    private int _executionCount;

    private PowerShellNotebookSession(PowerShellNotebookProcess worker) {
        _worker = worker;
    }

    /// <summary>Gets the version reported by the running PowerShell executable.</summary>
    public string PowerShellVersion => _worker.Version;

    /// <summary>Gets whether this session can accept another notebook execution.</summary>
    public bool IsUsable => _worker.IsOpen;

    /// <summary>Starts a PowerShell process and waits for its persistent notebook runspace.</summary>
    /// <param name="options">PowerShell executable, working directory, and startup deadline.</param>
    /// <param name="cancellationToken">Cancels startup and closes any process already started.</param>
    /// <returns>The started session, which the caller must dispose.</returns>
    public static async Task<PowerShellNotebookSession> StartAsync(PowerShellNotebookSessionOptions? options = null,
        CancellationToken cancellationToken = default) {
        string executable = options?.ExecutablePath ?? "pwsh";
        ArgumentException.ThrowIfNullOrWhiteSpace(executable);
        string? directory = options?.WorkingDirectory;
        if (directory != null) directory = Path.GetFullPath(directory);
        TimeSpan timeout = options?.StartupTimeout ?? TimeSpan.FromSeconds(30);
        ValidateTimeout(timeout, nameof(PowerShellNotebookSessionOptions.StartupTimeout));
        PowerShellNotebookProcess worker = await PowerShellNotebookProcess.StartAsync(executable, directory, timeout, cancellationToken)
            .ConfigureAwait(false);
        return new PowerShellNotebookSession(worker);
    }

    /// <summary>Runs selected PowerShell cells in notebook order and replaces their saved outputs and execution counts.</summary>
    /// <param name="notebook">The notebook to update. The caller retains ownership.</param>
    /// <param name="options">Cell selection, timeout and output settings.</param>
    /// <param name="cancellationToken">Cancels the run. An active cell retains received outputs and receives a cancellation error.</param>
    /// <returns>Execution outcomes and the updated notebook.</returns>
    /// <remarks>Later executions on this session see variables and functions previously defined in the runspace.
    /// A cell timeout ends the session and returns a report. Caller cancellation throws <see cref="OperationCanceledException"/>.
    /// Edit or save the notebook after this task completes.</remarks>
    public async Task<NotebookExecutionReport> ExecuteAsync(NotebookDocument notebook, NotebookExecutionOptions? options = null,
        CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(notebook);
        NotebookExecutionOptions settings = SnapshotOptions(options);
        await _executionLock.WaitAsync(cancellationToken).ConfigureAwait(false);
        try {
            if (!IsUsable) throw new InvalidOperationException("The PowerShell notebook session has ended. Start a new session to run more cells.");
            return await ExecuteSelectedAsync(notebook, settings, cancellationToken).ConfigureAwait(false);
        } finally {
            _executionLock.Release();
        }
    }

    /// <summary>Closes the session and stops the worker if it is still running.</summary>
    public void Dispose() => _worker.Dispose();

    private async Task<NotebookExecutionReport> ExecuteSelectedAsync(NotebookDocument notebook, NotebookExecutionOptions settings,
        CancellationToken cancellationToken) {
        var results = new List<NotebookCellExecutionResult>();
        var skipped = new List<int>();
        string? notebookLanguage = GetNotebookLanguage(notebook);
        IEnumerable<int> indices = settings.CellIndices == null
            ? Enumerable.Range(0, notebook.Cells.Count) : settings.CellIndices.Distinct().OrderBy(static index => index);
        int[] selected = indices.ToArray();
        foreach (int index in selected) {
            if (index < 0 || index >= notebook.Cells.Count) {
                throw new ArgumentOutOfRangeException(nameof(NotebookExecutionOptions.CellIndices), index, "A selected index is outside the notebook.");
            }
        }
        foreach (int index in selected) {
            cancellationToken.ThrowIfCancellationRequested();
            NotebookCell cell = notebook.Cells[index];
            if (cell.Kind != NotebookCellKind.Code || !IsPowerShell(cell.Language ?? notebookLanguage, settings.IncludeUnspecifiedLanguage)) {
                skipped.Add(index);
                continue;
            }
            NotebookCellExecutionResult result = await ExecuteCellAsync(index, cell, settings, cancellationToken).ConfigureAwait(false);
            results.Add(result);
            if (result.Status != NotebookCellExecutionStatus.Succeeded &&
                (result.Status != NotebookCellExecutionStatus.Failed || !settings.ContinueOnError)) {
                break;
            }
        }
        return new NotebookExecutionReport(notebook, results, skipped);
    }

    private async Task<NotebookCellExecutionResult> ExecuteCellAsync(int index, NotebookCell cell,
        NotebookExecutionOptions settings, CancellationToken cancellationToken) {
        string source = cell.Source;
        cancellationToken.ThrowIfCancellationRequested();
        int executionCount = checked(++_executionCount);
        cell.ClearOutputs();
        cell.ExecutionCount = executionCount;

        using var deadline = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        if (settings.CellTimeout is TimeSpan timeout) deadline.CancelAfter(timeout);
        using CancellationTokenRegistration interrupt = _worker.InterruptWhenCanceled(deadline.Token);
        bool outputTruncated = false;
        try {
            await _worker.SendAsync(source, settings.OutputCharactersLimit, settings.OutputRecordsLimit, deadline.Token).ConfigureAwait(false);
            while (true) {
                deadline.Token.ThrowIfCancellationRequested();
                using JsonDocument message = await _worker.ReadAsync().ConfigureAwait(false);
                JsonElement frame = message.RootElement;
                string kind = PowerShellNotebookProcess.GetKind(frame);
                if (kind == "complete") {
                    deadline.CancelAfter(Timeout.Infinite);
                    interrupt.Dispose();
                    deadline.Token.ThrowIfCancellationRequested();
                    NotebookCellExecutionStatus status = ReadFlag(frame, "failed")
                        ? NotebookCellExecutionStatus.Failed : NotebookCellExecutionStatus.Succeeded;
                    return new NotebookCellExecutionResult(index, cell, executionCount, status, ReadFlag(frame, "truncated") || outputTruncated);
                }
                if (kind == "truncated") {
                    outputTruncated = true;
                    continue;
                }
                if (kind == "fatal") {
                    throw new WorkerConnectionException(ReadText(frame, "text").GetString()!);
                }
                NotebookOutput output = ReadOutput(frame, kind, executionCount);
                cell.AddOutput(output);
                settings.OutputReceived?.Invoke(new NotebookExecutionEvent(index, cell.Id, executionCount, output));
            }
        } catch (Exception error) when (error is OperationCanceledException or WorkerConnectionException) {
            _worker.Dispose();
            if (cancellationToken.IsCancellationRequested) {
                SaveInterruption(cell, "NotebookExecutionCanceled", "The PowerShell notebook execution was canceled.");
                cancellationToken.ThrowIfCancellationRequested();
            }
            if (deadline.IsCancellationRequested) {
                SaveInterruption(cell, "NotebookCellTimeout", "The PowerShell cell exceeded its deadline.");
                return new NotebookCellExecutionResult(index, cell, executionCount, NotebookCellExecutionStatus.TimedOut, outputTruncated);
            }
            if (error is WorkerConnectionException) {
                SaveInterruption(cell, "NotebookSessionEnded", error.Message);
                return new NotebookCellExecutionResult(index, cell, executionCount, NotebookCellExecutionStatus.SessionEnded, outputTruncated);
            }
            throw;
        } catch {
            // A callback or a notebook mutation may fail after the worker has started writing a response.
            _worker.Dispose();
            throw;
        }
    }

    private static void SaveInterruption(NotebookCell cell, string name, string message) =>
        cell.AddOutput(NotebookOutput.CreateError(name, message, Array.Empty<string>()));

    private static NotebookOutput ReadOutput(JsonElement frame, string kind, int executionCount) {
        JsonElement text = ReadText(frame, "text");
        return kind switch {
            "output" => NotebookOutput.CreateExecuteResult(new Dictionary<string, JsonElement> { ["text/plain"] = text }, executionCount),
            "stdout" or "stderr" => NotebookOutput.CreateStream(kind, text.GetString()!),
            "error" => NotebookOutput.CreateError(ReadText(frame, "name").GetString()!, text.GetString()!, ReadTraceback(frame)),
            _ => throw new WorkerConnectionException("The PowerShell worker returned an unknown notebook output.")
        };
    }

    private static string[] ReadTraceback(JsonElement frame) =>
        frame.TryGetProperty("traceback", out JsonElement traceback) && traceback.ValueKind == JsonValueKind.Array
            ? traceback.EnumerateArray().Where(static line => line.ValueKind == JsonValueKind.String)
                .Select(static line => line.GetString()!).ToArray()
            : Array.Empty<string>();

    private static JsonElement ReadText(JsonElement frame, string name) =>
        frame.TryGetProperty(name, out JsonElement value) && value.ValueKind == JsonValueKind.String ? value :
        throw new WorkerConnectionException("The PowerShell worker did not send the required notebook output text.");

    private static bool ReadFlag(JsonElement frame, string name) =>
        frame.TryGetProperty(name, out JsonElement value) && value.ValueKind is JsonValueKind.True or JsonValueKind.False
            ? value.GetBoolean() : throw new WorkerConnectionException("The PowerShell worker did not send the required notebook result flag.");

    private static string? GetNotebookLanguage(NotebookDocument notebook) {
        string? language = notebook.Language;
        if (!string.IsNullOrWhiteSpace(language)) return language;
        return notebook.Metadata.TryGetValue("kernelspec", out JsonElement kernel) && kernel.ValueKind == JsonValueKind.Object &&
            kernel.TryGetProperty("name", out JsonElement name) && name.ValueKind == JsonValueKind.String ? name.GetString() : null;
    }

    private static bool IsPowerShell(string? language, bool includeUnspecified) {
        if (string.IsNullOrWhiteSpace(language)) return includeUnspecified;
        string name = language!.Trim();
        return name.Equals("powershell", StringComparison.OrdinalIgnoreCase) ||
            name.Equals("pwsh", StringComparison.OrdinalIgnoreCase) ||
            name.Equals("ps1", StringComparison.OrdinalIgnoreCase) ||
            name.Equals(".net-powershell", StringComparison.OrdinalIgnoreCase);
    }

    private static NotebookExecutionOptions SnapshotOptions(NotebookExecutionOptions? options) {
        var snapshot = options == null ? new NotebookExecutionOptions() : new NotebookExecutionOptions {
            CellIndices = options.CellIndices?.ToArray(),
            CellTimeout = options.CellTimeout,
            ContinueOnError = options.ContinueOnError,
            IncludeUnspecifiedLanguage = options.IncludeUnspecifiedLanguage,
            OutputCharactersLimit = options.OutputCharactersLimit,
            OutputRecordsLimit = options.OutputRecordsLimit,
            OutputReceived = options.OutputReceived
        };
        if (snapshot.CellTimeout is TimeSpan timeout) ValidateTimeout(timeout, nameof(NotebookExecutionOptions.CellTimeout));
        if (snapshot.OutputCharactersLimit < 1 || snapshot.OutputCharactersLimit > MaximumOutputCharacters) {
            throw new ArgumentOutOfRangeException(nameof(NotebookExecutionOptions.OutputCharactersLimit));
        }
        if (snapshot.OutputRecordsLimit < 1 || snapshot.OutputRecordsLimit > MaximumOutputRecords) {
            throw new ArgumentOutOfRangeException(nameof(NotebookExecutionOptions.OutputRecordsLimit));
        }
        return snapshot;
    }

    private static void ValidateTimeout(TimeSpan timeout, string parameterName) {
        if (timeout.TotalMilliseconds < 1 || timeout.TotalMilliseconds > int.MaxValue) {
            throw new ArgumentOutOfRangeException(parameterName, "The timeout must be between one millisecond and Int32.MaxValue milliseconds.");
        }
    }
}
