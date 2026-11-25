using System.Text.Json;

namespace OfficeIMO.Notebook.PowerShell;

/// <summary>Runs PowerShell notebook cells in a persistent, explicitly started PowerShell runspace.</summary>
/// <remarks>Only run notebooks whose code you trust. Cells execute with the current user's permissions, in a child process
/// or, with <see cref="PowerShellNotebookIsolation.Runspace"/>, in the calling PowerShell process.
/// Runs submitted to the same session are serialized. Interrupting an active cell closes its session.</remarks>
public sealed class PowerShellNotebookSession : IDisposable {
    private const int MaximumOutputCharacters = 16 * 1024 * 1024;
    private const int MaximumOutputRecords = 100_000;
    private readonly IPowerShellNotebookWorker _worker;
    private readonly SemaphoreSlim _executionLock = new(1, 1);
    private int _executionCount;

    private PowerShellNotebookSession(IPowerShellNotebookWorker worker) {
        _worker = worker;
    }

    /// <summary>Gets the version reported by the PowerShell that runs the cells.</summary>
    public string PowerShellVersion => _worker.Version;

    /// <summary>Gets whether this session can accept another notebook execution.</summary>
    public bool IsUsable => _worker.IsOpen;

    /// <summary>Starts a PowerShell process, or a runspace in the calling PowerShell process, and waits for its persistent notebook runspace.</summary>
    /// <param name="options">Isolation, PowerShell executable, working directory, and startup deadline.</param>
    /// <param name="cancellationToken">Cancels startup and closes any process already started.</param>
    /// <returns>The started session, which the caller must dispose.</returns>
    /// <exception cref="PlatformNotSupportedException"><see cref="PowerShellNotebookIsolation.Runspace"/> was requested outside a PowerShell host.</exception>
    public static async Task<PowerShellNotebookSession> StartAsync(PowerShellNotebookSessionOptions? options = null,
        CancellationToken cancellationToken = default) {
        string? directory = options?.WorkingDirectory;
        if (directory != null) directory = Path.GetFullPath(directory);
        TimeSpan timeout = options?.StartupTimeout ?? TimeSpan.FromSeconds(30);
        ValidateTimeout(timeout, nameof(PowerShellNotebookSessionOptions.StartupTimeout));
        PowerShellNotebookIsolation isolation = options?.Isolation ?? PowerShellNotebookIsolation.Process;
        if (isolation == PowerShellNotebookIsolation.Runspace) {
            if (options?.ExecutablePath != null) {
                throw new ArgumentException("A runspace session uses the calling PowerShell. Remove ExecutablePath or choose process isolation.", nameof(options));
            }
            cancellationToken.ThrowIfCancellationRequested();
            return new PowerShellNotebookSession(PowerShellNotebookRunspaceHost.Start(directory));
        }
        if (isolation != PowerShellNotebookIsolation.Process) {
            throw new ArgumentOutOfRangeException(nameof(options), isolation, "Choose process or runspace isolation.");
        }
        string executable = options?.ExecutablePath ?? "pwsh";
        ArgumentException.ThrowIfNullOrWhiteSpace(executable);
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
            await _worker.SendAsync(source, executionCount, settings.OutputCharactersLimit, settings.OutputRecordsLimit, deadline.Token).ConfigureAwait(false);
            while (true) {
                deadline.Token.ThrowIfCancellationRequested();
                PowerShellNotebookFrame frame = await _worker.ReadAsync().ConfigureAwait(false);
                if (frame.Kind == PowerShellNotebookFrameKind.Complete) {
                    deadline.CancelAfter(Timeout.Infinite);
                    interrupt.Dispose();
                    deadline.Token.ThrowIfCancellationRequested();
                    NotebookCellExecutionStatus status = frame.Failed
                        ? NotebookCellExecutionStatus.Failed : NotebookCellExecutionStatus.Succeeded;
                    return new NotebookCellExecutionResult(index, cell, executionCount, status, frame.Truncated || outputTruncated);
                }
                if (frame.Kind == PowerShellNotebookFrameKind.Truncated) {
                    outputTruncated = true;
                    continue;
                }
                NotebookOutput output = frame.Output!;
                cell.AddOutput(output);
                settings.OutputReceived?.Invoke(new NotebookExecutionEvent(index, cell.Id, executionCount, output));
            }
        } catch (Exception error) when (error is OperationCanceledException or PowerShellNotebookConnectionException) {
            _worker.Dispose();
            if (cancellationToken.IsCancellationRequested) {
                SaveInterruption(cell, "NotebookExecutionCanceled", "The PowerShell notebook execution was canceled.");
                cancellationToken.ThrowIfCancellationRequested();
            }
            if (deadline.IsCancellationRequested) {
                SaveInterruption(cell, "NotebookCellTimeout", "The PowerShell cell exceeded its deadline.");
                return new NotebookCellExecutionResult(index, cell, executionCount, NotebookCellExecutionStatus.TimedOut, outputTruncated);
            }
            if (error is PowerShellNotebookConnectionException) {
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
