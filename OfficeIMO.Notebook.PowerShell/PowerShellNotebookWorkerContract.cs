namespace OfficeIMO.Notebook.PowerShell;

/// <summary>Runs notebook cells for one session and reports each cell's outputs and completion in order.</summary>
internal interface IPowerShellNotebookWorker : IDisposable {
    bool IsOpen { get; }

    string Version { get; }

    /// <summary>Starts one cell. Frames for the cell follow through <see cref="ReadAsync"/> until a completion frame.</summary>
    Task SendAsync(string source, int executionCount, int outputCharactersLimit, int outputRecordsLimit, CancellationToken cancellationToken);

    /// <summary>Returns the next frame of the active cell. A lost worker throws <see cref="PowerShellNotebookConnectionException"/>.</summary>
    Task<PowerShellNotebookFrame> ReadAsync();

    /// <summary>Closes the worker when the token is canceled, which ends any pending read.</summary>
    CancellationTokenRegistration InterruptWhenCanceled(CancellationToken token);
}

internal enum PowerShellNotebookFrameKind {
    Output,
    Truncated,
    Complete
}

/// <summary>One cell update: an output, a truncation notice or the cell's completion.</summary>
internal readonly struct PowerShellNotebookFrame {
    private PowerShellNotebookFrame(PowerShellNotebookFrameKind kind, NotebookOutput? output, bool failed, bool truncated) {
        Kind = kind;
        Output = output;
        Failed = failed;
        Truncated = truncated;
    }

    internal PowerShellNotebookFrameKind Kind { get; }

    internal NotebookOutput? Output { get; }

    internal bool Failed { get; }

    internal bool Truncated { get; }

    internal static PowerShellNotebookFrame ForOutput(NotebookOutput output) =>
        new(PowerShellNotebookFrameKind.Output, output, false, false);

    internal static PowerShellNotebookFrame TruncationNotice => new(PowerShellNotebookFrameKind.Truncated, null, false, true);

    internal static PowerShellNotebookFrame Completed(bool failed, bool truncated) =>
        new(PowerShellNotebookFrameKind.Complete, null, failed, truncated);
}

/// <summary>The worker stopped, closed or sent a frame the session cannot use.</summary>
internal sealed class PowerShellNotebookConnectionException : IOException {
    internal PowerShellNotebookConnectionException(string message, Exception? innerException = null) : base(message, innerException) { }
}
