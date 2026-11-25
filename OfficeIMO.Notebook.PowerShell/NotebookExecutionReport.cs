namespace OfficeIMO.Notebook.PowerShell;

/// <summary>The observed outcome of an attempted PowerShell notebook cell.</summary>
public enum NotebookCellExecutionStatus {
    /// <summary>The PowerShell pipeline completed without errors.</summary>
    Succeeded,

    /// <summary>The PowerShell pipeline reported an error.</summary>
    Failed,

    /// <summary>The cell exceeded its deadline and the session ended.</summary>
    TimedOut,

    /// <summary>The PowerShell worker ended unexpectedly.</summary>
    SessionEnded
}

/// <summary>Execution evidence for one attempted notebook code cell.</summary>
public sealed class NotebookCellExecutionResult {
    internal NotebookCellExecutionResult(int cellIndex, NotebookCell cell, int executionCount,
        NotebookCellExecutionStatus status, bool outputTruncated) {
        CellIndex = cellIndex;
        CellId = cell.Id;
        ExecutionCount = executionCount;
        Status = status;
        OutputTruncated = outputTruncated;
    }

    /// <summary>Gets the zero-based position of the cell in the notebook.</summary>
    public int CellIndex { get; }

    /// <summary>Gets the cell's notebook identifier.</summary>
    public string CellId { get; }

    /// <summary>Gets the execution count assigned by this session.</summary>
    public int ExecutionCount { get; }

    /// <summary>Gets the result of executing the cell.</summary>
    public NotebookCellExecutionStatus Status { get; }

    /// <summary>Gets whether the worker omitted output under the configured limits.</summary>
    public bool OutputTruncated { get; }
}

/// <summary>One output saved to a notebook cell during execution.</summary>
public sealed class NotebookExecutionEvent {
    internal NotebookExecutionEvent(int cellIndex, string cellId, int executionCount, NotebookOutput output) {
        CellIndex = cellIndex;
        CellId = cellId;
        ExecutionCount = executionCount;
        Output = output;
    }

    /// <summary>Gets the zero-based position of the cell producing output.</summary>
    public int CellIndex { get; }

    /// <summary>Gets the notebook identifier of the cell producing output.</summary>
    public string CellId { get; }

    /// <summary>Gets the current execution count of the producing cell.</summary>
    public int ExecutionCount { get; }

    /// <summary>Gets the output appended to the producing cell.</summary>
    public NotebookOutput Output { get; }
}

/// <summary>Reports an explicit notebook execution.</summary>
public sealed class NotebookExecutionReport {
    internal NotebookExecutionReport(NotebookDocument notebook, List<NotebookCellExecutionResult> cells, List<int> skippedCellIndices) {
        Notebook = notebook;
        Cells = cells.AsReadOnly();
        SkippedCellIndices = skippedCellIndices.AsReadOnly();
    }

    /// <summary>Gets the caller-owned notebook with the outputs of attempted cells.</summary>
    public NotebookDocument Notebook { get; }

    /// <summary>Gets the cells that ran, in notebook order.</summary>
    public IReadOnlyList<NotebookCellExecutionResult> Cells { get; }

    /// <summary>Gets the selected Markdown, raw, or other-language cells that were skipped.</summary>
    public IReadOnlyList<int> SkippedCellIndices { get; }

    /// <summary>Gets whether every attempted cell finished without a PowerShell error or interruption.</summary>
    public bool Succeeded => Cells.All(static cell => cell.Status == NotebookCellExecutionStatus.Succeeded);
}
