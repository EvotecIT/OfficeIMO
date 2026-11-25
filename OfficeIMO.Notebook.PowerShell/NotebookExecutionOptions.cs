namespace OfficeIMO.Notebook.PowerShell;

/// <summary>Controls which notebook cells run and how their output is retained.</summary>
public sealed class NotebookExecutionOptions {
    /// <summary>Gets or sets the zero-based notebook cell indices. Selected cells run in notebook order. The default selects every cell.</summary>
    public IReadOnlyCollection<int>? CellIndices { get; set; }

    /// <summary>Gets or sets the deadline for each cell. Set to <see langword="null"/> to disable the deadline.</summary>
    public TimeSpan? CellTimeout { get; set; } = TimeSpan.FromSeconds(60);

    /// <summary>Gets or sets whether the session continues to later cells after a PowerShell error.</summary>
    public bool ContinueOnError { get; set; }

    /// <summary>Gets or sets whether to run code cells whose language is unspecified in both the cell and notebook.</summary>
    public bool IncludeUnspecifiedLanguage { get; set; }

    /// <summary>Gets or sets the maximum number of UTF-16 characters retained from each cell, including error text.</summary>
    public int OutputCharactersLimit { get; set; } = 1_048_576;

    /// <summary>Gets or sets the maximum number of output records retained from each cell.</summary>
    public int OutputRecordsLimit { get; set; } = 4_096;

    /// <summary>Gets or sets a synchronous callback for each output from the PowerShell worker as it arrives. The callback may run on a background thread.</summary>
    /// <remarks>Callbacks run in output order. They should return promptly and must not start another execution on this session.</remarks>
    public Action<NotebookExecutionEvent>? OutputReceived { get; set; }
}
