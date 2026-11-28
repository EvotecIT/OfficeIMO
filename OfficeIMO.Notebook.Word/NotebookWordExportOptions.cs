namespace OfficeIMO.Notebook.Word;

/// <summary>Controls the content of an editable Word report from saved notebook cells.</summary>
public sealed class NotebookWordExportOptions {
    /// <summary>Includes code inputs independently of the saved output.</summary>
    public bool IncludeCode { get; set; } = true;

    /// <summary>Includes saved outputs from code cells without executing the notebook.</summary>
    public bool IncludeOutputs { get; set; } = true;

    /// <summary>Includes the notebook's saved execution counts.</summary>
    public bool IncludeExecutionCounts { get; set; } = true;

    /// <summary>Optional visible report title.</summary>
    public string? Title { get; set; }

    /// <summary>Optional visible report author.</summary>
    public string? Author { get; set; }

    /// <summary>Rejects a report that would omit or approximate notebook content.</summary>
    public bool FailOnLoss { get; set; }

    /// <summary>Creates an independent snapshot of these settings.</summary>
    public NotebookWordExportOptions Clone() => (NotebookWordExportOptions)MemberwiseClone();
}
