namespace OfficeIMO.Notebook;

/// <summary>The kind of an output saved in a code cell.</summary>
public enum NotebookOutputKind {
    /// <summary>An unrecognized or extension output.</summary>
    Unknown,
    /// <summary>Text captured from a named stream.</summary>
    Stream,
    /// <summary>A displayed bundle of alternative MIME representations.</summary>
    DisplayData,
    /// <summary>A displayed execution result with an execution count.</summary>
    ExecuteResult,
    /// <summary>A saved execution error and traceback.</summary>
    Error
}
