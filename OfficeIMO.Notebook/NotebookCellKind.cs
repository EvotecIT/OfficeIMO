namespace OfficeIMO.Notebook;

/// <summary>The content kind of a notebook cell.</summary>
public enum NotebookCellKind {
    /// <summary>A cell type outside the recognized notebook profiles.</summary>
    Unknown,
    /// <summary>A cell containing Markdown prose.</summary>
    Markdown,
    /// <summary>A cell containing source code and optional saved outputs.</summary>
    Code,
    /// <summary>A raw cell for a target-specific exporter.</summary>
    Raw
}
