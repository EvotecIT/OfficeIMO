namespace OfficeIMO.Notebook.Import;

/// <summary>Controls resource limits and caller-selected inputs for notebook import.</summary>
public sealed class NotebookImportOptions {
    /// <summary>Resource limits for the individual input and the imported notebook.</summary>
    public NotebookLoadOptions Limits { get; set; } = NotebookLoadOptions.UntrustedDefaults;

    /// <summary>Language used for Jupytext comment markers. When omitted, the importer uses the script metadata or its default.</summary>
    public string? ScriptLanguage { get; set; }

    /// <summary>Optional caller-owned Jupyter companion for Jupytext. The import verifies cell alignment before using saved results.</summary>
    public NotebookDocument? PairedNotebook { get; set; }

    /// <summary>Exact notebook identifier when selecting a notebook from a Deepnote project.</summary>
    public string? DeepnoteNotebookId { get; set; }

    internal NotebookImportOptions Snapshot(NotebookImportFormat format) {
        ArgumentNullException.ThrowIfNull(Limits);
        if (!Enum.IsDefined(typeof(NotebookImportFormat), format)) throw new ArgumentOutOfRangeException(nameof(format));
        if (format != NotebookImportFormat.JupytextPercent && (ScriptLanguage != null || PairedNotebook != null)) {
            throw new ArgumentException("A script language and paired notebook require the Jupytext percent profile.");
        }
        if (format != NotebookImportFormat.Deepnote && DeepnoteNotebookId != null) {
            throw new ArgumentException("A Deepnote notebook ID requires the Deepnote profile.");
        }
        var copy = (NotebookImportOptions)MemberwiseClone();
        copy.Limits = Limits.Snapshot();
        return copy;
    }
}
