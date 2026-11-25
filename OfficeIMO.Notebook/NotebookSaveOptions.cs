namespace OfficeIMO.Notebook;

/// <summary>Controls how a notebook is serialized as nbformat 4.</summary>
public sealed class NotebookSaveOptions {
    /// <summary>Gets or sets whether JSON is indented for review. The default is true.</summary>
    public bool Indented { get; set; } = true;

    /// <summary>Gets or sets whether to allow reported normalization losses, including version 3 worksheet flattening.</summary>
    public bool AllowLoss { get; set; }
}
