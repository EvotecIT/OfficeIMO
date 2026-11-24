namespace OfficeIMO.Notebook.Import;

public static partial class NotebookImporter {
    /// <summary>Recognizes notebook-specific source extensions without opening the file.</summary>
    /// <remarks>Use an explicit profile for ambiguous names such as .dib, .md, .json, and script files.
    /// Open native .ipynb through <see cref="NotebookDocument"/>.</remarks>
    public static bool TryGetFormatFromPath(string path, out NotebookImportFormat format) {
        ArgumentNullException.ThrowIfNull(path);
        NotebookImportFormat? match = path.EndsWith(".nb.html", StringComparison.OrdinalIgnoreCase)
            ? NotebookImportFormat.RNotebookHtml
            : Path.GetExtension(path).ToLowerInvariant() switch {
                ".qmd" => NotebookImportFormat.Quarto,
                ".rmd" => NotebookImportFormat.RMarkdown,
                ".deepnote" => NotebookImportFormat.Deepnote,
                ".zpln" => NotebookImportFormat.Zeppelin,
                _ => null
            };
        format = match.GetValueOrDefault();
        return match.HasValue;
    }
}
