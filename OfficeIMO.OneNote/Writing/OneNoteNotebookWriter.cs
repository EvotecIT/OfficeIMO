namespace OfficeIMO.OneNote;

/// <summary>Creates a native offline notebook directory containing <c>.onetoc2</c> and <c>.one</c> files.</summary>
public static class OneNoteNotebookWriter {
    /// <summary>
    /// Writes a complete notebook to a new or empty directory. Existing files are never overwritten.
    /// </summary>
    public static void Write(OneNoteNotebook notebook, string directoryPath, OneNoteWriterOptions? options = null) {
        if (directoryPath == null) throw new ArgumentNullException(nameof(directoryPath));
        OneNoteWriterOptions effective = options ?? new OneNoteWriterOptions();
        OneNoteNotebookSerializationPlan plan = OneNoteNotebookSerializationPlan.Create(notebook, effective);
        string root = Path.GetFullPath(directoryPath);
        if (Directory.Exists(root) && Directory.EnumerateFileSystemEntries(root).Any()) {
            throw new IOException("The destination notebook directory is not empty.");
        }
        Directory.CreateDirectory(root);
        string rootPrefix = root.EndsWith(Path.DirectorySeparatorChar.ToString(), StringComparison.Ordinal) ? root : root + Path.DirectorySeparatorChar;
        foreach (OneNoteCabinetEntry entry in plan.Entries) {
            string relative = entry.Name.Replace('/', Path.DirectorySeparatorChar);
            string path = Path.GetFullPath(Path.Combine(root, relative));
            if (!path.StartsWith(rootPrefix, StringComparison.OrdinalIgnoreCase)) throw new IOException("A generated notebook entry escapes the destination directory.");
            string? parent = Path.GetDirectoryName(path);
            if (!string.IsNullOrEmpty(parent)) Directory.CreateDirectory(parent);
            using (var stream = new FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.None)) stream.Write(entry.Data, 0, entry.Data.Length);
        }
    }
}
