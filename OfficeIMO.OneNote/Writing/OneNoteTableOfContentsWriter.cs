namespace OfficeIMO.OneNote;

/// <summary>Creates native offline <c>.onetoc2</c> table-of-contents files.</summary>
public static class OneNoteTableOfContentsWriter {
    /// <summary>Serializes the root notebook hierarchy to a new byte array.</summary>
    public static byte[] Write(OneNoteNotebook notebook, OneNoteWriterOptions? options = null) =>
        OneNoteNotebookSerializationPlan.CreateRootTableOfContents(notebook, options ?? new OneNoteWriterOptions());

    /// <summary>Serializes the root notebook hierarchy to a caller-owned stream.</summary>
    public static void Write(OneNoteNotebook notebook, Stream destination, OneNoteWriterOptions? options = null) {
        if (destination == null) throw new ArgumentNullException(nameof(destination));
        if (!destination.CanWrite) throw new ArgumentException("The destination stream must be writable.", nameof(destination));
        byte[] data = Write(notebook, options);
        destination.Write(data, 0, data.Length);
    }

    /// <summary>Serializes the root notebook hierarchy to a file.</summary>
    public static void Write(OneNoteNotebook notebook, string path, OneNoteWriterOptions? options = null) {
        if (path == null) throw new ArgumentNullException(nameof(path));
        byte[] data = Write(notebook, options);
        string fullPath = Path.GetFullPath(path);
        string? directory = Path.GetDirectoryName(fullPath);
        if (!string.IsNullOrEmpty(directory)) Directory.CreateDirectory(directory);
        File.WriteAllBytes(fullPath, data);
    }
}
