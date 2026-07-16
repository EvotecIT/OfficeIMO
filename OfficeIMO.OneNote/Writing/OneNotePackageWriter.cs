namespace OfficeIMO.OneNote;

/// <summary>Creates portable Cabinet-based <c>.onepkg</c> notebook archives.</summary>
public static class OneNotePackageWriter {
    /// <summary>Serializes a complete notebook to a new <c>.onepkg</c> byte array.</summary>
    public static byte[] Write(OneNoteNotebook notebook, OneNoteWriterOptions? options = null) {
        OneNoteWriterOptions effective = options ?? new OneNoteWriterOptions();
        OneNoteNotebookSerializationPlan plan = OneNoteNotebookSerializationPlan.Create(notebook, effective);
        byte[] data = OneNoteCabinetArchiveWriter.Write(plan.Entries, effective.MaxOutputBytes);
        if (effective.ValidateRoundTrip) {
            using (var stream = new MemoryStream(data, false)) {
                OneNotePackageReader.Read(stream, "notebook.onepkg", new OneNoteNotebookReaderOptions {
                    ContinueOnSectionError = false,
                    IncludeRecycleBin = true,
                    MaxSectionGroupDepth = Math.Max(
                        OneNoteNotebookReaderOptions.DefaultMaxSectionGroupDepth,
                        effective.MaxSectionGroupDepth),
                    MaxPackageEntries = effective.MaxPackageEntries,
                    MaxPackageExpandedBytes = effective.MaxOutputBytes,
                    MaxPackageEntryBytes = effective.MaxOutputBytes,
                    OneNoteOptions = OneNoteWriterValidation.CreateReaderOptions(effective, effective.MaxOutputBytes)
                });
            }
        }
        return data;
    }

    /// <summary>Serializes a complete notebook to a caller-owned stream.</summary>
    public static void Write(OneNoteNotebook notebook, Stream destination, OneNoteWriterOptions? options = null) {
        if (destination == null) throw new ArgumentNullException(nameof(destination));
        if (!destination.CanWrite) throw new ArgumentException("The destination stream must be writable.", nameof(destination));
        byte[] data = Write(notebook, options);
        destination.Write(data, 0, data.Length);
    }

    /// <summary>Serializes a complete notebook to a <c>.onepkg</c> file.</summary>
    public static void Write(OneNoteNotebook notebook, string path, OneNoteWriterOptions? options = null) {
        if (path == null) throw new ArgumentNullException(nameof(path));
        byte[] data = Write(notebook, options);
        string fullPath = Path.GetFullPath(path);
        string? directory = Path.GetDirectoryName(fullPath);
        if (!string.IsNullOrEmpty(directory)) Directory.CreateDirectory(directory);
        File.WriteAllBytes(fullPath, data);
    }
}
