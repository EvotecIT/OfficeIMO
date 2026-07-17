namespace OfficeIMO.OneNote;

/// <summary>Loads a desktop <c>.one</c> section into the typed OfficeIMO.OneNote model.</summary>
public static class OneNoteSectionReader {
    /// <summary>Reads a section from a file path.</summary>
    public static OneNoteSection Read(string path, OneNoteReaderOptions? options = null) {
        if (path == null) throw new ArgumentNullException(nameof(path));
        OneNoteRevisionStore store = OneNoteRevisionStoreReader.Read(path, options);
        OneNoteSection section = OneNoteSemanticMapper.MapSection(store, options ?? new OneNoteReaderOptions());
        section.SourcePath = Path.GetFullPath(path);
        if (string.IsNullOrWhiteSpace(section.Name)) section.Name = Path.GetFileNameWithoutExtension(path);
        return section;
    }

    /// <summary>Reads a section from a caller-owned seekable stream.</summary>
    public static OneNoteSection Read(Stream stream, OneNoteReaderOptions? options = null) {
        OneNoteRevisionStore store = OneNoteRevisionStoreReader.Read(stream, options);
        return OneNoteSemanticMapper.MapSection(store, options ?? new OneNoteReaderOptions());
    }
}
