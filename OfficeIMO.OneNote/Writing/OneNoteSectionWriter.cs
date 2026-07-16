namespace OfficeIMO.OneNote;

/// <summary>Creates native desktop or FSSHTTP package-store <c>.one</c> section files.</summary>
public static class OneNoteSectionWriter {
    /// <summary>Serializes a section to a new byte array.</summary>
    public static byte[] Write(OneNoteSection section, OneNoteWriterOptions? options = null) {
        if (section == null) throw new ArgumentNullException(nameof(section));
        OneNoteWriterOptions effective = options ?? new OneNoteWriterOptions();
        OneNoteWriterValidation.ValidateSectionOptions(effective);
        OneNoteWriteGraph graph = new OneNoteWriteGraphBuilder(
            effective.MaxOutputBytes,
            effective.PreserveUnknownData,
            effective.MaxPageRelationshipDepth,
            effective.MaxContentDepth).BuildSection(section);
        byte[] data = OneNoteGraphSerializer.Write(graph, effective, section.StorageFormat);
        if (effective.ValidateRoundTrip) {
            using (var stream = new MemoryStream(data, false)) {
                OneNoteSection roundTrip = OneNoteSectionReader.Read(
                    stream,
                    OneNoteWriterValidation.CreateReaderOptions(effective, effective.MaxOutputBytes));
                OneNoteWriteRoundTripValidator.ValidateSection(section, roundTrip);
            }
        }
        return data;
    }

    /// <summary>Serializes a section to a caller-owned writable stream.</summary>
    public static void Write(OneNoteSection section, Stream destination, OneNoteWriterOptions? options = null) {
        if (destination == null) throw new ArgumentNullException(nameof(destination));
        if (!destination.CanWrite) throw new ArgumentException("The destination stream must be writable.", nameof(destination));
        byte[] data = Write(section, options);
        destination.Write(data, 0, data.Length);
    }

    /// <summary>Serializes a section to a file, replacing an existing file.</summary>
    public static void Write(OneNoteSection section, string path, OneNoteWriterOptions? options = null) {
        if (path == null) throw new ArgumentNullException(nameof(path));
        if (path.Length == 0) throw new ArgumentException("Path cannot be empty.", nameof(path));
        byte[] data = Write(section, options);
        string fullPath = Path.GetFullPath(path);
        string? directory = Path.GetDirectoryName(fullPath);
        if (!string.IsNullOrEmpty(directory)) Directory.CreateDirectory(directory);
        File.WriteAllBytes(fullPath, data);
    }
}
