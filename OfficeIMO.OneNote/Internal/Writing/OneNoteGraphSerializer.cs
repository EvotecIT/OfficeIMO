namespace OfficeIMO.OneNote;

/// <summary>Selects and emits the requested physical MS-ONESTORE encoding.</summary>
internal static class OneNoteGraphSerializer {
    internal static byte[] Write(
        OneNoteWriteGraph graph,
        OneNoteWriterOptions options,
        OneNoteStorageFormat sourceStorageFormat = OneNoteStorageFormat.Unknown,
        long? maxOutputBytes = null) {
        if (graph == null) throw new ArgumentNullException(nameof(graph));
        if (options == null) throw new ArgumentNullException(nameof(options));
        long outputLimit = maxOutputBytes ?? options.MaxOutputBytes;
        if (outputLimit < 1 || outputLimit > options.MaxOutputBytes) throw new ArgumentOutOfRangeException(nameof(maxOutputBytes));

        OneNoteStorageFormat target = options.StorageFormat != OneNoteStorageFormat.Unknown
            ? options.StorageFormat
            : sourceStorageFormat;
        if (target == OneNoteStorageFormat.Unknown) target = OneNoteStorageFormat.RevisionStore;

        byte[] data;
        switch (target) {
            case OneNoteStorageFormat.RevisionStore:
                data = OneNoteRevisionStoreWriter.Write(graph, outputLimit);
                break;
            case OneNoteStorageFormat.FileSynchronizationPackage:
                data = OneNotePackageStoreWriter.Write(graph, outputLimit);
                break;
            default:
                throw new ArgumentOutOfRangeException(
                    nameof(options),
                    "StorageFormat must be Unknown, RevisionStore, or FileSynchronizationPackage for .one and .onetoc2 output.");
        }

        if (data.LongLength > outputLimit) throw new IOException("OneNote output exceeds MaxOutputBytes.");
        return data;
    }
}
