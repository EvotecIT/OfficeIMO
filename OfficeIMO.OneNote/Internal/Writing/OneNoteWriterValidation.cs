namespace OfficeIMO.OneNote;

internal static class OneNoteWriterValidation {
    /// <summary>
    /// Creates read-back limits that accept any asset set already bounded by the serialized output limit.
    /// </summary>
    internal static OneNoteReaderOptions CreateReaderOptions(long maxOutputBytes) => new OneNoteReaderOptions {
        MaxInputBytes = maxOutputBytes,
        MaxAssetBytes = maxOutputBytes,
        MaxTotalAssetBytes = maxOutputBytes
    };
}
