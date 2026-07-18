namespace OfficeIMO.OneNote;

internal static class OneNoteWriterValidation {
    internal static void ValidateSectionOptions(OneNoteWriterOptions options) {
        if (options == null) throw new ArgumentNullException(nameof(options));
        if (options.MaxOutputBytes < 1) {
            throw new ArgumentOutOfRangeException(nameof(options.MaxOutputBytes), "MaxOutputBytes must be greater than zero.");
        }
        if (options.MaxInkPathValues < 1) {
            throw new ArgumentOutOfRangeException(nameof(options.MaxInkPathValues), "MaxInkPathValues must be greater than zero.");
        }
        ValidateTraversalDepth(options.MaxPageRelationshipDepth, nameof(options.MaxPageRelationshipDepth));
        ValidateTraversalDepth(options.MaxContentDepth, nameof(options.MaxContentDepth));
    }

    internal static void ValidateNotebookOptions(OneNoteWriterOptions options) {
        ValidateSectionOptions(options);
        ValidateTraversalDepth(options.MaxSectionGroupDepth, nameof(options.MaxSectionGroupDepth));
        if (options.MaxPackageEntries < 1 || options.MaxPackageEntries > ushort.MaxValue) {
            throw new ArgumentOutOfRangeException(nameof(options.MaxPackageEntries), "MaxPackageEntries must be between 1 and 65535.");
        }
    }

    /// <summary>
    /// Creates read-back limits that accept any asset set already bounded by the serialized output limit.
    /// </summary>
    internal static OneNoteReaderOptions CreateReaderOptions(OneNoteWriterOptions options, long maxOutputBytes) => new OneNoteReaderOptions {
        MaxInputBytes = maxOutputBytes,
        MaxAssetBytes = maxOutputBytes,
        MaxTotalAssetBytes = maxOutputBytes,
        MaxInkPathValues = options.MaxInkPathValues,
        MaxPageRelationshipDepth = Math.Max(
            OneNoteReaderOptions.DefaultMaxPageRelationshipDepth,
            options.MaxPageRelationshipDepth),
        MaxPropertySetDepth = Math.Max(
            OneNoteReaderOptions.DefaultMaxPropertySetDepth,
            options.MaxContentDepth)
    };

    private static void ValidateTraversalDepth(int value, string parameterName) {
        if (value < 1 || value > OneNoteWriterOptions.MaximumTraversalDepth) {
            throw new ArgumentOutOfRangeException(
                parameterName,
                "Writer traversal depths must be between 1 and " + OneNoteWriterOptions.MaximumTraversalDepth + ".");
        }
    }
}
