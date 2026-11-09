namespace OfficeIMO.OneNote;

internal static class OneNoteWriterValidation {
    internal static void ValidateSectionOptions(OneNoteSaveOptions options) {
        if (options.GeneratedTimestampUtc.HasValue && options.GeneratedTimestampUtc.Value.Kind != DateTimeKind.Utc)
            throw new ArgumentException("GeneratedTimestampUtc must have UTC kind.", nameof(options.GeneratedTimestampUtc));
        if (options.IdentitySeed.HasValue && !options.GeneratedTimestampUtc.HasValue)
            throw new ArgumentException("Reproducible identity generation requires an explicit GeneratedTimestampUtc.", nameof(options.GeneratedTimestampUtc));
        ArgumentNullException.ThrowIfNull(options);
        if (options.MaxOutputBytes < 1) {
            throw new ArgumentOutOfRangeException(nameof(options.MaxOutputBytes), "MaxOutputBytes must be greater than zero.");
        }
        if (options.MaxInkPathValues < 1) {
            throw new ArgumentOutOfRangeException(nameof(options.MaxInkPathValues), "MaxInkPathValues must be greater than zero.");
        }
        ValidateTraversalDepth(options.MaxPageRelationshipDepth, nameof(options.MaxPageRelationshipDepth));
        ValidateTraversalDepth(options.MaxContentDepth, nameof(options.MaxContentDepth));
    }

    internal static void ValidateNotebookOptions(OneNoteSaveOptions options) {
        ValidateSectionOptions(options);
        ValidateTraversalDepth(options.MaxSectionGroupDepth, nameof(options.MaxSectionGroupDepth));
        if (options.MaxPackageEntries < 1 || options.MaxPackageEntries > ushort.MaxValue) {
            throw new ArgumentOutOfRangeException(nameof(options.MaxPackageEntries), "MaxPackageEntries must be between 1 and 65535.");
        }
    }

    /// <summary>
    /// Creates read-back limits that accept any asset set already bounded by the serialized output limit.
    /// </summary>
    internal static OneNoteLoadOptions CreateReaderOptions(OneNoteSaveOptions options, long maxOutputBytes) => new OneNoteLoadOptions {
        MaxInputBytes = maxOutputBytes,
        MaxAssetBytes = maxOutputBytes,
        MaxTotalAssetBytes = maxOutputBytes,
        MaxInkPathValues = options.MaxInkPathValues,
        MaxPageRelationshipDepth = Math.Max(
            OneNoteLoadOptions.DefaultMaxPageRelationshipDepth,
            options.MaxPageRelationshipDepth),
        MaxPropertySetDepth = Math.Max(
            OneNoteLoadOptions.DefaultMaxPropertySetDepth,
            options.MaxContentDepth)
    };

    private static void ValidateTraversalDepth(int value, string parameterName) {
        if (value < 1 || value > OneNoteSaveOptions.MaximumTraversalDepth) {
            throw new ArgumentOutOfRangeException(
                parameterName,
                "Writer traversal depths must be between 1 and " + OneNoteSaveOptions.MaximumTraversalDepth + ".");
        }
    }
}
