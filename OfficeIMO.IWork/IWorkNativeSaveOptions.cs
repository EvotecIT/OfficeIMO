namespace OfficeIMO.IWork;

/// <summary>Operation limits and explicit cache acceptance for saving a loaded native iWork snapshot.</summary>
public sealed class IWorkNativeSaveOptions {
    /// <summary>Gets or sets the maximum sum of captured file-entry bytes and the maximum encoded ZIP length, each applied separately.</summary>
    public int MaximumPackageBytes { get; set; } = 64 * 1024 * 1024;
    /// <summary>Gets or sets the combined decompressed bytes of archives changed by this operation.</summary>
    public int MaximumEditedArchiveBytes { get; set; } = 64 * 1024 * 1024;
    /// <summary>Gets or sets the number of distinct edited source records permitted in one save.</summary>
    public int MaximumEditedRecords { get; set; } = 1_024;
    /// <summary>Gets or sets explicit acceptance that edited packages retain native previews and other visual caches without refreshing them.</summary>
    /// <remarks>Every edited save requires this acceptance and reports the cache boundary as unassessed.
    /// Unchanged saves preserve existing caches and do not require it.</remarks>
    public bool AllowUnassessedVisualCaches { get; set; }

    internal IWorkNativeSaveOptions Snapshot() {
        var copy = new IWorkNativeSaveOptions { MaximumPackageBytes = MaximumPackageBytes,
            MaximumEditedArchiveBytes = MaximumEditedArchiveBytes, MaximumEditedRecords = MaximumEditedRecords,
            AllowUnassessedVisualCaches = AllowUnassessedVisualCaches };
        if (copy.MaximumPackageBytes <= 0) throw new ArgumentOutOfRangeException(nameof(MaximumPackageBytes));
        if (copy.MaximumEditedArchiveBytes <= 0) throw new ArgumentOutOfRangeException(nameof(MaximumEditedArchiveBytes));
        if (copy.MaximumEditedRecords <= 0) throw new ArgumentOutOfRangeException(nameof(MaximumEditedRecords));
        return copy;
    }
}
