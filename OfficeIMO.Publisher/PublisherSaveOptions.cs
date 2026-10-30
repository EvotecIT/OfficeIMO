namespace OfficeIMO.Publisher;

/// <summary>Native source-save fidelity, output bounds and destination-conflict policy.</summary>
public sealed class PublisherSaveOptions {
    /// <summary>Blocks output with possible or unassessed loss by default. Allow still rejects unsupported edits and unsafe source carriers.</summary>
    public OfficeConversionLossPolicy LossPolicy { get; set; } = OfficeConversionLossPolicy.Block;
    /// <summary>Existing file destinations are preserved by default. Replace requires the shared atomic replacement path.</summary>
    public OfficeConversionFileConflictPolicy FileConflictPolicy { get; set; } = OfficeConversionFileConflictPolicy.FailIfExists;
    /// <summary>Maximum native output bytes, including the compound container. Defaults to 64 MiB.</summary>
    public long MaxOutputBytes { get; set; } = 64L * 1024 * 1024;

    /// <summary>Returns a validated independent settings snapshot.</summary>
    public PublisherSaveOptions Clone() {
        if (!Enum.IsDefined(typeof(OfficeConversionLossPolicy), LossPolicy)) throw new ArgumentOutOfRangeException(nameof(LossPolicy));
        if (!Enum.IsDefined(typeof(OfficeConversionFileConflictPolicy), FileConflictPolicy)) throw new ArgumentOutOfRangeException(nameof(FileConflictPolicy));
        if (MaxOutputBytes < 1 || MaxOutputBytes > int.MaxValue) throw new ArgumentOutOfRangeException(nameof(MaxOutputBytes));
        return new PublisherSaveOptions { LossPolicy = LossPolicy, FileConflictPolicy = FileConflictPolicy, MaxOutputBytes = MaxOutputBytes };
    }
}
