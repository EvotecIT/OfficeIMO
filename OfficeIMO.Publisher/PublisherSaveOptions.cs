namespace OfficeIMO.Publisher;

/// <summary>Native source-save fidelity, output bounds and destination-conflict policy.</summary>
public sealed class PublisherSaveOptions {
    /// <summary>Allows prepared output with possible or unassessed loss. Unsupported edits and unsafe source carriers always fail.</summary>
    public bool AllowLoss { get; set; }
    /// <summary>Existing file destinations are preserved by default. Replace requires the shared atomic replacement path.</summary>
    public OfficeConversionFileConflictPolicy FileConflictPolicy { get; set; } = OfficeConversionFileConflictPolicy.FailIfExists;
    /// <summary>Maximum native output bytes, including the compound container. Defaults to 64 MiB.</summary>
    public long MaxOutputBytes { get; set; } = 64L * 1024 * 1024;

    /// <summary>Returns a validated independent settings snapshot.</summary>
    internal PublisherSaveOptions Clone() {
        if (!Enum.IsDefined(typeof(OfficeConversionFileConflictPolicy), FileConflictPolicy)) throw new ArgumentOutOfRangeException(nameof(FileConflictPolicy));
        if (MaxOutputBytes < 1 || MaxOutputBytes > int.MaxValue) throw new ArgumentOutOfRangeException(nameof(MaxOutputBytes));
        return new PublisherSaveOptions { AllowLoss = AllowLoss, FileConflictPolicy = FileConflictPolicy, MaxOutputBytes = MaxOutputBytes };
    }
}
