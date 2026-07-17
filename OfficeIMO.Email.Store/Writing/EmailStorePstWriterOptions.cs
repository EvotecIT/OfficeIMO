namespace OfficeIMO.Email.Store;

/// <summary>Controls creation of a new managed Unicode PST file.</summary>
public sealed class EmailStorePstWriterOptions {
    /// <summary>Creates PST writer options.</summary>
    public EmailStorePstWriterOptions(
        string? displayName = null,
        bool overwriteExisting = false,
        bool failOnDataLoss = false,
        int maxFolderCount = 100_000,
        int maxItemCount = int.MaxValue,
        int maxNestedMessageDepth = 32) {
        if (maxFolderCount <= 0) throw new ArgumentOutOfRangeException(nameof(maxFolderCount));
        if (maxItemCount <= 0) throw new ArgumentOutOfRangeException(nameof(maxItemCount));
        if (maxNestedMessageDepth < 0) throw new ArgumentOutOfRangeException(nameof(maxNestedMessageDepth));
        DisplayName = string.IsNullOrWhiteSpace(displayName) ? "OfficeIMO Personal Folders" : displayName!;
        OverwriteExisting = overwriteExisting;
        FailOnDataLoss = failOnDataLoss;
        MaxFolderCount = maxFolderCount;
        MaxItemCount = maxItemCount;
        MaxNestedMessageDepth = maxNestedMessageDepth;
    }

    /// <summary>Display name stored in the message-store object.</summary>
    public string DisplayName { get; }

    /// <summary>Whether an existing destination may be atomically replaced.</summary>
    public bool OverwriteExisting { get; }

    /// <summary>Whether a fidelity warning prevents completion.</summary>
    public bool FailOnDataLoss { get; }

    /// <summary>Maximum folders accepted by one writer.</summary>
    public int MaxFolderCount { get; }

    /// <summary>Maximum top-level items accepted by one writer.</summary>
    public int MaxItemCount { get; }

    /// <summary>Maximum embedded-message nesting depth written.</summary>
    public int MaxNestedMessageDepth { get; }
}
