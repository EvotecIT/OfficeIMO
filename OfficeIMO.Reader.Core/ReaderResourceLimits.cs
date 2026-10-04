namespace OfficeIMO.Reader;

/// <summary>
/// Whole-read resource budgets shared by nested readers and by files in a folder operation.
/// Null values leave a budget unlimited. These complement format parser limits; rich-model
/// budgets are checked when the engine produces its model, not before its internal allocation.
/// </summary>
public sealed class ReaderResourceLimits {
    /// <summary>Maximum distinct generated chunk objects, including nested content.</summary>
    public long? MaxChunks { get; set; }
    /// <summary>Maximum aggregate characters in chunk Text and Markdown projections.</summary>
    public long? MaxChunkCharacters { get; set; }
    /// <summary>Maximum distinct rich block objects across root and nested results.</summary>
    public long? MaxBlocks { get; set; }
    /// <summary>Maximum distinct asset records across root and nested results.</summary>
    public long? MaxAssets { get; set; }
    /// <summary>Maximum aggregate known asset payload bytes. Shared byte-array payloads count once.</summary>
    public long? MaxAssetBytes { get; set; }
    /// <summary>Maximum aggregate decoded bytes admitted to nested readers, including nested archive bytes.</summary>
    public long? MaxNestedInputBytes { get; set; }
    /// <summary>Maximum nested inputs decoded or dispatched across one operation.</summary>
    public long? MaxNestedDocuments { get; set; }
    /// <summary>Maximum nested input depth, with the root at zero.</summary>
    public int? MaxNestedDepth { get; set; }

    internal ReaderResourceLimits CloneValidated() {
        var clone = (ReaderResourceLimits)MemberwiseClone();
        if (MaxChunks < 0 || MaxChunkCharacters < 0 || MaxBlocks < 0 || MaxAssets < 0 || MaxAssetBytes < 0 ||
            MaxNestedInputBytes < 0 || MaxNestedDocuments < 0 || MaxNestedDepth < 0) {
            throw new ArgumentOutOfRangeException(nameof(ReaderResourceLimits), "Resource budgets cannot be negative.");
        }
        return clone;
    }
}

/// <summary>A whole-operation budget was exceeded. This failure must not be downgraded to a recoverable parse warning.</summary>
public sealed class ReaderResourceLimitException : IOException {
    /// <summary>Creates a named resource-budget failure.</summary>
    public ReaderResourceLimitException(string limitName, long maximum)
        : base($"Reader operation exceeds {limitName} ({maximum}).") { LimitName = limitName; Maximum = maximum; }
    /// <summary>The exceeded limit property.</summary>
    public string LimitName { get; }
    /// <summary>The configured budget.</summary>
    public long Maximum { get; }
}
