namespace OfficeIMO.Workflows;

/// <summary>The operation represented by an ONIX message. This does not establish recipient acceptance.</summary>
public enum BookOnixMessageKind {
    /// <summary>Complete records with their original notification types.</summary>
    CompleteRecords,
    /// <summary>Explicit whole-block replacements or clears, notification 04.</summary>
    BlockUpdates,
    /// <summary>Explicit withdrawal of metadata records, notification 05; not product cancellation.</summary>
    Deletions
}

/// <summary>ONIX blocks in their schema serialization order.</summary>
public enum BookOnixBlock {
    /// <summary>Product description, block 1; cannot be cleared.</summary>
    DescriptiveDetail,
    /// <summary>Collateral, block 2.</summary>
    CollateralDetail,
    /// <summary>Promotional events, block 7.</summary>
    PromotionDetail,
    /// <summary>Detailed contents, block 3.</summary>
    ContentDetail,
    /// <summary>Publisher and rights, block 4; cannot be cleared.</summary>
    PublishingDetail,
    /// <summary>Related products and works, block 5.</summary>
    RelatedMaterial,
    /// <summary>Production manifests, block 8.</summary>
    ProductionDetail,
    /// <summary>All unnamed product supply repeats, block 6; cannot be cleared. Named markets use explicit reference selections.</summary>
    ProductSupply
}

/// <summary>An explicit block-update instruction derived from a complete, unaltered export result.</summary>
public sealed record BookOnixBlockUpdate(BookOnixExportResult Source) {
    /// <summary>Market references to copy whole from Source, in selection order. Every source supply must be named. Cannot accompany a whole ProductSupply replacement.</summary>
    public IReadOnlyList<string> ReplaceMarketReferences { get; init; } = [];
    /// <summary>Explicit market removals, emitted as reference-only supply blocks. References may be absent from Source; recipient existence is not checked. At most 32 market operations in total.</summary>
    public IReadOnlyList<string> RemoveMarketReferences { get; init; } = [];
    /// <summary>Distinct blocks to copy in full from Source. Missing blocks are errors, not inferred deletions.</summary>
    public IReadOnlyList<BookOnixBlock> ReplaceBlocks { get; init; } = [];
    /// <summary>Distinct optional blocks to emit empty. DescriptiveDetail, PublishingDetail and ProductSupply cannot be cleared.</summary>
    public IReadOnlyList<BookOnixBlock> ClearBlocks { get; init; } = [];
}

/// <summary>A plain-text reason for withdrawing a metadata record, optionally tagged with an ONIX list 74 language.</summary>
public sealed record BookOnixDeletionReason(string Text, string? LanguageCode = null);

/// <summary>Explicit withdrawal of a metadata record issued in error; does not describe a book going out of print or being withdrawn from sale.</summary>
public sealed record BookOnixDeletion(BookOnixExportResult Source) {
    /// <summary>Up to 16 optional translations, at most 100 UTF-16 code units each. Repeated reasons require distinct explicit languages.</summary>
    public IReadOnlyList<BookOnixDeletionReason> Reasons { get; init; } = [];
}
