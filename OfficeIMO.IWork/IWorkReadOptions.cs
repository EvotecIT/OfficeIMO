namespace OfficeIMO.IWork;

/// <summary>Bounds and preservation preferences applied while reading an iWork source.</summary>
public sealed class IWorkReadOptions {
    /// <summary>Gets or sets whether raw source records are included in the conversion report for inspection.</summary>
    public bool PreserveSourceRecords { get; set; } = true;

    /// <summary>Gets or sets the maximum source package size in bytes.</summary>
    public long MaximumPackageBytes { get; set; } = 512L * 1024 * 1024;

    /// <summary>Gets or sets the maximum number of outer and nested package entries.</summary>
    public int MaximumEntryCount { get; set; } = 8192;

    /// <summary>Gets or sets the maximum uncompressed size of one package entry.</summary>
    public long MaximumEntryBytes { get; set; } = 128L * 1024 * 1024;

    /// <summary>Gets or sets the maximum combined uncompressed package-entry size.</summary>
    public long MaximumTotalEntryBytes { get; set; } = 512L * 1024 * 1024;

    /// <summary>Gets or sets the maximum compressed size of one IWA archive.</summary>
    public int MaximumIwaBytes { get; set; } = 64 * 1024 * 1024;

    /// <summary>Gets or sets the maximum decompressed size of one IWA archive.</summary>
    public int MaximumDecompressedIwaBytes { get; set; } = 256 * 1024 * 1024;

    /// <summary>Gets or sets the maximum combined decompressed size of all IWA archives in one source.</summary>
    public long MaximumTotalDecompressedIwaBytes { get; set; } = 512L * 1024 * 1024;

    /// <summary>Gets or sets the maximum declared decompressed size of one raw Snappy chunk.</summary>
    public int MaximumSnappyChunkBytes { get; set; } = 64 * 1024 * 1024;

    /// <summary>Gets or sets the maximum encoded ArchiveInfo header size.</summary>
    public int MaximumArchiveInfoBytes { get; set; } = 1024 * 1024;

    /// <summary>Gets or sets the maximum payload size of one archived record.</summary>
    public int MaximumRecordBytes { get; set; } = 128 * 1024 * 1024;

    /// <summary>Gets or sets the maximum number of IWA records across the package.</summary>
    public int MaximumRecordCount { get; set; } = 1_000_000;

    /// <summary>Gets or sets the maximum combined ArchiveInfo version, object-reference, and data-reference values across the package.</summary>
    public int MaximumArchiveReferenceCount { get; set; } = 8_000_000;

    /// <summary>Gets or sets the maximum declared reference occurrences inspected in distinct source fields for unresolved-reference evidence across one semantic projection, including readable siblings in a field with failures.</summary>
    public int MaximumSourceReferenceIssues { get; set; } = 100_000;

    /// <summary>Gets or sets the maximum distinct selected source paths retained as unreadable or rejected declaration evidence across one semantic projection.</summary>
    public int MaximumSourceDeclarationIssues { get; set; } = 100_000;

    /// <summary>Gets or sets the maximum number of fields decoded from one protobuf message.</summary>
    public int MaximumProtobufFieldCount { get; set; } = 1_000_000;

    /// <summary>Gets or sets the maximum nested protobuf message depth.</summary>
    public int MaximumProtobufDepth { get; set; } = 64;

    /// <summary>Gets or sets the maximum projected row count of one iWork table.</summary>
    public int MaximumTableRows { get; set; } = 1_048_576;

    /// <summary>Gets or sets the maximum projected column count of one iWork table.</summary>
    public int MaximumTableColumns { get; set; } = 16_384;

    /// <summary>Gets or sets the maximum number of materialized non-empty cells across one iWork source.</summary>
    public int MaximumMaterializedCells { get; set; } = 10_000_000;

    /// <summary>Gets or sets the maximum combined string and formula catalog entries decoded across one iWork source.</summary>
    public int MaximumTableCatalogEntries { get; set; } = 100_000;

    /// <summary>Gets or sets the maximum combined table dimension headers, bucket references, hidden-state declarations and filter references inspected across one iWork source.</summary>
    public int MaximumTableDimensionEntries { get; set; } = 1_100_000;

    /// <summary>Gets or sets the maximum number of projected sheets across one Numbers source.</summary>
    public int MaximumProjectedSheets { get; set; } = 4096;

    /// <summary>Gets or sets the maximum number of projected slides across one Keynote source.</summary>
    public int MaximumProjectedSlides { get; set; } = 10_000;

    /// <summary>Gets or sets the maximum number of projected tables across one iWork source.</summary>
    public int MaximumProjectedTables { get; set; } = 4096;

    /// <summary>Gets or sets the maximum number of projected images across one iWork source.</summary>
    public int MaximumProjectedImages { get; set; } = 4096;

    /// <summary>Gets or sets the maximum number of package image-metadata entries decoded across one iWork source.</summary>
    public int MaximumImageMetadataEntries { get; set; } = 16_384;

    /// <summary>Gets or sets the maximum combined encoded image bytes emitted by one semantic projection, counting repeated uses.</summary>
    public long MaximumProjectedImageBytes { get; set; } = 512L * 1024 * 1024;

    /// <summary>Gets or sets the maximum cumulative decoded image work for preview inspection and for each semantic projection. Default: 64 MiB.</summary>
    public long MaximumDecodedImageBytes { get; set; } = 64L * 1024 * 1024;

    /// <summary>Gets or sets the maximum number of merged ranges projected from one iWork table.</summary>
    public int MaximumTableMergedRanges { get; set; } = 100_000;

    /// <summary>Gets or sets the maximum number of syntax nodes decoded from one iWork formula.</summary>
    public int MaximumFormulaNodes { get; set; } = 4096;

    /// <summary>Gets or sets the maximum reconstructed formula length in characters.</summary>
    public int MaximumFormulaCharacters { get; set; } = 8192;

    /// <summary>Gets or sets the maximum combined formula-rendering work across one semantic projection, counting repeated formula uses.</summary>
    public long MaximumFormulaRenderingOperations { get; set; } = 64L * 1024 * 1024;

    /// <summary>Gets or sets the maximum number of projected text items across one semantic projection.</summary>
    public int MaximumProjectedTextItems { get; set; } = 100_000;

    /// <summary>Gets or sets the maximum number of decoded text characters across one semantic projection.</summary>
    public long MaximumProjectedTextCharacters { get; set; } = 16L * 1024 * 1024;

    /// <summary>Gets or sets the maximum cross-record inheritance depth of an iWork text or table style.</summary>
    public int MaximumTextStyleInheritanceDepth { get; set; } = 64;

    /// <summary>Returns a validated copy suitable for an isolated reader registration.</summary>
    public IWorkReadOptions Clone() => Snapshot();

    internal IWorkReadOptions Snapshot() {
        ValidatePositive(MaximumPackageBytes, nameof(MaximumPackageBytes));
        ValidatePositive(MaximumEntryCount, nameof(MaximumEntryCount));
        ValidatePositive(MaximumEntryBytes, nameof(MaximumEntryBytes));
        ValidatePositive(MaximumTotalEntryBytes, nameof(MaximumTotalEntryBytes));
        ValidatePositive(MaximumIwaBytes, nameof(MaximumIwaBytes));
        ValidatePositive(MaximumDecompressedIwaBytes, nameof(MaximumDecompressedIwaBytes));
        ValidatePositive(MaximumTotalDecompressedIwaBytes, nameof(MaximumTotalDecompressedIwaBytes));
        ValidatePositive(MaximumSnappyChunkBytes, nameof(MaximumSnappyChunkBytes));
        ValidatePositive(MaximumArchiveInfoBytes, nameof(MaximumArchiveInfoBytes));
        ValidatePositive(MaximumRecordBytes, nameof(MaximumRecordBytes));
        ValidatePositive(MaximumRecordCount, nameof(MaximumRecordCount));
        ValidatePositive(MaximumArchiveReferenceCount, nameof(MaximumArchiveReferenceCount));
        ValidatePositive(MaximumSourceReferenceIssues, nameof(MaximumSourceReferenceIssues));
        ValidatePositive(MaximumSourceDeclarationIssues, nameof(MaximumSourceDeclarationIssues));
        ValidatePositive(MaximumProtobufFieldCount, nameof(MaximumProtobufFieldCount));
        ValidatePositive(MaximumProtobufDepth, nameof(MaximumProtobufDepth));
        ValidatePositive(MaximumTableRows, nameof(MaximumTableRows));
        ValidatePositive(MaximumTableColumns, nameof(MaximumTableColumns));
        ValidatePositive(MaximumMaterializedCells, nameof(MaximumMaterializedCells));
        ValidatePositive(MaximumTableCatalogEntries, nameof(MaximumTableCatalogEntries));
        ValidatePositive(MaximumTableDimensionEntries, nameof(MaximumTableDimensionEntries));
        ValidatePositive(MaximumProjectedSheets, nameof(MaximumProjectedSheets));
        ValidatePositive(MaximumProjectedSlides, nameof(MaximumProjectedSlides));
        ValidatePositive(MaximumProjectedTables, nameof(MaximumProjectedTables));
        ValidatePositive(MaximumProjectedImages, nameof(MaximumProjectedImages));
        ValidatePositive(MaximumImageMetadataEntries, nameof(MaximumImageMetadataEntries));
        ValidatePositive(MaximumProjectedImageBytes, nameof(MaximumProjectedImageBytes));
        ValidatePositive(MaximumDecodedImageBytes, nameof(MaximumDecodedImageBytes));
        ValidatePositive(MaximumTableMergedRanges, nameof(MaximumTableMergedRanges));
        ValidatePositive(MaximumFormulaNodes, nameof(MaximumFormulaNodes));
        ValidatePositive(MaximumFormulaCharacters, nameof(MaximumFormulaCharacters));
        ValidatePositive(MaximumFormulaRenderingOperations, nameof(MaximumFormulaRenderingOperations));
        ValidatePositive(MaximumProjectedTextItems, nameof(MaximumProjectedTextItems));
        ValidatePositive(MaximumProjectedTextCharacters, nameof(MaximumProjectedTextCharacters));
        ValidatePositive(MaximumTextStyleInheritanceDepth, nameof(MaximumTextStyleInheritanceDepth));

        if (MaximumEntryBytes > MaximumTotalEntryBytes) {
            throw new ArgumentException($"{nameof(MaximumEntryBytes)} cannot exceed {nameof(MaximumTotalEntryBytes)}.");
        }
        if (MaximumSnappyChunkBytes > MaximumDecompressedIwaBytes) {
            throw new ArgumentException($"{nameof(MaximumSnappyChunkBytes)} cannot exceed {nameof(MaximumDecompressedIwaBytes)}.");
        }

        return (IWorkReadOptions)MemberwiseClone();
    }

    private static void ValidatePositive(long value, string name) {
        if (value <= 0) throw new ArgumentOutOfRangeException(name, "The limit must be positive.");
    }
}
