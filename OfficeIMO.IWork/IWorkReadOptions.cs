namespace OfficeIMO.IWork;

/// <summary>Bounds and preservation preferences applied while reading an iWork source.</summary>
public sealed class IWorkReadOptions {
    /// <summary>Gets or sets whether raw source records are included in the conversion report for inspection.</summary>
    public bool PreserveSourceRecords { get; set; } = true;

    /// <summary>Gets or sets the maximum source package size in bytes.</summary>
    public long MaxPackageBytes { get; set; } = 512L * 1024 * 1024;

    /// <summary>Gets or sets the maximum number of outer and nested package entries.</summary>
    public int MaxEntryCount { get; set; } = 8192;

    /// <summary>Gets or sets the maximum uncompressed size of one package entry.</summary>
    public long MaxEntryBytes { get; set; } = 128L * 1024 * 1024;

    /// <summary>Gets or sets the maximum combined uncompressed package-entry size.</summary>
    public long MaxTotalEntryBytes { get; set; } = 512L * 1024 * 1024;

    /// <summary>Gets or sets the maximum compressed size of one IWA archive.</summary>
    public int MaxIwaBytes { get; set; } = 64 * 1024 * 1024;

    /// <summary>Gets or sets the maximum decompressed size of one IWA archive.</summary>
    public int MaxDecompressedIwaBytes { get; set; } = 256 * 1024 * 1024;

    /// <summary>Gets or sets the maximum combined decompressed size of all IWA archives in one source.</summary>
    public long MaxTotalDecompressedIwaBytes { get; set; } = 512L * 1024 * 1024;

    /// <summary>Gets or sets the maximum declared decompressed size of one raw Snappy chunk.</summary>
    public int MaxSnappyChunkBytes { get; set; } = 64 * 1024 * 1024;

    /// <summary>Gets or sets the maximum encoded ArchiveInfo header size.</summary>
    public int MaxArchiveInfoBytes { get; set; } = 1024 * 1024;

    /// <summary>Gets or sets the maximum payload size of one archived record.</summary>
    public int MaxRecordBytes { get; set; } = 128 * 1024 * 1024;

    /// <summary>Gets or sets the maximum number of IWA records across the package.</summary>
    public int MaxRecordCount { get; set; } = 1_000_000;

    /// <summary>Gets or sets the maximum combined ArchiveInfo version, object-reference, and data-reference values across the package.</summary>
    public int MaxArchiveReferenceCount { get; set; } = 8_000_000;

    /// <summary>Gets or sets the maximum declared reference occurrences inspected in distinct source fields for unresolved-reference evidence across one semantic projection, including readable siblings in a field with failures.</summary>
    public int MaxSourceReferenceIssues { get; set; } = 100_000;

    /// <summary>Gets or sets the maximum distinct selected source paths retained as unreadable or rejected declaration evidence across one semantic projection.</summary>
    public int MaxSourceDeclarationIssues { get; set; } = 100_000;

    /// <summary>Gets or sets the maximum number of fields decoded from one protobuf message.</summary>
    public int MaxProtobufFieldCount { get; set; } = 1_000_000;

    /// <summary>Gets or sets the maximum nested protobuf message depth.</summary>
    public int MaxProtobufDepth { get; set; } = 64;

    /// <summary>Gets or sets the maximum projected row count of one iWork table.</summary>
    public int MaxTableRows { get; set; } = 1_048_576;

    /// <summary>Gets or sets the maximum projected column count of one iWork table.</summary>
    public int MaxTableColumns { get; set; } = 16_384;

    /// <summary>Gets or sets the maximum number of materialized non-empty cells across one iWork source.</summary>
    public int MaxMaterializedCells { get; set; } = 10_000_000;

    /// <summary>Gets or sets the maximum combined string and formula catalog entries decoded across one iWork source.</summary>
    public int MaxTableCatalogEntries { get; set; } = 100_000;

    /// <summary>Gets or sets the maximum combined table dimension headers, bucket references, bounded modern row offset slots, hidden-state declarations, selected UUID position maps, grouping metadata and filter references inspected across one iWork source.</summary>
    public int MaxTableDimensionEntries { get; set; } = 1_100_000;

    /// <summary>Gets or sets the maximum combined table stroke layer and run declarations inspected across one semantic projection.</summary>
    public int MaxTableBorderEntries { get; set; } = 100_000;

    /// <summary>Gets or sets the maximum number of projected sheets across one Numbers source.</summary>
    public int MaxProjectedSheets { get; set; } = 4096;

    /// <summary>Gets or sets the maximum number of projected slides across one Keynote source.</summary>
    public int MaxProjectedSlides { get; set; } = 10_000;

    /// <summary>Gets or sets the maximum number of projected tables across one iWork source.</summary>
    public int MaxProjectedTables { get; set; } = 4096;

    /// <summary>Gets or sets the maximum number of projected images across one iWork source.</summary>
    public int MaxProjectedImages { get; set; } = 4096;

    /// <summary>Gets or sets the maximum number of package image-metadata entries decoded across one iWork source.</summary>
    public int MaxImageMetadataEntries { get; set; } = 16_384;

    /// <summary>Gets or sets the maximum combined encoded image bytes emitted by one semantic projection, counting repeated uses.</summary>
    public long MaxProjectedImageBytes { get; set; } = 512L * 1024 * 1024;

    /// <summary>Gets or sets the maximum cumulative decoded image work for preview inspection and for each semantic projection. Default: 64 MiB.</summary>
    public long MaxDecodedImageBytes { get; set; } = 64L * 1024 * 1024;

    /// <summary>Gets or sets the maximum number of merged ranges projected from one iWork table.</summary>
    public int MaxTableMergedRanges { get; set; } = 100_000;

    /// <summary>Gets or sets the maximum number of syntax nodes decoded from one iWork formula.</summary>
    /// <remarks>Also bounds the cumulative referenced cell positions inspected while qualifying GCD/LCM operands in that formula. Repeated references count again; larger references retain their caches with incomplete expressions.</remarks>
    public int MaxFormulaNodes { get; set; } = 4096;

    /// <summary>Gets or sets the maximum reconstructed formula length in characters.</summary>
    public int MaxFormulaCharacters { get; set; } = 8192;

    /// <summary>Gets or sets the maximum combined formula-rendering work across one semantic projection, counting repeated formula uses.</summary>
    /// <remarks>Includes numeric cell inspections for bounded GCD/LCM reference qualification.</remarks>
    public long MaxFormulaRenderingOperations { get; set; } = 64L * 1024 * 1024;

    /// <summary>Gets or sets the maximum number of projected text items across one semantic projection.</summary>
    public int MaxProjectedTextItems { get; set; } = 100_000;

    /// <summary>Gets or sets the maximum number of decoded text characters across one semantic projection.</summary>
    public long MaxProjectedTextCharacters { get; set; } = 16L * 1024 * 1024;

    /// <summary>Gets or sets the maximum cross-record inheritance depth of an iWork text or table style.</summary>
    public int MaxTextStyleInheritanceDepth { get; set; } = 64;

    /// <summary>Returns a validated copy suitable for an isolated reader registration.</summary>
    public IWorkReadOptions Clone() => Snapshot();

    internal IWorkReadOptions Snapshot() {
        ValidatePositive(MaxPackageBytes, nameof(MaxPackageBytes));
        ValidatePositive(MaxEntryCount, nameof(MaxEntryCount));
        ValidatePositive(MaxEntryBytes, nameof(MaxEntryBytes));
        ValidatePositive(MaxTotalEntryBytes, nameof(MaxTotalEntryBytes));
        ValidatePositive(MaxIwaBytes, nameof(MaxIwaBytes));
        ValidatePositive(MaxDecompressedIwaBytes, nameof(MaxDecompressedIwaBytes));
        ValidatePositive(MaxTotalDecompressedIwaBytes, nameof(MaxTotalDecompressedIwaBytes));
        ValidatePositive(MaxSnappyChunkBytes, nameof(MaxSnappyChunkBytes));
        ValidatePositive(MaxArchiveInfoBytes, nameof(MaxArchiveInfoBytes));
        ValidatePositive(MaxRecordBytes, nameof(MaxRecordBytes));
        ValidatePositive(MaxRecordCount, nameof(MaxRecordCount));
        ValidatePositive(MaxArchiveReferenceCount, nameof(MaxArchiveReferenceCount));
        ValidatePositive(MaxSourceReferenceIssues, nameof(MaxSourceReferenceIssues));
        ValidatePositive(MaxSourceDeclarationIssues, nameof(MaxSourceDeclarationIssues));
        ValidatePositive(MaxProtobufFieldCount, nameof(MaxProtobufFieldCount));
        ValidatePositive(MaxProtobufDepth, nameof(MaxProtobufDepth));
        ValidatePositive(MaxTableRows, nameof(MaxTableRows));
        ValidatePositive(MaxTableColumns, nameof(MaxTableColumns));
        ValidatePositive(MaxMaterializedCells, nameof(MaxMaterializedCells));
        ValidatePositive(MaxTableCatalogEntries, nameof(MaxTableCatalogEntries));
        ValidatePositive(MaxTableDimensionEntries, nameof(MaxTableDimensionEntries));
        ValidatePositive(MaxTableBorderEntries, nameof(MaxTableBorderEntries));
        ValidatePositive(MaxProjectedSheets, nameof(MaxProjectedSheets));
        ValidatePositive(MaxProjectedSlides, nameof(MaxProjectedSlides));
        ValidatePositive(MaxProjectedTables, nameof(MaxProjectedTables));
        ValidatePositive(MaxProjectedImages, nameof(MaxProjectedImages));
        ValidatePositive(MaxImageMetadataEntries, nameof(MaxImageMetadataEntries));
        ValidatePositive(MaxProjectedImageBytes, nameof(MaxProjectedImageBytes));
        ValidatePositive(MaxDecodedImageBytes, nameof(MaxDecodedImageBytes));
        ValidatePositive(MaxTableMergedRanges, nameof(MaxTableMergedRanges));
        ValidatePositive(MaxFormulaNodes, nameof(MaxFormulaNodes));
        ValidatePositive(MaxFormulaCharacters, nameof(MaxFormulaCharacters));
        ValidatePositive(MaxFormulaRenderingOperations, nameof(MaxFormulaRenderingOperations));
        ValidatePositive(MaxProjectedTextItems, nameof(MaxProjectedTextItems));
        ValidatePositive(MaxProjectedTextCharacters, nameof(MaxProjectedTextCharacters));
        ValidatePositive(MaxTextStyleInheritanceDepth, nameof(MaxTextStyleInheritanceDepth));

        if (MaxEntryBytes > MaxTotalEntryBytes) {
            throw new ArgumentException($"{nameof(MaxEntryBytes)} cannot exceed {nameof(MaxTotalEntryBytes)}.");
        }
        if (MaxSnappyChunkBytes > MaxDecompressedIwaBytes) {
            throw new ArgumentException($"{nameof(MaxSnappyChunkBytes)} cannot exceed {nameof(MaxDecompressedIwaBytes)}.");
        }

        return (IWorkReadOptions)MemberwiseClone();
    }

    private static void ValidatePositive(long value, string name) {
        if (value <= 0) throw new ArgumentOutOfRangeException(name, "The limit must be positive.");
    }
}
