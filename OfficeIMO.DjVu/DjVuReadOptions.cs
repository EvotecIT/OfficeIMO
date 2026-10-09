namespace OfficeIMO.DjVu;

/// <summary>Limits applied to the owned source, document structure and expanded codec data.</summary>
public sealed class DjVuReadOptions {
    /// <summary>Maximum aggregate source bytes, including explicitly resolved components.</summary>
    public long MaxSourceBytes { get; set; } = 256L * 1024 * 1024;
    /// <summary>Maximum number of component files in the document.</summary>
    public int MaxComponents { get; set; } = 50_000;
    /// <summary>Maximum number of pages.</summary>
    public int MaxPages { get; set; } = 10_000;
    /// <summary>Maximum total IFF chunks, including nested containers.</summary>
    public int MaxChunks { get; set; } = 200_000;
    /// <summary>Maximum nesting of IFF containers and included components.</summary>
    public int MaxDepth { get; set; } = 32;
    /// <summary>Maximum aggregate bytes expanded while reading directories and text.</summary>
    public long MaxExpandedBytes { get; set; } = 64L * 1024 * 1024;
    /// <summary>Maximum size of one BZZ block, including its end marker.</summary>
    public int MaxBzzBlockBytes { get; set; } = 4 * 1024 * 1024;
    /// <summary>Maximum aggregate retained text characters.</summary>
    public int MaxTextCharacters { get; set; } = 32 * 1024 * 1024;
    /// <summary>Maximum aggregate text-zone records.</summary>
    public int MaxTextZones { get; set; } = 1_000_000;
    /// <summary>Maximum nesting of text zones.</summary>
    public int MaxTextZoneDepth { get; set; } = 64;
    /// <summary>Maximum bookmark records in the compressed document outline.</summary>
    public int MaxBookmarks { get; set; } = 50_000;
    /// <summary>Maximum nesting of document bookmarks.</summary>
    public int MaxBookmarkDepth { get; set; } = 64;
    /// <summary>Maximum aggregate UTF-16 title and target characters in the document outline.</summary>
    public int MaxBookmarkCharacters { get; set; } = 2 * 1024 * 1024;
    /// <summary>Maximum native pixels in one page.</summary>
    public long MaxPagePixels { get; set; } = 32L * 1024 * 1024;
    /// <summary>Maximum bytes of concurrently retained codec working buffers in an operation.</summary>
    public long MaxCodecBytes { get; set; } = 256L * 1024 * 1024;
    /// <summary>Maximum decoded JB2 dictionary and page symbols in an operation.</summary>
    public int MaxSymbols { get; set; } = 1_000_000;
    /// <summary>Maximum symbol placements in one page.</summary>
    public int MaxSymbolPlacements { get; set; } = 2_000_000;
    /// <summary>Maximum aggregate bitmap samples decoded for JB2 page and inherited dictionary symbols in one render.</summary>
    public long MaxJb2DecodedSamples { get; set; } = 128L * 1024 * 1024;
    /// <summary>Maximum aggregate comment bytes decoded in JB2 page and inherited dictionary streams in one render.</summary>
    public long MaxJb2CommentBytes { get; set; } = 1024L * 1024;
    /// <summary>Maximum cumulative progressive slices in one IW44 image, across all chunks.</summary>
    public int MaxIw44Slices { get; set; } = 256;
    /// <summary>Maximum padded coefficient positions visited by IW44 slices across all colour planes and layers in a render operation.</summary>
    public long MaxIw44CoefficientSamples { get; set; } = 512L * 1024 * 1024;
    /// <summary>
    /// Optional resolver for an indirect document. IDs are passed as opaque strings; OfficeIMO never
    /// opens paths or URLs named by the document. Returned bytes are copied into the owned snapshot.
    /// </summary>
    public Func<string, CancellationToken, byte[]>? ComponentResolver { get; set; }

    /// <summary>Creates an independent, validated settings snapshot. The trusted resolver is retained by reference.</summary>
    public DjVuReadOptions Clone() => Snapshot();

    internal DjVuReadOptions Snapshot() {
        var copy = (DjVuReadOptions)MemberwiseClone();
        copy.Validate();
        return copy;
    }

    internal void Validate() {
        if (MaxSourceBytes <= 0 || MaxSourceBytes > int.MaxValue) throw new ArgumentOutOfRangeException(nameof(MaxSourceBytes));
        if (MaxComponents <= 0) throw new ArgumentOutOfRangeException(nameof(MaxComponents));
        if (MaxPages <= 0) throw new ArgumentOutOfRangeException(nameof(MaxPages));
        if (MaxChunks <= 0) throw new ArgumentOutOfRangeException(nameof(MaxChunks));
        if (MaxDepth <= 0 || MaxDepth > 256) throw new ArgumentOutOfRangeException(nameof(MaxDepth));
        if (MaxExpandedBytes <= 0 || MaxExpandedBytes > int.MaxValue) throw new ArgumentOutOfRangeException(nameof(MaxExpandedBytes));
        if (MaxBzzBlockBytes <= 0 || MaxBzzBlockBytes > 4 * 1024 * 1024) throw new ArgumentOutOfRangeException(nameof(MaxBzzBlockBytes));
        if (MaxTextCharacters <= 0) throw new ArgumentOutOfRangeException(nameof(MaxTextCharacters));
        if (MaxTextZones <= 0) throw new ArgumentOutOfRangeException(nameof(MaxTextZones));
        if (MaxTextZoneDepth <= 0 || MaxTextZoneDepth > 256) throw new ArgumentOutOfRangeException(nameof(MaxTextZoneDepth));
        if (MaxBookmarks <= 0) throw new ArgumentOutOfRangeException(nameof(MaxBookmarks));
        if (MaxBookmarkDepth <= 0 || MaxBookmarkDepth > 256) throw new ArgumentOutOfRangeException(nameof(MaxBookmarkDepth));
        if (MaxBookmarkCharacters <= 0) throw new ArgumentOutOfRangeException(nameof(MaxBookmarkCharacters));
        if (MaxPagePixels <= 0) throw new ArgumentOutOfRangeException(nameof(MaxPagePixels));
        if (MaxCodecBytes <= 0) throw new ArgumentOutOfRangeException(nameof(MaxCodecBytes));
        if (MaxSymbols <= 0) throw new ArgumentOutOfRangeException(nameof(MaxSymbols));
        if (MaxSymbolPlacements <= 0) throw new ArgumentOutOfRangeException(nameof(MaxSymbolPlacements));
        if (MaxJb2DecodedSamples <= 0) throw new ArgumentOutOfRangeException(nameof(MaxJb2DecodedSamples));
        if (MaxJb2CommentBytes <= 0) throw new ArgumentOutOfRangeException(nameof(MaxJb2CommentBytes));
        if (MaxIw44Slices <= 0) throw new ArgumentOutOfRangeException(nameof(MaxIw44Slices));
        if (MaxIw44CoefficientSamples <= 0) throw new ArgumentOutOfRangeException(nameof(MaxIw44CoefficientSamples));
    }
}

/// <summary>A document or decoded representation exceeds an explicit operation limit.</summary>
public sealed class DjVuResourceLimitException : IOException {
    /// <summary>Creates a resource-limit failure.</summary>
    public DjVuResourceLimitException(string limitName) : base("DjVu input exceeds " + limitName + ".") => LimitName = limitName;
    /// <summary>Name of the limit that was exceeded.</summary>
    public string LimitName { get; }
}
