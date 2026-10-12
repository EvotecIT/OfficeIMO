namespace OfficeIMO.Publisher;

/// <summary>Resource bounds for managed Publisher decoding. Limits reject input rather than silently truncate it.</summary>
public sealed class PublisherReadOptions {
    /// <summary>Source byte, text, record, object, and compound-stream limits. MaxInputBytes also bounds cumulative encoded image processing; MaxItems bounds image-store entries including unavailable assets, cumulative projected gradient stops, and recovered table tracks/cells. MaxTextCharacters also bounds copied table text.</summary>
    public OfficeLegacyImportLimits Limits { get; set; } = new OfficeLegacyImportLimits();
    /// <summary>Maximum number of document and master pages combined.</summary>
    public int MaximumPages { get; set; } = 1024;
    /// <summary>Maximum nesting of publication containers and drawing groups.</summary>
    public int MaximumNestingDepth { get; set; } = 64;
    /// <summary>Maximum extracted bytes for one image, including decompressed metafiles.</summary>
    public int MaximumImageBytes { get; set; } = 16 * 1024 * 1024;
    /// <summary>Maximum total extracted and projected image bytes.</summary>
    public int MaximumTotalImageBytes { get; set; } = 64 * 1024 * 1024;
    /// <summary>Optional trusted application codec for projecting native WMF/EMF pictures and supported raster payloads requiring picture effects. Original bytes remain in Images.</summary>
    public OfficeIMO.Drawing.IOfficeRasterImageCodec? ImageCodec { get; set; }
    /// <summary>Maximum decoded pixels for one projected picture, including application codec output and picture-effect decoding.</summary>
    public long MaximumRasterPixels { get; set; } = 8_000_000;
    /// <summary>Maximum cumulative picture-effect work in pixels. Charges the selected raster once for decoding and once for filtering, plus each inspected GIF, WebP or icon frame.</summary>
    public long MaximumImageProcessingPixels { get; set; } = 64_000_000;

    /// <summary>Creates a validated independent options copy. A supplied application image codec remains shared.</summary>
    public PublisherReadOptions Clone() {
        if (Limits == null) throw new ArgumentNullException(nameof(Limits));
        Limits.Validate();
        if (MaximumPages < 1) throw new ArgumentOutOfRangeException(nameof(MaximumPages));
        if (MaximumNestingDepth < 1 || MaximumNestingDepth > 128) throw new ArgumentOutOfRangeException(nameof(MaximumNestingDepth));
        if (MaximumImageBytes < 1) throw new ArgumentOutOfRangeException(nameof(MaximumImageBytes));
        if (MaximumTotalImageBytes < 1) throw new ArgumentOutOfRangeException(nameof(MaximumTotalImageBytes));
        if (MaximumRasterPixels < 1) throw new ArgumentOutOfRangeException(nameof(MaximumRasterPixels));
        if (MaximumImageProcessingPixels < 1) throw new ArgumentOutOfRangeException(nameof(MaximumImageProcessingPixels));
        return new PublisherReadOptions {
            Limits = Limits.Clone(), MaximumPages = MaximumPages, MaximumNestingDepth = MaximumNestingDepth,
            MaximumImageBytes = MaximumImageBytes, MaximumTotalImageBytes = MaximumTotalImageBytes,
            ImageCodec = ImageCodec, MaximumRasterPixels = MaximumRasterPixels,
            MaximumImageProcessingPixels = MaximumImageProcessingPixels
        };
    }
}
