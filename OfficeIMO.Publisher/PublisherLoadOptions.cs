namespace OfficeIMO.Publisher;

/// <summary>Resource bounds for loading, editing and projecting a Publisher publication. In-memory creation uses the same bounds for its model and projection.</summary>
public sealed class PublisherLoadOptions {
    /// <summary>Returns fresh bounded settings for untrusted publications using the existing supported import and projection limits.</summary>
    public static PublisherLoadOptions UntrustedDefaults => new PublisherLoadOptions();
    /// <summary>Source byte, text, record, object, and compound-stream limits. MaxInputBytes also bounds cumulative encoded image processing. MaxItems bounds image-store entries, parsed and projected gradient stops, table tracks and cells, and native table utility paint records. MaxTextCharacters also bounds copied table text and edited or projected labels.</summary>
    public OfficeLegacyImportLimits Limits { get; set; } = new OfficeLegacyImportLimits();
    /// <summary>Maximum number of document and master pages combined.</summary>
    public int MaxPages { get; set; } = 1024;
    /// <summary>Maximum nesting of publication containers and drawing groups.</summary>
    public int MaxNestingDepth { get; set; } = 64;
    /// <summary>Maximum extracted bytes for one image, including decompressed metafiles.</summary>
    public int MaxImageBytes { get; set; } = 16 * 1024 * 1024;
    /// <summary>Maximum total extracted and projected image bytes.</summary>
    public int MaxTotalImageBytes { get; set; } = 64 * 1024 * 1024;
    /// <summary>Optional trusted application codec for projecting native WMF/EMF pictures and supported raster payloads requiring picture effects. Original bytes remain in Images.</summary>
    public OfficeIMO.Drawing.IOfficeRasterImageCodec? ImageCodec { get; set; }
    /// <summary>Optional scoped font faces used for story flow and drawing rendering. The collection is copied; providers remain shared.</summary>
    public OfficeIMO.Drawing.OfficeFontFaceCollection? Fonts { get; set; }
    /// <summary>Maximum decoded pixels for one projected picture, including application codec output and picture-effect decoding.</summary>
    public long MaxRasterPixels { get; set; } = 8_000_000;
    /// <summary>Maximum cumulative picture-effect work in pixels. Charges the selected raster once for decoding and once for filtering, plus each inspected GIF, WebP or icon frame.</summary>
    public long MaxImageProcessingPixels { get; set; } = 64_000_000;

    /// <summary>Creates a validated independent options copy. Application image codecs and font providers remain shared.</summary>
    public PublisherLoadOptions Clone() {
        ArgumentNullException.ThrowIfNull(Limits);
        Limits.Validate();
        ArgumentOutOfRangeException.ThrowIfLessThan(MaxPages, 1);
        ArgumentOutOfRangeException.ThrowIfLessThan(MaxNestingDepth, 1);
        ArgumentOutOfRangeException.ThrowIfGreaterThan(MaxNestingDepth, 128);
        ArgumentOutOfRangeException.ThrowIfLessThan(MaxImageBytes, 1);
        ArgumentOutOfRangeException.ThrowIfLessThan(MaxTotalImageBytes, 1);
        ArgumentOutOfRangeException.ThrowIfLessThan(MaxRasterPixels, 1);
        ArgumentOutOfRangeException.ThrowIfLessThan(MaxImageProcessingPixels, 1);
        return new PublisherLoadOptions {
            Limits = Limits.Clone(), MaxPages = MaxPages, MaxNestingDepth = MaxNestingDepth,
            MaxImageBytes = MaxImageBytes, MaxTotalImageBytes = MaxTotalImageBytes,
            ImageCodec = ImageCodec, Fonts = Fonts?.Clone(), MaxRasterPixels = MaxRasterPixels,
            MaxImageProcessingPixels = MaxImageProcessingPixels
        };
    }
}
