using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher;

/// <summary>Page size and resource bounds for an in-memory publication. Native PUB saving requires a supported loaded source.</summary>
public sealed class PublisherCreateOptions {
    /// <summary>Width of the initial page. Defaults to A4 portrait, 210 millimeters.</summary>
    public OfficeLength Width { get; set; } = OfficeLength.Millimeters(210);
    /// <summary>Height of the initial page. Defaults to A4 portrait, 297 millimeters.</summary>
    public OfficeLength Height { get; set; } = OfficeLength.Millimeters(297);
    /// <summary>Text, record, object, compound-stream and image resource limits used by the model and projection.</summary>
    public OfficeLegacyImportLimits Limits { get; set; } = new OfficeLegacyImportLimits();
    /// <summary>Maximum number of document and master pages combined.</summary>
    public int MaxPages { get; set; } = 1024;
    /// <summary>Maximum nesting of publication containers and drawing groups.</summary>
    public int MaxNestingDepth { get; set; } = 64;
    /// <summary>Maximum extracted bytes for one image, including decompressed metafiles.</summary>
    public int MaxImageBytes { get; set; } = 16 * 1024 * 1024;
    /// <summary>Maximum total extracted and projected image bytes.</summary>
    public int MaxTotalImageBytes { get; set; } = 64 * 1024 * 1024;
    /// <summary>Optional trusted application codec for supported picture projection and effects. The codec remains caller-owned.</summary>
    public IOfficeRasterImageCodec? ImageCodec { get; set; }
    /// <summary>Optional scoped font faces for story flow and rendering. The collection is copied; providers remain shared.</summary>
    public OfficeFontFaceCollection? Fonts { get; set; }
    /// <summary>Maximum decoded pixels for one projected picture, including application codec output and picture-effect decoding.</summary>
    public long MaxRasterPixels { get; set; } = 8_000_000;
    /// <summary>Maximum cumulative picture-effect work in pixels, including decoding, filtering and inspected frames.</summary>
    public long MaxImageProcessingPixels { get; set; } = 64_000_000;

    /// <summary>Snapshots settings into the existing model and projection budget owner without exposing load settings as a public bag.</summary>
    internal PublisherLoadOptions ToLoadOptions() => new PublisherLoadOptions {
        Limits = Limits, MaxPages = MaxPages, MaxNestingDepth = MaxNestingDepth,
        MaxImageBytes = MaxImageBytes, MaxTotalImageBytes = MaxTotalImageBytes,
        ImageCodec = ImageCodec, Fonts = Fonts, MaxRasterPixels = MaxRasterPixels,
        MaxImageProcessingPixels = MaxImageProcessingPixels
    }.Clone();
}
