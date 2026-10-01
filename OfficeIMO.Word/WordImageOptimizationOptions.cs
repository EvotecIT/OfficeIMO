using OfficeIMO.Drawing;

namespace OfficeIMO.Word;

/// <summary>Managed embedded-media optimization settings. Changes remain in memory until a normal save.</summary>
public sealed class WordImageOptimizationOptions {
    /// <summary>Defaults to placement-aware downsampling without same-size JPEG recompression.</summary>
    public OfficeImageOptimizationMode Mode { get; set; } = OfficeImageOptimizationMode.Downsample;
    /// <summary>Required pixels per displayed inch, between 36 and 1200. Defaults to 144.</summary>
    public double TargetDpi { get; set; } = 144;
    /// <summary>JPEG encoding quality, from 1 through 100. Defaults to 85.</summary>
    public int JpegQuality { get; set; } = 85;
    /// <summary>Preserves original bytes when the candidate is not smaller.</summary>
    public bool KeepOriginalWhenNotSmaller { get; set; } = true;
    /// <summary>Filtering used for pixel reduction.</summary>
    public OfficeRasterResamplingMode ResamplingMode { get; set; } = OfficeRasterResamplingMode.Bilinear;
    /// <summary>Metadata policy. Metadata loss blocks replacement unless explicitly allowed.</summary>
    public OfficeImageMetadataPolicy MetadataPolicy { get; set; } = OfficeImageMetadataPolicy.Preserve;
    /// <summary>Categories retained by the selective metadata policy.</summary>
    public OfficeImageMetadataKinds MetadataSelection { get; set; } = OfficeImageMetadataKinds.All;
    /// <summary>Allows reported metadata loss, including intentional stripping. Defaults to false.</summary>
    public bool AllowMetadataLoss { get; set; }
    /// <summary>Maximum encoded bytes read from one image part. Defaults to 32 MiB.</summary>
    public long MaxImageBytes { get; set; } = 32L * 1024 * 1024;
    /// <summary>Maximum combined original and candidate bytes retained for transactional replacement. Defaults to 128 MiB.</summary>
    public long MaxStagedBytes { get; set; } = 128L * 1024 * 1024;
    /// <summary>Maximum package parts visited. Defaults to 10000.</summary>
    public int MaxPackageParts { get; set; } = 10_000;
    /// <summary>Save policy used to authorize invalidating a signed package during mutation.</summary>
    public WordSignedDocumentSavePolicy SignedDocumentPolicy { get; set; } = WordSignedDocumentSavePolicy.Block;

    /// <summary>Creates an independently validated copy of these settings.</summary>
    public WordImageOptimizationOptions Clone() {
        if (double.IsNaN(TargetDpi) || double.IsInfinity(TargetDpi) || TargetDpi < 36 || TargetDpi > 1200)
            throw new ArgumentOutOfRangeException(nameof(TargetDpi));
        if (JpegQuality < 1 || JpegQuality > 100) throw new ArgumentOutOfRangeException(nameof(JpegQuality));
        if (Mode < OfficeImageOptimizationMode.Downsample || Mode > OfficeImageOptimizationMode.DownsampleAndRecompress)
            throw new ArgumentOutOfRangeException(nameof(Mode));
        if (MaxImageBytes <= 0 || MaxImageBytes > int.MaxValue) throw new ArgumentOutOfRangeException(nameof(MaxImageBytes));
        if (MaxStagedBytes <= 0) throw new ArgumentOutOfRangeException(nameof(MaxStagedBytes));
        if (MaxPackageParts <= 0) throw new ArgumentOutOfRangeException(nameof(MaxPackageParts));
        if (ResamplingMode < OfficeRasterResamplingMode.NearestNeighbor || ResamplingMode > OfficeRasterResamplingMode.Lanczos3)
            throw new ArgumentOutOfRangeException(nameof(ResamplingMode));
        if (MetadataPolicy < OfficeImageMetadataPolicy.Preserve || MetadataPolicy > OfficeImageMetadataPolicy.SelectiveCopy)
            throw new ArgumentOutOfRangeException(nameof(MetadataPolicy));
        if ((MetadataSelection & ~OfficeImageMetadataKinds.All) != 0)
            throw new ArgumentOutOfRangeException(nameof(MetadataSelection));
        if (SignedDocumentPolicy != WordSignedDocumentSavePolicy.Block && SignedDocumentPolicy != WordSignedDocumentSavePolicy.AllowSignatureInvalidation)
            throw new ArgumentOutOfRangeException(nameof(SignedDocumentPolicy));
        return (WordImageOptimizationOptions)MemberwiseClone();
    }
}
