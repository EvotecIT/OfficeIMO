using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

/// <summary>Controls optional placement-aware image downsampling during PDF generation.</summary>
public sealed class PdfImageOptimizationOptions {
    private double _targetDpi = 144D;
    private double _downsampleThreshold = 1.1D;
    private int _jpegQuality = 85;
    private OfficeImageOptimizationMode _mode;

    /// <summary>Enables managed placement-aware image optimization. Disabled by default.</summary>
    public bool Enabled { get; set; }

    /// <summary>Selects placement downsampling, same-size JPEG recompression, or both. Defaults to downsampling.</summary>
    public OfficeImageOptimizationMode Mode {
        get => _mode;
        set {
            if (value < OfficeImageOptimizationMode.Downsample || value > OfficeImageOptimizationMode.DownsampleAndRecompress)
                throw new System.ArgumentOutOfRangeException(nameof(Mode));
            _mode = value;
        }
    }

    /// <summary>Target image resolution in pixels per inch. Defaults to 144.</summary>
    public double TargetDpi {
        get => _targetDpi;
        set {
            if (double.IsNaN(value) || double.IsInfinity(value) || value < 36D || value > 1200D) {
                throw new System.ArgumentOutOfRangeException(nameof(TargetDpi), "PDF image target DPI must be between 36 and 1200.");
            }
            _targetDpi = value;
        }
    }

    /// <summary>
    /// Minimum ratio between source pixels and placement-required pixels before downsampling.
    /// Defaults to 1.1 to avoid re-encoding for negligible reductions.
    /// </summary>
    public double DownsampleThreshold {
        get => _downsampleThreshold;
        set {
            if (double.IsNaN(value) || double.IsInfinity(value) || value < 1D || value > 10D) {
                throw new System.ArgumentOutOfRangeException(nameof(DownsampleThreshold), "PDF image downsample threshold must be between 1 and 10.");
            }
            _downsampleThreshold = value;
        }
    }

    /// <summary>Sampling mode used when source pixels are reduced.</summary>
    public OfficeRasterResamplingMode ResamplingMode { get; set; } = OfficeRasterResamplingMode.Bilinear;

    /// <summary>Quality used when a JPEG source is downsampled. Defaults to 85.</summary>
    public int JpegQuality {
        get => _jpegQuality;
        set {
            if (value < 1 || value > 100) {
                throw new System.ArgumentOutOfRangeException(nameof(JpegQuality), "PDF image JPEG quality must be between 1 and 100.");
            }
            _jpegQuality = value;
        }
    }

    /// <summary>Keeps the original encoded payload when optimized bytes are not smaller.</summary>
    public bool KeepOriginalWhenNotSmaller { get; set; } = true;

    /// <summary>Metadata behavior for rewritten images. Defaults to preservation.</summary>
    public OfficeImageMetadataPolicy MetadataPolicy { get; set; } = OfficeImageMetadataPolicy.Preserve;
    /// <summary>Metadata categories retained by selective copying.</summary>
    public OfficeImageMetadataKinds MetadataSelection { get; set; } = OfficeImageMetadataKinds.All;
    /// <summary>Permits reported metadata loss. False preserves the original image if metadata would be lost.</summary>
    public bool AllowMetadataLoss { get; set; }

    /// <summary>Creates an independent copy of these options.</summary>
    public PdfImageOptimizationOptions Clone() {
        if (ResamplingMode < OfficeRasterResamplingMode.NearestNeighbor || ResamplingMode > OfficeRasterResamplingMode.Lanczos3)
            throw new System.ArgumentOutOfRangeException(nameof(ResamplingMode));
        if (MetadataPolicy < OfficeImageMetadataPolicy.Preserve || MetadataPolicy > OfficeImageMetadataPolicy.SelectiveCopy)
            throw new System.ArgumentOutOfRangeException(nameof(MetadataPolicy));
        if ((MetadataSelection & ~OfficeImageMetadataKinds.All) != 0)
            throw new System.ArgumentOutOfRangeException(nameof(MetadataSelection));
        return new PdfImageOptimizationOptions {
        Enabled = Enabled,
        Mode = Mode,
        TargetDpi = TargetDpi,
        DownsampleThreshold = DownsampleThreshold,
        ResamplingMode = ResamplingMode,
        JpegQuality = JpegQuality,
        KeepOriginalWhenNotSmaller = KeepOriginalWhenNotSmaller,
        MetadataPolicy = MetadataPolicy,
        MetadataSelection = MetadataSelection,
        AllowMetadataLoss = AllowMetadataLoss
        };
    }
}
