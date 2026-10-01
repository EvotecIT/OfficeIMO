namespace OfficeIMO.Rtf;

/// <summary>
/// RTF picture payload preserved without external imaging dependencies.
/// </summary>
public sealed partial class RtfImage : IRtfBlock, IRtfInline {
    /// <summary>Creates an image block.</summary>
    public RtfImage(RtfImageFormat format, byte[] data) {
        Format = format;
        Data = data ?? Array.Empty<byte>();
    }

    /// <summary>Image format.</summary>
    public RtfImageFormat Format { get; set; }

    /// <summary>Raw image bytes.</summary>
    public byte[] Data { get; set; }

    /// <summary>Original width in pixels when present.</summary>
    public int? SourceWidth { get; set; }

    /// <summary>Original height in pixels when present.</summary>
    public int? SourceHeight { get; set; }

    /// <summary>Desired width in twips when present.</summary>
    public int? DesiredWidthTwips { get; set; }

    /// <summary>Desired height in twips when present.</summary>
    public int? DesiredHeightTwips { get; set; }

    /// <summary>Horizontal picture scale as a percentage. An omitted value means 100.</summary>
    public int? ScaleXPercent { get; set; }

    /// <summary>Vertical picture scale as a percentage. An omitted value means 100.</summary>
    public int? ScaleYPercent { get; set; }

    /// <summary>Unscaled twips cropped from the left edge. Negative values add space.</summary>
    public int? CropLeftTwips { get; set; }

    /// <summary>Unscaled twips cropped from the top edge. Negative values add space.</summary>
    public int? CropTopTwips { get; set; }

    /// <summary>Unscaled twips cropped from the right edge. Negative values add space.</summary>
    public int? CropRightTwips { get; set; }

    /// <summary>Unscaled twips cropped from the bottom edge. Negative values add space.</summary>
    public int? CropBottomTwips { get; set; }

    /// <summary>Alternative text or description.</summary>
    public string? Description { get; set; }
}
