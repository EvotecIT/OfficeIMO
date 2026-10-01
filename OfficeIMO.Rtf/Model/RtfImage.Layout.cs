namespace OfficeIMO.Rtf;

public sealed partial class RtfImage {
    /// <summary>Resolves the picture's unscaled and visible sizes in twips. Pixel sizes use 96 DPI when no goal size is supplied.</summary>
    public RtfImageLayout ResolveLayout(double? fallbackWidthTwips = null, double? fallbackHeightTwips = null) {
        double? width = DesiredWidthTwips ?? (SourceWidth.HasValue ? SourceWidth.Value * 15d : fallbackWidthTwips);
        double? height = DesiredHeightTwips ?? (SourceHeight.HasValue ? SourceHeight.Value * 15d : fallbackHeightTwips);
        double scaleX = (ScaleXPercent ?? 100) / 100d;
        double scaleY = (ScaleYPercent ?? 100) / 100d;
        double? visibleWidth = (width - (double)(CropLeftTwips ?? 0) - (CropRightTwips ?? 0)) * scaleX;
        double? visibleHeight = (height - (double)(CropTopTwips ?? 0) - (CropBottomTwips ?? 0)) * scaleY;
        if (scaleX <= 0 || scaleY <= 0 || width <= 0 || height <= 0 || visibleWidth <= 0 || visibleHeight <= 0) {
            throw new InvalidDataException("Picture scaling and cropping must leave a positive visible size.");
        }
        return new RtfImageLayout(width, height, visibleWidth, visibleHeight, scaleX, scaleY);
    }
}

/// <summary>Resolved picture dimensions in twips, before and after cropping and scaling.</summary>
public sealed class RtfImageLayout {
    internal RtfImageLayout(double? width, double? height, double? visibleWidth, double? visibleHeight, double scaleX, double scaleY) {
        WidthTwips = width; HeightTwips = height; VisibleWidthTwips = visibleWidth; VisibleHeightTwips = visibleHeight; ScaleX = scaleX; ScaleY = scaleY;
    }
    /// <summary>Width before cropping and scaling.</summary>
    public double? WidthTwips { get; }
    /// <summary>Height before cropping and scaling.</summary>
    public double? HeightTwips { get; }
    /// <summary>Visible width after cropping and scaling.</summary>
    public double? VisibleWidthTwips { get; }
    /// <summary>Visible height after cropping and scaling.</summary>
    public double? VisibleHeightTwips { get; }
    /// <summary>Horizontal scale multiplier.</summary>
    public double ScaleX { get; }
    /// <summary>Vertical scale multiplier.</summary>
    public double ScaleY { get; }
}
