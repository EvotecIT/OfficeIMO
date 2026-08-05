using System;

namespace OfficeIMO.Drawing;

/// <summary>
/// Shared fit and source-crop placement plan for rendering an image into a target rectangle.
/// </summary>
public readonly struct OfficeImageRenderPlan {
    private OfficeImageRenderPlan(
        OfficeImagePlacement targetPlacement,
        OfficeImagePlacement visiblePlacement,
        OfficeImagePlacement imagePlacement,
        OfficeImageSourceCrop sourceCrop,
        OfficeImageFit fit,
        bool requiresTargetClip) {
        TargetPlacement = targetPlacement;
        VisiblePlacement = visiblePlacement;
        ImagePlacement = imagePlacement;
        SourceCrop = sourceCrop;
        Fit = fit;
        RequiresTargetClip = requiresTargetClip;
    }

    /// <summary>Original target rectangle requested by the caller.</summary>
    public OfficeImagePlacement TargetPlacement { get; }

    /// <summary>Rectangle occupied by the visible image content after fit and source crop are applied.</summary>
    public OfficeImagePlacement VisiblePlacement { get; }

    /// <summary>Rectangle used to draw the underlying full image before source-crop clipping.</summary>
    public OfficeImagePlacement ImagePlacement { get; }

    /// <summary>Source-image crop fractions used by the plan.</summary>
    public OfficeImageSourceCrop SourceCrop { get; }

    /// <summary>Requested image fit mode.</summary>
    public OfficeImageFit Fit { get; }

    /// <summary>Whether the rendered image must be clipped to <see cref="TargetPlacement" />.</summary>
    public bool RequiresTargetClip { get; }

    /// <summary>
    /// Creates a render plan for coordinate systems where Y grows downward from the top edge.
    /// </summary>
    public static OfficeImageRenderPlan CreateTopLeft(
        double sourceWidth,
        double sourceHeight,
        double targetX,
        double targetY,
        double targetWidth,
        double targetHeight,
        OfficeImageFit fit = OfficeImageFit.Stretch,
        OfficeImageSourceCrop sourceCrop = default) =>
        Create(sourceWidth, sourceHeight, targetX, targetY, targetWidth, targetHeight, fit, sourceCrop, cropOffsetFromStart: sourceCrop.Top);

    /// <summary>
    /// Creates a render plan for coordinate systems where Y grows upward from the bottom edge.
    /// </summary>
    public static OfficeImageRenderPlan CreateBottomLeft(
        double sourceWidth,
        double sourceHeight,
        double targetX,
        double targetBottomY,
        double targetWidth,
        double targetHeight,
        OfficeImageFit fit = OfficeImageFit.Stretch,
        OfficeImageSourceCrop sourceCrop = default) =>
        Create(sourceWidth, sourceHeight, targetX, targetBottomY, targetWidth, targetHeight, fit, sourceCrop, cropOffsetFromStart: sourceCrop.Bottom);

    /// <summary>
    /// Returns a visible-placement projection suitable for top-left SVG and raster renderers.
    /// </summary>
    public OfficeImageProjection ToVisibleProjection(
        double rotationDegrees = 0D,
        double? rotationCenterX = null,
        double? rotationCenterY = null,
        bool flipHorizontal = false,
        bool flipVertical = false) =>
        new OfficeImageProjection(
            VisiblePlacement,
            SourceCrop,
            rotationDegrees,
            rotationCenterX,
            rotationCenterY,
            flipHorizontal,
            flipVertical);

    private static OfficeImageRenderPlan Create(
        double sourceWidth,
        double sourceHeight,
        double targetX,
        double targetY,
        double targetWidth,
        double targetHeight,
        OfficeImageFit fit,
        OfficeImageSourceCrop sourceCrop,
        double cropOffsetFromStart) {
        OfficeImagePlacement targetPlacement = new OfficeImagePlacement(targetX, targetY, targetWidth, targetHeight);
        EnsureValidFit(fit, nameof(fit));

        double visibleSourceWidth = sourceCrop.HasCrop ? sourceWidth * sourceCrop.VisibleWidth : sourceWidth;
        double visibleSourceHeight = sourceCrop.HasCrop ? sourceHeight * sourceCrop.VisibleHeight : sourceHeight;
        OfficeImagePlacement visiblePlacement = fit == OfficeImageFit.Stretch
            ? targetPlacement
            : OfficeImagePlacement.Fit(visibleSourceWidth, visibleSourceHeight, targetX, targetY, targetWidth, targetHeight, fit);

        OfficeImagePlacement imagePlacement = visiblePlacement;
        if (sourceCrop.HasCrop) {
            double imageWidth = visiblePlacement.Width / sourceCrop.VisibleWidth;
            double imageHeight = visiblePlacement.Height / sourceCrop.VisibleHeight;
            imagePlacement = new OfficeImagePlacement(
                visiblePlacement.X - (sourceCrop.Left * imageWidth),
                visiblePlacement.Y - (cropOffsetFromStart * imageHeight),
                imageWidth,
                imageHeight);
        }

        return new OfficeImageRenderPlan(
            targetPlacement,
            visiblePlacement,
            imagePlacement,
            sourceCrop,
            fit,
            fit == OfficeImageFit.Cover);
    }

    private static void EnsureValidFit(OfficeImageFit fit, string paramName) {
        if (fit != OfficeImageFit.Stretch && fit != OfficeImageFit.Contain && fit != OfficeImageFit.Cover) {
            throw new ArgumentOutOfRangeException(paramName, "Unsupported image fit mode.");
        }
    }
}
