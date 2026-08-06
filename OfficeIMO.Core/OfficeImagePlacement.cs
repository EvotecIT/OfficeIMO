using System;

namespace OfficeIMO.Drawing;

/// <summary>
/// Describes the destination rectangle used when placing an image into a target box.
/// </summary>
public readonly struct OfficeImagePlacement {
    /// <summary>
    /// Creates a new image placement rectangle.
    /// </summary>
    public OfficeImagePlacement(double x, double y, double width, double height) {
        EnsureFinite(x, nameof(x));
        EnsureFinite(y, nameof(y));
        EnsurePositive(width, nameof(width));
        EnsurePositive(height, nameof(height));

        X = x;
        Y = y;
        Width = width;
        Height = height;
    }

    /// <summary>Destination left coordinate.</summary>
    public double X { get; }

    /// <summary>Destination top coordinate.</summary>
    public double Y { get; }

    /// <summary>Destination width.</summary>
    public double Width { get; }

    /// <summary>Destination height.</summary>
    public double Height { get; }

    /// <summary>
    /// Fits a source image into a target rectangle using the requested image fit mode.
    /// </summary>
    /// <param name="sourceWidth">Source image width.</param>
    /// <param name="sourceHeight">Source image height.</param>
    /// <param name="targetX">Target rectangle left coordinate.</param>
    /// <param name="targetY">Target rectangle top coordinate.</param>
    /// <param name="targetWidth">Target rectangle width.</param>
    /// <param name="targetHeight">Target rectangle height.</param>
    /// <param name="fit">Image fit mode.</param>
    /// <returns>The destination rectangle to render.</returns>
    public static OfficeImagePlacement Fit(
        double sourceWidth,
        double sourceHeight,
        double targetX,
        double targetY,
        double targetWidth,
        double targetHeight,
        OfficeImageFit fit = OfficeImageFit.Stretch) {
        EnsureFinite(targetX, nameof(targetX));
        EnsureFinite(targetY, nameof(targetY));
        EnsurePositive(targetWidth, nameof(targetWidth));
        EnsurePositive(targetHeight, nameof(targetHeight));
        EnsureValidFit(fit, nameof(fit));

        if (fit == OfficeImageFit.Stretch) {
            return new OfficeImagePlacement(targetX, targetY, targetWidth, targetHeight);
        }

        EnsurePositive(sourceWidth, nameof(sourceWidth));
        EnsurePositive(sourceHeight, nameof(sourceHeight));

        double scaleX = targetWidth / sourceWidth;
        double scaleY = targetHeight / sourceHeight;
        double scale = fit == OfficeImageFit.Contain ? Math.Min(scaleX, scaleY) : Math.Max(scaleX, scaleY);
        double width = sourceWidth * scale;
        double height = sourceHeight * scale;

        return new OfficeImagePlacement(
            targetX + ((targetWidth - width) / 2D),
            targetY + ((targetHeight - height) / 2D),
            width,
            height);
    }

    /// <summary>
    /// Calculates the visible aspect-ratio change between a source image and a target rectangle.
    /// </summary>
    /// <param name="sourceWidth">Source image width.</param>
    /// <param name="sourceHeight">Source image height.</param>
    /// <param name="targetWidth">Target rectangle width.</param>
    /// <param name="targetHeight">Target rectangle height.</param>
    /// <returns>A ratio greater than or equal to 1 where 1 means the source and target aspect ratios match.</returns>
    public static double GetAspectRatioDistortionRatio(
        double sourceWidth,
        double sourceHeight,
        double targetWidth,
        double targetHeight) {
        EnsurePositive(sourceWidth, nameof(sourceWidth));
        EnsurePositive(sourceHeight, nameof(sourceHeight));
        EnsurePositive(targetWidth, nameof(targetWidth));
        EnsurePositive(targetHeight, nameof(targetHeight));

        double sourceAspect = sourceWidth / sourceHeight;
        double targetAspect = targetWidth / targetHeight;
        return Math.Max(sourceAspect, targetAspect) / Math.Min(sourceAspect, targetAspect);
    }

    /// <summary>
    /// Returns whether a source image would visibly distort when stretched into a target rectangle.
    /// </summary>
    /// <param name="sourceWidth">Source image width.</param>
    /// <param name="sourceHeight">Source image height.</param>
    /// <param name="targetWidth">Target rectangle width.</param>
    /// <param name="targetHeight">Target rectangle height.</param>
    /// <param name="thresholdRatio">Minimum distortion ratio considered visible. Values below 1 are invalid.</param>
    public static bool ExceedsAspectRatioDistortion(
        double sourceWidth,
        double sourceHeight,
        double targetWidth,
        double targetHeight,
        double thresholdRatio) {
        if (thresholdRatio < 1D || double.IsNaN(thresholdRatio) || double.IsInfinity(thresholdRatio)) {
            throw new ArgumentOutOfRangeException(nameof(thresholdRatio), "Aspect-ratio distortion threshold must be finite and greater than or equal to 1.");
        }

        return GetAspectRatioDistortionRatio(sourceWidth, sourceHeight, targetWidth, targetHeight) > thresholdRatio;
    }

    /// <summary>
    /// Converts the placement into a tuple.
    /// </summary>
    public (double X, double Y, double Width, double Height) ToTuple() => (X, Y, Width, Height);

    private static void EnsureValidFit(OfficeImageFit fit, string paramName) {
        if (fit != OfficeImageFit.Stretch && fit != OfficeImageFit.Contain && fit != OfficeImageFit.Cover) {
            throw new ArgumentOutOfRangeException(paramName, "Unsupported image fit mode.");
        }
    }

    private static void EnsurePositive(double value, string paramName) {
        if (value <= 0D || double.IsNaN(value) || double.IsInfinity(value)) {
            throw new ArgumentOutOfRangeException(paramName, "Value must be positive and finite.");
        }
    }

    private static void EnsureFinite(double value, string paramName) {
        if (double.IsNaN(value) || double.IsInfinity(value)) {
            throw new ArgumentOutOfRangeException(paramName, "Value must be finite.");
        }
    }
}
