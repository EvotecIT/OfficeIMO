using System;
using global::ChartForgeX.VisualArtifacts;
using OfficeIMO.Drawing;

namespace OfficeIMO.ChartForgeX;

/// <summary>Configures conversion from a ChartForgeX visual artifact to OfficeIMO document surfaces.</summary>
public sealed class OfficeVisualConversionOptions {
    private double _pointsPerPixel = 0.75D;
    private double? _widthPoints;
    private double? _heightPoints;
    private int _maximumSvgElements = OfficeSvgDrawingReaderOptions.DefaultMaximumElements;
    private double _maximumSvgViewportDimension = OfficeSvgDrawingReaderOptions.DefaultMaximumViewportDimension;
    private double _maximumSvgViewportPixels = OfficeSvgDrawingReaderOptions.DefaultMaximumViewportPixels;
    private OfficeVisualSvgPolicy _svgPolicy = OfficeVisualSvgPolicy.PreserveVector;

    /// <summary>Gets or sets ChartForgeX rendering options, including watermarks and PNG metadata.</summary>
    public VisualArtifactRenderOptions? RenderOptions { get; set; }

    /// <summary>Gets or sets SVG import behavior. The default preserves vector output and reports limitations.</summary>
    public OfficeVisualSvgPolicy SvgPolicy {
        get => _svgPolicy;
        set {
            if (!Enum.IsDefined(typeof(OfficeVisualSvgPolicy), value)) throw new ArgumentOutOfRangeException(nameof(SvgPolicy), value, "Unknown SVG fidelity policy.");
            _svgPolicy = value;
        }
    }

    /// <summary>Gets or sets the conversion factor from ChartForgeX pixels to Office points.</summary>
    public double PointsPerPixel {
        get => _pointsPerPixel;
        set {
            ValidatePositiveFinite(value, nameof(PointsPerPixel));
            _pointsPerPixel = value;
        }
    }

    /// <summary>Gets or sets an optional output width in points. Aspect ratio is preserved when height is omitted.</summary>
    public double? WidthPoints {
        get => _widthPoints;
        set {
            if (value.HasValue) ValidatePositiveFinite(value.Value, nameof(WidthPoints));
            _widthPoints = value;
        }
    }

    /// <summary>Gets or sets an optional output height in points. Aspect ratio is preserved when width is omitted.</summary>
    public double? HeightPoints {
        get => _heightPoints;
        set {
            if (value.HasValue) ValidatePositiveFinite(value.Value, nameof(HeightPoints));
            _heightPoints = value;
        }
    }

    /// <summary>Gets or sets the SVG importer element limit.</summary>
    public int MaximumSvgElements {
        get => _maximumSvgElements;
        set {
            if (value <= 0 || value > OfficeSvgDrawingReaderOptions.MaximumAllowedElements) {
                throw new ArgumentOutOfRangeException(nameof(MaximumSvgElements), value, "SVG element limit is outside the supported range.");
            }
            _maximumSvgElements = value;
        }
    }

    /// <summary>
    /// Gets or sets the maximum accepted SVG viewport width or height. The safe default is 8,192;
    /// increase this only for trusted input.
    /// </summary>
    public double MaximumSvgViewportDimension {
        get => _maximumSvgViewportDimension;
        set {
            ValidateSvgLimit(value, OfficeSvgDrawingReaderOptions.MaximumAllowedViewportDimension, nameof(MaximumSvgViewportDimension));
            _maximumSvgViewportDimension = value;
        }
    }

    /// <summary>
    /// Gets or sets the maximum accepted SVG viewport area in pixels. The safe default is 16 megapixels;
    /// increase this only for trusted input.
    /// </summary>
    public double MaximumSvgViewportPixels {
        get => _maximumSvgViewportPixels;
        set {
            ValidateSvgLimit(value, OfficeSvgDrawingReaderOptions.MaximumAllowedViewportPixels, nameof(MaximumSvgViewportPixels));
            _maximumSvgViewportPixels = value;
        }
    }

    internal (double Width, double Height) ResolveSize(double sourceWidth, double sourceHeight) {
        ValidatePositiveFinite(sourceWidth, nameof(sourceWidth));
        ValidatePositiveFinite(sourceHeight, nameof(sourceHeight));
        if (WidthPoints.HasValue && HeightPoints.HasValue) return (WidthPoints.Value, HeightPoints.Value);
        if (WidthPoints.HasValue) return (WidthPoints.Value, WidthPoints.Value * sourceHeight / sourceWidth);
        if (HeightPoints.HasValue) return (HeightPoints.Value * sourceWidth / sourceHeight, HeightPoints.Value);
        return (sourceWidth * PointsPerPixel, sourceHeight * PointsPerPixel);
    }

    private static void ValidatePositiveFinite(double value, string parameterName) {
        if (value <= 0D || double.IsNaN(value) || double.IsInfinity(value)) {
            throw new ArgumentOutOfRangeException(parameterName, value, "Value must be positive and finite.");
        }
    }

    private static void ValidateSvgLimit(double value, double maximum, string parameterName) {
        ValidatePositiveFinite(value, parameterName);
        if (value > maximum) {
            throw new ArgumentOutOfRangeException(parameterName, value, "SVG limit exceeds the supported hard maximum.");
        }
    }
}
