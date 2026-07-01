namespace OfficeIMO.Drawing;

/// <summary>
/// Describes one positioned raster or SVG layer in a composed dependency-free image.
/// </summary>
public sealed class OfficeImageLayer {
    private OfficeImageLayer(double x, double y, double width, double height, OfficeRasterImage? rasterImage, string? svgInnerContent) {
        ValidateCoordinate(x, nameof(x));
        ValidateCoordinate(y, nameof(y));
        ValidatePositiveDimension(width, nameof(width));
        ValidatePositiveDimension(height, nameof(height));

        X = x;
        Y = y;
        Width = width;
        Height = height;
        RasterImage = rasterImage;
        SvgInnerContent = svgInnerContent;
    }

    /// <summary>Layer x-coordinate in output pixels or SVG user units.</summary>
    public double X { get; }

    /// <summary>Layer y-coordinate in output pixels or SVG user units.</summary>
    public double Y { get; }

    /// <summary>Layer width in output pixels or SVG user units.</summary>
    public double Width { get; }

    /// <summary>Layer height in output pixels or SVG user units.</summary>
    public double Height { get; }

    /// <summary>Raster image to draw when composing PNG output.</summary>
    public OfficeRasterImage? RasterImage { get; }

    /// <summary>Inner SVG markup to nest when composing SVG output.</summary>
    public string? SvgInnerContent { get; }

    /// <summary>
    /// Creates a positioned raster image layer.
    /// </summary>
    public static OfficeImageLayer FromRaster(OfficeRasterImage rasterImage, double x, double y, double width, double height) {
        if (rasterImage == null) {
            throw new System.ArgumentNullException(nameof(rasterImage));
        }

        return new OfficeImageLayer(x, y, width, height, rasterImage, null);
    }

    /// <summary>
    /// Creates a positioned SVG layer from inner SVG markup.
    /// </summary>
    public static OfficeImageLayer FromSvgInner(string svgInnerContent, double x, double y, double width, double height) {
        if (svgInnerContent == null) {
            throw new System.ArgumentNullException(nameof(svgInnerContent));
        }

        return new OfficeImageLayer(x, y, width, height, null, svgInnerContent);
    }

    private static void ValidateCoordinate(double value, string paramName) {
        if (double.IsNaN(value) || double.IsInfinity(value)) {
            throw new System.ArgumentOutOfRangeException(paramName, "Layer coordinates must be finite numbers.");
        }
    }

    private static void ValidatePositiveDimension(double value, string paramName) {
        if (value <= 0D || double.IsNaN(value) || double.IsInfinity(value)) {
            throw new System.ArgumentOutOfRangeException(paramName, "Layer dimensions must be finite positive numbers.");
        }
    }
}
