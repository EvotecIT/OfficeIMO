using OfficeIMO.Drawing;

namespace OfficeIMO.Core.Internal;

/// <summary>
/// Shared Office 2016 SVG picture contract: the SVG part is referenced from an <c>asvg:svgBlip</c> blip extension,
/// while the ordinary blip carries a PNG fallback for older Office versions and other consumers.
/// </summary>
internal static class OfficeSvgPictureFallback {
    internal const string BlipExtensionUri = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}";
    private const double RasterScale = 2D;

    /// <summary>
    /// Geometry command budget for an Office SVG picture. Generated symbols such as large QR codes draw thousands of
    /// module runs; renderers that draw the picture's SVG use the same budget as its stored fallback.
    /// </summary>
    internal const int MaximumGeometryCommands = 250_000;

    /// <summary>Renders the PNG fallback for an SVG picture.</summary>
    /// <remarks>
    /// A lossy import would make a misleading fallback, so an SVG with unsupported features returns false and keeps
    /// its single blip.
    /// </remarks>
    internal static bool TryCreate(byte[] svg, out byte[] png, out double width, out double height) {
        png = Array.Empty<byte>();
        width = height = 0D;
        if (svg == null || svg.Length == 0) return false;
        try {
            // The pixel budget still bounds the raster.
            var options = new OfficeSvgDrawingReaderOptions { MaximumGeometryCommands = MaximumGeometryCommands };
            if (!OfficeSvgDrawingReader.TryRead(svg, options, out OfficeDrawing? drawing, out int unsupportedFeatures) || unsupportedFeatures > 0 ||
                drawing == null || drawing.Width <= 0D || drawing.Height <= 0D) return false;
            width = drawing.Width;
            height = drawing.Height;
            png = OfficeDrawingRasterRenderer.ToPng(drawing, RasterScale);
            return png.Length > 0;
        } catch (Exception ex) when (ex is NotSupportedException or InvalidOperationException or ArgumentException or InvalidDataException) {
            return false;
        }
    }
}
