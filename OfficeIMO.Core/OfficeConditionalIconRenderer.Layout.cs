using System;
using System.Text;

namespace OfficeIMO.Drawing;

public static partial class OfficeConditionalIconRenderer {
    /// <summary>
    /// Gets local CSS-pixel bounds using the current estimated 3px left inset and 16px size cap.
    /// The icon is bounded by the cell; exact native Excel inset and size are not qualified.
    /// </summary>
    public static (double X, double Y, double Size) GetCellBounds(double width, double height, double scale = 1D) {
        if (double.IsNaN(width) || double.IsInfinity(width) || width <= 0D) throw new ArgumentOutOfRangeException(nameof(width));
        if (double.IsNaN(height) || double.IsInfinity(height) || height <= 0D) throw new ArgumentOutOfRangeException(nameof(height));
        if (double.IsNaN(scale) || double.IsInfinity(scale) || scale <= 0D) throw new ArgumentOutOfRangeException(nameof(scale));
        double x = Math.Min(3D, width);
        double size = Math.Max(0D, Math.Min(16D, Math.Min(height, width - x)));
        return (x * scale, (height - size) / 2D * scale, size * scale);
    }

    /// <summary>Creates a vector drawing using the same artwork as SVG and raster conditional icons.</summary>
    public static OfficeDrawing CreateDrawing(OfficeConditionalIconKind kind) {
        const double size = 16D;
        if (kind == OfficeConditionalIconKind.None) return new OfficeDrawing(size, size);
        var svg = new StringBuilder("<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"16\" height=\"16\" viewBox=\"0 0 16 16\">");
        AppendSvg(svg, 0D, 0D, size, kind);
        svg.Append("</svg>");
        if (!OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg.ToString()), out OfficeDrawing? drawing) || drawing == null) {
            throw new InvalidOperationException("Conditional icon artwork could not be read as a vector drawing.");
        }
        return drawing;
    }
}
