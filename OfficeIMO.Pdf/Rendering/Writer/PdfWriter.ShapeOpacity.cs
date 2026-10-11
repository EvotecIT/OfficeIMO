using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    /// <summary>Combines shape opacity with solid-color alpha consistently in body and running content.</summary>
    private static (double Fill, double Stroke) ResolveShapePaintOpacity(OfficeShape shape) {
        bool hasFill = (shape.FillColor.HasValue || shape.FillGradient != null || shape.FillRadialGradient != null)
            && shape.Kind != OfficeShapeKind.Line;
        bool hasStroke = shape.StrokeColor.HasValue && shape.StrokeWidth > 0;
        // Gradient stop alpha is handled by the gradient resource, independently of solid-color alpha.
        double fillColorOpacity = shape.FillGradient == null && shape.FillRadialGradient == null && shape.FillColor.HasValue
            ? shape.FillColor.Value.A / 255D : 1D;
        double strokeColorOpacity = shape.StrokeColor.HasValue ? shape.StrokeColor.Value.A / 255D : 1D;
        return (hasFill ? (shape.FillOpacity ?? 1D) * fillColorOpacity : 1D,
            hasStroke ? (shape.StrokeOpacity ?? 1D) * strokeColorOpacity : 1D);
    }
}
