using System;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    // Preserve the PDF painter's established presets. Other formats use their
    // own dash lengths and cap defaults through the shared drawing model.
    private static double[]? PdfPresetDashPattern(OfficeStrokeDashStyle style, double width) {
        double unit = Math.Max(.1D, width);
        return style switch {
            OfficeStrokeDashStyle.Dash => new[] { 3 * unit, 1.5 * unit },
            OfficeStrokeDashStyle.Dot => new[] { unit, 1.5 * unit },
            OfficeStrokeDashStyle.DashDot => new[] { 3 * unit, 1.5 * unit, unit, 1.5 * unit },
            _ => null
        };
    }

    private static OfficeStrokeLineCap PdfPresetLineCap(OfficeStrokeDashStyle style) =>
        style is OfficeStrokeDashStyle.Dot or OfficeStrokeDashStyle.DashDot
            ? OfficeStrokeLineCap.Round : OfficeStrokeLineCap.Butt;

    private static double PdfPaintMiterLimit(OfficeShape shape) =>
        // The compact painter retains PDF's implicit limit of ten; the exact
        // painter emits the model's limit, including an explicit miter join.
        shape.Transform.HasValue || RequiresExactStrokeWriter(shape) ? shape.StrokeMiterLimit : 10D;
}
