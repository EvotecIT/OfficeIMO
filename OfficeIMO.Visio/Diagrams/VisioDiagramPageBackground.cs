using OfficeIMO.Drawing;

namespace OfficeIMO.Visio.Diagrams;

/// <summary>Persists a diagram's page fill without adding a background page or affecting its layout.</summary>
internal static class VisioDiagramPageBackground {
    internal static void Apply(VisioPage page, OfficeColor? color) {
        if (!color.HasValue || color.Value.A == 0) return;
        // Apply after fitting and routing so the canvas never acts as a layout obstacle.
        VisioShape background = page.AddRectangle(page.Width / 2D, page.Height / 2D,
            page.Width, page.Height, string.Empty, VisioMeasurementUnit.Inches);
        background.NameU = "OfficeIMO Page Background";
        background.FillColor = color.Value;
        background.LinePattern = 0;
        background.MarkAsBackgroundSurface();
        background.Protect(protection => protection.Size().Position().Selection());
        // Shapes are saved and drawn in list order. The fill precedes every editable object.
        page.Shapes.Remove(background);
        page.Shapes.Insert(0, background);
    }
}
