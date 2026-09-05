using System;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeDrawing {
    /// <summary>
    /// Adds vector paint that represents one logical text string. The paint remains native while
    /// PDF and SVG exporters can expose the string once for extraction and accessibility.
    /// </summary>
    public OfficeDrawing AddActualTextDrawing(
        string actualText,
        OfficeDrawing drawing,
        double actualTextAnchorX,
        double actualTextAnchorY) {
        if (string.IsNullOrEmpty(actualText)) throw new ArgumentException("Logical text cannot be empty.", nameof(actualText));
        if (drawing == null) throw new ArgumentNullException(nameof(drawing));
        ValidateFinite(actualTextAnchorX, nameof(actualTextAnchorX));
        ValidateFinite(actualTextAnchorY, nameof(actualTextAnchorY));
        if (drawing.Width > Width || drawing.Height > Height) {
            throw new ArgumentOutOfRangeException(nameof(drawing), "Logical text paint must fit inside the drawing bounds.");
        }
        Fonts.AddRange(drawing.Fonts);
        _elements.Add(new OfficeDrawingGroup(
            drawing,
            0D,
            0D,
            OfficeClipPath.Rectangle(drawing.Width, drawing.Height),
            0D,
            0D,
            null,
            actualText,
            actualTextAnchorX,
            actualTextAnchorY));
        return this;
    }
}
