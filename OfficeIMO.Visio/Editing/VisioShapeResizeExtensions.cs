using System;

namespace OfficeIMO.Visio;

/// <summary>Explicit resizing of an existing shape and its editable child tree.</summary>
public static class VisioShapeResizeExtensions {
    /// <summary>
    /// Resizes a page-owned shape in its local axes while preserving its pin, rotation,
    /// identities, text, formatting and master links. Child frames, supported geometry,
    /// image placement and connection points follow the resize.
    /// </summary>
    /// <param name="page">Page owning the shape, including nested shapes.</param>
    /// <param name="shape">Shape to resize.</param>
    /// <param name="width">Finite positive requested width.</param>
    /// <param name="height">Finite positive requested height.</param>
    /// <param name="unit">Size unit, or the page's default unit when omitted.</param>
    /// <returns>The same shape instance.</returns>
    /// <exception cref="NotSupportedException">The original frame is invalid, or native geometry or unequal scaling of a rotated child/text frame cannot be represented by this resize profile.</exception>
    /// <remarks>
    /// Preparation validates the complete tree and affected native connector paths before
    /// mutation. Arbitrary ShapeSheet recalculation and sheared child/text frames are not
    /// supported. Physical font sizes, margins and line weights remain unchanged.
    /// </remarks>
    public static VisioShape ResizeShape(this VisioPage page, VisioShape shape, double width, double height,
        VisioMeasurementUnit? unit = null) {
        if (page == null) throw new ArgumentNullException(nameof(page));
        if (shape == null) throw new ArgumentNullException(nameof(shape));
        if (!ReferenceEquals(shape.OwnerPage, page)) throw new InvalidOperationException("The shape is not part of this page.");
        VisioMeasurementUnit sizeUnit = unit ?? page.DefaultUnit;
        width = width.ToInches(sizeUnit); height = height.ToInches(sizeUnit);
        VisioShapeResizing.Resize(page, shape, width, height);
        return shape;
    }
}
