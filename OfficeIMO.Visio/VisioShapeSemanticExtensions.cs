using System;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Helpers for marking shapes with OfficeIMO semantic roles used by layout, routing, inspection, and quality analysis.
    /// </summary>
    public static class VisioShapeSemanticExtensions {
        /// <summary>
        /// Marks a shape as a background surface such as a zone, subnet, lane, band, or grouping region.
        /// </summary>
        /// <param name="shape">Shape to mark.</param>
        public static VisioShape MarkAsBackgroundSurface(this VisioShape shape) {
            if (shape == null) {
                throw new ArgumentNullException(nameof(shape));
            }

            shape.SetUserCell(VisioSemanticUserCells.Kind, VisioSemanticUserCells.BackgroundSurfaceKind, "STR", prompt: "OfficeIMO semantic kind");
            return shape;
        }

        /// <summary>
        /// Marks a generated label, title, legend item, or caption as a diagram adornment that quality checks and routing can treat as non-domain content.
        /// </summary>
        /// <param name="shape">Shape to mark.</param>
        public static VisioShape MarkAsGeneratedDiagramAdornment(this VisioShape shape) {
            VisioSemanticUserCells.MarkGeneratedAdornment(shape);
            return shape;
        }
    }
}
