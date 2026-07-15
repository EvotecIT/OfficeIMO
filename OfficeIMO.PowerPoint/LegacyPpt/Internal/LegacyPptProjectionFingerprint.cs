using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint.LegacyPpt.Internal {
    /// <summary>
    /// Fingerprints the projected package while normalizing geometry fields that the preservation writer owns.
    /// Any other package mutation prevents the conservative incremental-edit path.
    /// </summary>
    internal static class LegacyPptProjectionFingerprint {
        internal static string Create(PresentationDocument document, LegacyPptProjectionMap projectionMap) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            if (projectionMap == null) throw new ArgumentNullException(nameof(projectionMap));

            return PowerPointPackageFingerprint.Create(document, (part, root) => {
                if (part is SlidePart) NormalizeProjectedShapeGeometry(root, part.Uri, projectionMap);
            });
        }

        private static void NormalizeProjectedShapeGeometry(OpenXmlElement root, Uri partUri,
            LegacyPptProjectionMap projectionMap) {
            LegacyPptSlideProjection? slideProjection = projectionMap.Slides.FirstOrDefault(slide =>
                string.Equals(slide.SlidePartUri, partUri.ToString(), StringComparison.Ordinal));
            if (slideProjection == null) return;

            foreach (P.Shape shape in root.Descendants<P.Shape>()) {
                uint? shapeId = shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value;
                if (!shapeId.HasValue || !slideProjection.TryGetShape(shapeId.Value, out _)) continue;
                var transform = shape.ShapeProperties?.Transform2D;
                if (transform?.Offset != null) {
                    transform.Offset.X = 0L;
                    transform.Offset.Y = 0L;
                }
                if (transform?.Extents != null) {
                    transform.Extents.Cx = 0L;
                    transform.Extents.Cy = 0L;
                }
            }
        }
    }
}
