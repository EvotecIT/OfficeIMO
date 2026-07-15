using System.Collections.ObjectModel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint.LegacyPpt.Internal {
    /// <summary>
    /// Retains a global package guard and independent normalized slide guards. Existing projected slides can be
    /// reordered or removed, while any unsupported mutation to a retained slide or shared package part is rejected.
    /// </summary>
    internal sealed class LegacyPptProjectionFingerprint {
        private readonly IReadOnlyDictionary<string, string> _slides;

        private LegacyPptProjectionFingerprint(string global, IReadOnlyDictionary<string, string> slides) {
            Global = global;
            _slides = new ReadOnlyDictionary<string, string>(slides.ToDictionary(
                pair => pair.Key, pair => pair.Value, StringComparer.Ordinal));
        }

        internal string Global { get; }

        internal static LegacyPptProjectionFingerprint Create(PresentationDocument document,
            LegacyPptProjectionMap projectionMap) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            if (projectionMap == null) throw new ArgumentNullException(nameof(projectionMap));
            var slides = new Dictionary<string, string>(StringComparer.Ordinal);
            foreach (SlidePart slidePart in document.PresentationPart?.SlideParts ?? Enumerable.Empty<SlidePart>()) {
                if (projectionMap.Slides.Any(slide => string.Equals(slide.SlidePartUri,
                        slidePart.Uri.ToString(), StringComparison.Ordinal))) {
                    slides.Add(slidePart.Uri.ToString(), CreateSlide(document, slidePart, projectionMap));
                }
            }
            if (slides.Count != projectionMap.Slides.Count) {
                throw new InvalidDataException("The projected slide fingerprint set is incomplete.");
            }
            return new LegacyPptProjectionFingerprint(CreateGlobal(document), slides);
        }

        internal bool Matches(PresentationDocument document, LegacyPptProjectionMap projectionMap) {
            if (!string.Equals(Global, CreateGlobal(document), StringComparison.Ordinal)) return false;
            SlidePart[] currentSlides = document.PresentationPart?.SlideParts.ToArray() ?? Array.Empty<SlidePart>();
            foreach (SlidePart slidePart in currentSlides) {
                string uri = slidePart.Uri.ToString();
                if (_slides.TryGetValue(uri, out string? expected)
                    && !string.Equals(expected, CreateSlide(document, slidePart, projectionMap),
                        StringComparison.Ordinal)) {
                    return false;
                }
            }
            return true;
        }

        private static string CreateGlobal(PresentationDocument document) =>
            PowerPointPackageFingerprint.Create(document,
                (part, root) => {
                    if (part is PresentationPart) NormalizePresentationTopology(root);
                },
                part => !(part is SlidePart),
                (owner, relationship) => !(relationship.OpenXmlPart is SlidePart));

        private static string CreateSlide(PresentationDocument document, SlidePart slidePart,
            LegacyPptProjectionMap projectionMap) => PowerPointPackageFingerprint.Create(document,
            (part, root) => NormalizeProjectedSlide(root, part.Uri, projectionMap),
            part => string.Equals(part.Uri.ToString(), slidePart.Uri.ToString(), StringComparison.Ordinal));

        private static void NormalizePresentationTopology(OpenXmlElement root) {
            if (root is not P.Presentation presentation || presentation.SlideIdList == null) return;
            presentation.SlideIdList.RemoveAllChildren<P.SlideId>();
        }

        private static void NormalizeProjectedSlide(OpenXmlElement root, Uri partUri,
            LegacyPptProjectionMap projectionMap) {
            LegacyPptSlideProjection? slideProjection = projectionMap.Slides.FirstOrDefault(slide =>
                string.Equals(slide.SlidePartUri, partUri.ToString(), StringComparison.Ordinal));
            if (slideProjection == null) return;
            if (root is P.Slide slideRoot) slideRoot.Show = null;

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
                if (shape.TextBody != null) shape.TextBody.RemoveAllChildren<A.Paragraph>();
            }
            foreach (P.Picture picture in root.Descendants<P.Picture>()) {
                uint? shapeId = picture.NonVisualPictureProperties?.NonVisualDrawingProperties?.Id?.Value;
                if (!shapeId.HasValue || !slideProjection.TryGetShape(shapeId.Value, out _)) continue;
                A.Transform2D? transform = picture.ShapeProperties?.Transform2D;
                if (transform?.Offset != null) {
                    transform.Offset.X = 0L;
                    transform.Offset.Y = 0L;
                }
                if (transform?.Extents != null) {
                    transform.Extents.Cx = 0L;
                    transform.Extents.Cy = 0L;
                }
            }
            foreach (P.ConnectionShape connection in root.Descendants<P.ConnectionShape>()) {
                uint? shapeId = connection.NonVisualConnectionShapeProperties?
                    .NonVisualDrawingProperties?.Id?.Value;
                if (!shapeId.HasValue || !slideProjection.TryGetShape(shapeId.Value, out _)) continue;
                A.Transform2D? transform = connection.ShapeProperties?.Transform2D;
                if (transform?.Offset != null) {
                    transform.Offset.X = 0L;
                    transform.Offset.Y = 0L;
                }
                if (transform?.Extents != null) {
                    transform.Extents.Cx = 0L;
                    transform.Extents.Cy = 0L;
                }
            }
            foreach (P.GroupShape group in root.Descendants<P.GroupShape>()) {
                uint? shapeId = group.NonVisualGroupShapeProperties?.NonVisualDrawingProperties?.Id?.Value;
                if (!shapeId.HasValue || !slideProjection.TryGetShape(shapeId.Value, out _)) continue;
                A.TransformGroup? transform = group.GroupShapeProperties?.TransformGroup;
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
