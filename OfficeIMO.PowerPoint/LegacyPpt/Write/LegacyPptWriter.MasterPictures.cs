using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.PowerPoint.LegacyPpt.Capabilities;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint.LegacyPpt.Write {
    internal static partial class LegacyPptWriter {
        private static bool TryValidateMaterializedLayoutPictures(
            PowerPointPresentation presentation,
            ISet<OpenXmlElement> materializedPictures,
            out string? reason) {
            IEnumerable<SlideLayoutPart> layoutParts = presentation
                .OpenXmlDocument.PresentationPart?.SlideMasterParts
                .SelectMany(master => master.SlideLayoutParts)
                ?? Enumerable.Empty<SlideLayoutPart>();
            foreach (SlideLayoutPart part in layoutParts) {
                P.Picture? unmaterialized = part.SlideLayout?.CommonSlideData?
                    .ShapeTree?.Descendants<P.Picture>().FirstOrDefault(
                        picture => !materializedPictures.Contains(picture));
                if (unmaterialized == null) continue;
                string name = part.SlideLayout?.CommonSlideData?.Name?.Value
                    ?? part.Uri.ToString();
                string pictureName = unmaterialized
                    .NonVisualPictureProperties?.NonVisualDrawingProperties?
                    .Name?.Value ?? "unnamed picture";
                reason = $"Slide layout '{name}' contains '{pictureName}', which does not materialize into any owning slide; classic binary PowerPoint has no independent ordinary-layout persist object in which to retain it.";
                return false;
            }
            reason = null;
            return true;
        }

        private static bool TryAddMasterPictures(
            LegacyPptWriterPictureCatalog catalog,
            IEnumerable<PowerPointShape> shapes,
            out LegacyPptFeature failureFeature, out string? reason) {
            if (catalog == null) throw new ArgumentNullException(
                nameof(catalog));
            if (shapes == null) throw new ArgumentNullException(nameof(shapes));
            IReadOnlyList<PowerPointShape> flattened =
                FlattenShapeTreeForWrite(shapes, out reason);
            if (reason != null) {
                failureFeature = LegacyPptFeature.Groups;
                return false;
            }
            foreach (PowerPointPicture picture in flattened
                         .OfType<PowerPointPicture>()) {
                if (!TryReadPicture(picture, out byte[] imageBytes,
                        out string? contentType, out reason)
                    || !catalog.TryAdd(picture, imageBytes,
                        contentType!, out reason)) {
                    failureFeature = LegacyPptFeature.RasterPictures;
                    return false;
                }
            }
            failureFeature = LegacyPptFeature.RasterPictures;
            reason = null;
            return true;
        }

        internal static bool IsSupportedMasterShape(PowerPointShape shape) =>
            IsSupportedShape(shape, includeOleObjects: false,
                includePictures: true);
    }
}
