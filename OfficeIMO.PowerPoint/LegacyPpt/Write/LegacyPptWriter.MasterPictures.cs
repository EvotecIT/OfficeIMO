using OfficeIMO.OpenXml.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Capabilities;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;

namespace OfficeIMO.PowerPoint.LegacyPpt.Write {
    internal static partial class LegacyPptWriter {
        /// <summary>
        /// Fails when a layout of the view holds a picture (<c>p:pic</c>) that no slide materialized: one missing from
        /// <paramref name="materializedPictures"/>, which holds pictures of the same view.
        /// </summary>
        private static bool TryValidateMaterializedLayoutPictures(
            LegacyPptPackageView view,
            ISet<OfficeXmlNode> materializedPictures,
            out string? reason) {
            foreach (LegacyPptPartView master in SlideMasterPartsInRelationshipOrder(view)) {
                foreach (LegacyPptPartView part in master.RelatedParts(
                             LegacyPptPackageView.SlideLayoutRelationshipType)) {
                    OfficeXmlNode unmaterialized = ShapeTree(part).Descendants(Pml, "pic")
                        .FirstOrDefault(picture => !materializedPictures.Contains(picture));
                    if (unmaterialized.IsNone) continue;
                    string name = CommonSlideData(part).Attribute("name")
                        ?? part.Name;
                    string pictureName = unmaterialized.Path(Pml, "nvPicPr", "cNvPr")
                        .Attribute("name") ?? "unnamed picture";
                    reason = $"Slide layout '{name}' contains '{pictureName}', which does not materialize into any owning slide; classic binary PowerPoint has no independent ordinary-layout persist object in which to retain it.";
                    return false;
                }
            }
            reason = null;
            return true;
        }

        private static bool TryAddMasterPictures(
            LegacyPptWriterPictureCatalog catalog,
            IEnumerable<PowerPointShape> shapes,
            LegacyPptWriterFontCatalog tableFonts,
            LegacyPptWriterPictureBulletCatalog pictureBullets,
            bool convertUnsupportedTables,
            out LegacyPptFeature failureFeature, out string? reason) {
            ArgumentNullException.ThrowIfNull(catalog);
            ArgumentNullException.ThrowIfNull(shapes);
            IReadOnlyList<PowerPointShape> flattened =
                FlattenShapeTreeForWrite(shapes, out reason);
            if (reason != null) {
                failureFeature = LegacyPptFeature.Groups;
                return false;
            }
            foreach (PowerPointShape shape in flattened) {
                byte[] imageBytes;
                string contentType;
                if (shape is PowerPointImage picture) {
                    if (!TryReadPicture(picture, out imageBytes,
                            out string? pictureContentType, out reason)) {
                        failureFeature = LegacyPptFeature.RasterPictures;
                        return false;
                    }
                    contentType = pictureContentType!;
                } else if (convertUnsupportedTables
                    && shape is PowerPointTable table
                    && !TryReadTableForWrite(table, tableFonts,
                        pictureBullets, out _)) {
                    if (!TryRenderTablePicture(table, out imageBytes,
                            out reason)) {
                        failureFeature = LegacyPptFeature.Tables;
                        return false;
                    }
                    contentType = "image/png";
                } else {
                    continue;
                }
                if (!catalog.TryAdd(shape, imageBytes, contentType,
                        out reason)) {
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
