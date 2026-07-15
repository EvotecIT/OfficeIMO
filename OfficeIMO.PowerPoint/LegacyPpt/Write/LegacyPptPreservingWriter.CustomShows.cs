using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Model;

namespace OfficeIMO.PowerPoint.LegacyPpt.Write {
    internal static partial class LegacyPptPreservingWriter {
        private static bool CustomShowsEqual(LegacyPptProjectionMap projectionMap,
            LegacyPptWriter.LegacyPptWriterCustomShowCatalog current) {
            if (projectionMap.CustomShows.Count != current.Shows.Count) return false;
            for (int showIndex = 0; showIndex < projectionMap.CustomShows.Count;
                 showIndex++) {
                LegacyPptCustomShow source = projectionMap.CustomShows[showIndex];
                LegacyPptWriter.LegacyPptWriterCustomShow target =
                    current.Shows[showIndex];
                if (!string.Equals(source.Name, target.Name,
                        StringComparison.Ordinal)
                    || source.SlideIds.Count != target.SlidePartUris.Count) {
                    return false;
                }
                for (int slideIndex = 0; slideIndex < source.SlideIds.Count;
                     slideIndex++) {
                    if (!projectionMap.TryGetSlide(source.SlideIds[slideIndex],
                            out LegacyPptSlideProjection? slide)
                        || slide == null
                        || !string.Equals(slide.SlidePartUri,
                            target.SlidePartUris[slideIndex],
                            StringComparison.Ordinal)) {
                        return false;
                    }
                }
            }
            return true;
        }

        private static bool TryRewriteCustomShows(LegacyPptPackage package,
            byte[]? currentDocumentBytes,
            LegacyPptWriter.LegacyPptWriterCustomShowCatalog customShows,
            PreservingInteractionContext interactionContext, out byte[] bytes) {
            bytes = Array.Empty<byte>();
            LegacyPptRecord document;
            if (currentDocumentBytes != null) {
                document = LegacyPptRecordReader.ReadSingle(currentDocumentBytes, 0,
                    new LegacyPptImportOptions());
            } else if (!TryReadDocument(package, out LegacyPptRecord? source)
                       || source == null) {
                return false;
            } else {
                document = source;
            }
            return LegacyPptWriter.TryBuildNamedShowsRecord(customShows,
                       interactionContext.ResolveSlideId,
                       out byte[] namedShows)
                && LegacyPptWriter.TryRewriteDocumentNamedShows(document,
                    namedShows, out bytes);
        }
    }
}
