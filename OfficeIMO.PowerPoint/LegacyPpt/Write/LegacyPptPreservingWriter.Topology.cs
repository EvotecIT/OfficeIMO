using OfficeIMO.PowerPoint.LegacyPpt.Internal;

namespace OfficeIMO.PowerPoint.LegacyPpt.Write {
    internal static partial class LegacyPptPreservingWriter {
        private static bool TryRewriteDocumentSlideOrder(LegacyPptPackage package,
            LegacyPptProjectionMap projectionMap, IReadOnlyList<LegacyPptSlideProjection> slideOrder,
            out byte[] bytes) {
            bytes = Array.Empty<byte>();
            if (!package.PersistObjects.TryGetValue(package.DocumentPersistId,
                    out LegacyPptPersistObject? documentObject) || documentObject == null) {
                return false;
            }
            LegacyPptRecord document = LegacyPptRecordReader.ReadSingle(documentObject.RecordBytes, 0,
                new LegacyPptImportOptions());
            if (slideOrder.Count < projectionMap.Slides.Count
                && document.DescendantsAndSelf().Any(record => record.Type == RecordNamedShows)) {
                return false;
            }
            LegacyPptRecord? slideList = document.Children.FirstOrDefault(child =>
                child.Type == RecordSlideListWithText && child.Instance == 0);
            if (slideList == null || !TryReorderSlideList(slideList, projectionMap, slideOrder,
                    out byte[] reorderedSlideList)) {
                return false;
            }

            var children = new List<byte[]>(document.Children.Count);
            foreach (LegacyPptRecord child in document.Children) {
                children.Add(ReferenceEquals(child, slideList) ? reorderedSlideList : child.CopyRecordBytes());
            }
            bytes = BuildRecord(document.Version, document.Instance, document.Type, Concat(children));
            return true;
        }

        private static bool TryReorderSlideList(LegacyPptRecord slideList,
            LegacyPptProjectionMap projectionMap, IReadOnlyList<LegacyPptSlideProjection> slideOrder,
            out byte[] bytes) {
            var prefix = new List<byte[]>();
            var groups = new Dictionary<uint, List<byte[]>>();
            List<byte[]>? currentGroup = null;
            uint currentPersistId = 0;
            foreach (LegacyPptRecord child in slideList.Children) {
                if (child.Type == RecordSlidePersistAtom) {
                    if (currentGroup != null && !TryAddSlideGroup(groups, currentPersistId, currentGroup)) {
                        bytes = slideList.CopyRecordBytes();
                        return false;
                    }
                    if (child.PayloadLength < 4) {
                        bytes = slideList.CopyRecordBytes();
                        return false;
                    }
                    currentPersistId = child.ReadUInt32(0);
                    currentGroup = new List<byte[]> { child.CopyRecordBytes() };
                } else if (currentGroup == null) {
                    prefix.Add(child.CopyRecordBytes());
                } else {
                    currentGroup.Add(child.CopyRecordBytes());
                }
            }
            if (currentGroup != null && !TryAddSlideGroup(groups, currentPersistId, currentGroup)) {
                bytes = slideList.CopyRecordBytes();
                return false;
            }
            if (groups.Count != projectionMap.Slides.Count || slideOrder.Count > projectionMap.Slides.Count
                || slideOrder.Select(slide => slide.PersistId).Distinct().Count() != slideOrder.Count) {
                bytes = slideList.CopyRecordBytes();
                return false;
            }

            var reordered = new List<byte[]>(slideList.Children.Count);
            reordered.AddRange(prefix);
            foreach (LegacyPptSlideProjection slide in slideOrder) {
                if (!groups.TryGetValue(slide.PersistId, out List<byte[]>? group) || group == null) {
                    bytes = slideList.CopyRecordBytes();
                    return false;
                }
                reordered.AddRange(group);
            }
            bytes = BuildRecord(slideList.Version, slideList.Instance, slideList.Type, Concat(reordered));
            return true;
        }

        private static bool TryAddSlideGroup(IDictionary<uint, List<byte[]>> groups, uint persistId,
            List<byte[]> group) {
            if (groups.ContainsKey(persistId)) return false;
            groups.Add(persistId, group);
            return true;
        }

        private static IReadOnlyList<uint> GetCurrentSlideIds(PowerPointPresentation presentation,
            LegacyPptProjectionMap projectionMap) {
            var slideIds = new List<uint>(presentation.Slides.Count);
            foreach (PowerPointSlide slide in presentation.Slides) {
                if (!projectionMap.TryGetSlide(slide, out LegacyPptSlideProjection? projection)
                    || projection == null) {
                    throw new InvalidOperationException("The projected slide map changed during binary saving.");
                }
                slideIds.Add(projection.SlideId);
            }
            return slideIds;
        }
    }
}
