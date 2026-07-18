using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;

namespace OfficeIMO.PowerPoint.LegacyPpt.Write {
    internal static partial class LegacyPptPreservingWriter {
        private static bool TryAppendNewSlides(LegacyPptPackage package,
            LegacyPptProjectionMap projectionMap, IReadOnlyList<PowerPointSlide> addedSlides,
            IDictionary<uint, byte[]> rewritten,
            LegacyPptWriter.LegacyPptWriterInteractionCatalog interactionCatalog,
            PreservingInteractionContext interactionContext,
            LegacyPptWriter.LegacyPptWriterFontCatalog fonts,
            LegacyPptWriter.LegacyPptWriterPictureBulletCatalog pictureBullets,
            out IReadOnlyList<uint> addedSlideIds) {
            var slideIds = new List<uint>(addedSlides.Count);
            addedSlideIds = slideIds;
            if (addedSlides.Count == 0 || projectionMap.Slides.Count == 0
                || !TryReadDocument(package, out LegacyPptRecord? document)
                || document == null) {
                return false;
            }

            LegacyPptRecord? slideList = document.Children.FirstOrDefault(child =>
                child.Type == RecordSlideListWithText && child.Instance == 0);
            LegacyPptRecord? dgg = document.DescendantsAndSelf().FirstOrDefault(record =>
                record.Type == OfficeArtDgg);
            if (slideList == null || dgg == null || !TryReadDggClusters(dgg,
                    out List<KeyValuePair<uint, uint>> clusters)) {
                return false;
            }

            uint persistId = package.PersistObjects.Count == 0 ? 1U : package.PersistObjects.Keys.Max();
            uint slideId = projectionMap.Slides.Max(slide => slide.SlideId);
            uint drawingId = clusters.Count == 0 ? 0U : clusters.Max(cluster => cluster.Key);
            LegacyPptWriter.LegacyPptWriterInteractionCatalog remappedInteractions =
                interactionContext.RemapCatalog(addedSlides, interactionCatalog);
            var appendedSlideAtoms = new List<byte[]>(addedSlides.Count);
            foreach (PowerPointSlide slide in addedSlides) {
                if (persistId >= 0x000FFFFE || slideId >= 0x7FFFFFFF || drawingId >= 0x003FFFFF) {
                    return false;
                }
                if (!projectionMap.TryGetMasterId(slide, out uint masterIdRef)) return false;
                persistId++;
                slideId++;
                drawingId++;
                IReadOnlyList<PowerPointShape> writableShapes =
                    LegacyPptWriter.ReadSlideShapesForWrite(slide, out _);
                SlideLayoutPart? layoutPart = slide.SlidePart.SlideLayoutPart;
                bool layoutIsIndependentMaster = layoutPart != null
                    && projectionMap.TryGetTitleMaster(layoutPart, out _);
                uint nextShapeIndex = checked(unchecked((uint)
                    LegacyPptWriter.CountDrawingShapes(writableShapes)) + 2U);
                clusters.Add(new KeyValuePair<uint, uint>(drawingId, nextShapeIndex));
                rewritten.Add(persistId,
                    LegacyPptWriter.BuildIncrementalSlideRecord(slide, drawingId,
                        masterIdRef, remappedInteractions, fonts,
                        pictureBullets,
                        layoutIsIndependentMaster));
                appendedSlideAtoms.Add(BuildSlidePersistAtom(persistId, slideId));
                slideIds.Add(slideId);
            }

            byte[] rewrittenDgg = BuildDggWithAppendedClusters(dgg, clusters, addedSlides.Count);
            var slideListChildren = slideList.Children.Select(child => child.CopyRecordBytes()).ToList();
            slideListChildren.AddRange(appendedSlideAtoms);
            byte[] rewrittenSlideList = BuildRecord(slideList.Version, slideList.Instance, slideList.Type,
                Concat(slideListChildren));

            bool patchedDgg = false;
            var documentChildren = new List<byte[]>(document.Children.Count);
            foreach (LegacyPptRecord child in document.Children) {
                if (ReferenceEquals(child, slideList)) {
                    documentChildren.Add(rewrittenSlideList);
                } else if (!patchedDgg && child.DescendantsAndSelf().Any(record => ReferenceEquals(record, dgg))) {
                    if (!TryReplaceDescendant(child, dgg.Offset, rewrittenDgg, out byte[] rewrittenChild)) {
                        return false;
                    }
                    documentChildren.Add(rewrittenChild);
                    patchedDgg = true;
                } else {
                    documentChildren.Add(child.CopyRecordBytes());
                }
            }
            if (!patchedDgg) return false;
            rewritten.Add(package.DocumentPersistId, BuildRecord(document.Version, document.Instance,
                document.Type, Concat(documentChildren)));
            return true;
        }

        private static bool TryReadDocument(LegacyPptPackage package, out LegacyPptRecord? document) {
            document = null;
            if (!package.PersistObjects.TryGetValue(package.DocumentPersistId,
                    out LegacyPptPersistObject? documentObject) || documentObject == null) {
                return false;
            }
            document = LegacyPptRecordReader.ReadSingle(documentObject.RecordBytes, 0,
                new LegacyPptImportOptions());
            return true;
        }

        private static bool TryReadDggClusters(LegacyPptRecord dgg,
            out List<KeyValuePair<uint, uint>> clusters) {
            clusters = new List<KeyValuePair<uint, uint>>();
            if (dgg.PayloadLength < 16) return false;
            uint storedClusterCount = dgg.ReadUInt32(4);
            if (storedClusterCount == 0 || storedClusterCount - 1U > int.MaxValue) return false;
            int clusterCount = unchecked((int)(storedClusterCount - 1U));
            if (dgg.PayloadLength < checked(16 + clusterCount * 8)) return false;
            for (int index = 0; index < clusterCount; index++) {
                clusters.Add(new KeyValuePair<uint, uint>(dgg.ReadUInt32(16 + index * 8),
                    dgg.ReadUInt32(20 + index * 8)));
            }
            return true;
        }

        private static byte[] BuildDggWithAppendedClusters(LegacyPptRecord dgg,
            IReadOnlyList<KeyValuePair<uint, uint>> clusters, int addedDrawingCount) {
            var payload = new byte[checked(16 + clusters.Count * 8)];
            uint lastDrawingId = clusters.Count == 0 ? 0U : clusters[clusters.Count - 1].Key;
            uint lastNextShapeIndex = clusters.Count == 0 ? 1U : clusters[clusters.Count - 1].Value;
            WriteUInt32(payload, 0, checked((lastDrawingId << 10) + lastNextShapeIndex));
            WriteUInt32(payload, 4, checked(unchecked((uint)clusters.Count) + 1U));
            uint addedShapeCount = unchecked((uint)clusters.Skip(clusters.Count - addedDrawingCount)
                .Sum(cluster => checked((int)cluster.Value - 1)));
            WriteUInt32(payload, 8, checked(dgg.ReadUInt32(8) + addedShapeCount));
            WriteUInt32(payload, 12, checked(dgg.ReadUInt32(12) + unchecked((uint)addedDrawingCount)));
            for (int index = 0; index < clusters.Count; index++) {
                WriteUInt32(payload, 16 + index * 8, clusters[index].Key);
                WriteUInt32(payload, 20 + index * 8, clusters[index].Value);
            }
            return BuildRecord(dgg.Version, dgg.Instance, dgg.Type, payload);
        }

        private static byte[] BuildSlidePersistAtom(uint persistId, uint slideId) {
            var payload = new byte[20];
            WriteUInt32(payload, 0, persistId);
            WriteUInt32(payload, 4, 4);
            WriteUInt32(payload, 12, slideId);
            return BuildRecord(version: 0, instance: 0, RecordSlidePersistAtom, payload);
        }

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

    }
}
