using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint.LegacyPpt.Write {
    internal static partial class LegacyPptPreservingWriter {
        private const ushort RecordDrawingForLayout = 0x040C;
        private const ushort OfficeArtDgContainerForLayout = 0xF002;
        private const ushort OfficeArtSpgrContainerForLayout = 0xF003;
        private const ushort OfficeArtDgForLayout = 0xF008;

        private static bool TryAppendMaterializedLayoutShapes(
            LegacyPptRecord slideRecord,
            IReadOnlyList<PowerPointShape> layoutShapes,
            LegacyPptWriter.LegacyPptWriterFontCatalog fonts,
            LegacyPptWriter.LegacyPptWriterPictureBulletCatalog
                pictureBullets,
            out byte[] bytes, out uint drawingId,
            out MaterializedLayoutDrawingUpdate update) {
            bytes = slideRecord.CopyRecordBytes();
            drawingId = 0;
            update = default;
            if (layoutShapes.Count == 0) return true;
            if (layoutShapes.Any(shape => !LegacyPptWriter
                    .IsLayoutShape(shape)
                || HasLayoutInteraction(shape))) return false;

            LegacyPptRecord[] drawings = slideRecord.Children.Where(child =>
                child.Type == RecordDrawingForLayout).ToArray();
            if (drawings.Length != 1) return false;
            LegacyPptRecord drawing = drawings[0];
            LegacyPptRecord[] drawingContainers = drawing.Children.Where(
                child => child.Type == OfficeArtDgContainerForLayout)
                .ToArray();
            if (drawingContainers.Length != 1) return false;
            LegacyPptRecord drawingContainer = drawingContainers[0];
            LegacyPptRecord[] drawingAtoms = drawingContainer.Children.Where(
                child => child.Type == OfficeArtDgForLayout).ToArray();
            LegacyPptRecord[] shapeGroups = drawingContainer.Children.Where(
                child => child.Type == OfficeArtSpgrContainerForLayout)
                .ToArray();
            if (drawingAtoms.Length != 1 || shapeGroups.Length != 1
                || drawingAtoms[0].PayloadLength < 8) return false;
            LegacyPptRecord drawingAtom = drawingAtoms[0];
            LegacyPptRecord shapeGroup = shapeGroups[0];
            drawingId = drawingAtom.Instance;
            if (drawingId == 0 || drawingId >= 0x00400000) return false;

            uint shapeIdBase = checked(drawingId << 10);
            uint shapeIdLimit = checked(shapeIdBase + 1024U);
            uint[] shapeIds = drawingContainer.DescendantsAndSelf()
                .Where(record => record.Type == OfficeArtFsp
                    && record.PayloadLength >= 4)
                .Select(record => record.ReadUInt32(0))
                .Where(id => id >= shapeIdBase && id < shapeIdLimit)
                .ToArray();
            if (shapeIds.Length == 0) return false;
            uint nextShapeId = checked(shapeIds.Max() + 1U);
            int addedShapeCount = LegacyPptWriter.CountDrawingShapes(
                layoutShapes);
            if (addedShapeCount <= 0
                || checked(nextShapeId
                    + unchecked((uint)addedShapeCount)) > shapeIdLimit) {
                return false;
            }

            IReadOnlyList<byte[]> appended = LegacyPptWriter
                .BuildAppendedSlideShapeRecords(layoutShapes,
                    ref nextShapeId, fonts, pictureBullets);
            if (appended.Count != layoutShapes.Count) return false;
            update = new MaterializedLayoutDrawingUpdate(
                checked(nextShapeId - shapeIdBase),
                unchecked((uint)addedShapeCount));

            byte[] drawingAtomBytes = drawingAtom.CopyRecordBytes();
            WriteUInt32(drawingAtomBytes, 8, checked(
                drawingAtom.ReadUInt32(0)
                + unchecked((uint)addedShapeCount)));
            WriteUInt32(drawingAtomBytes, 12,
                checked(nextShapeId - 1U));

            var shapeGroupChildren = new List<byte[]>(
                checked(shapeGroup.Children.Count + appended.Count));
            bool inserted = false;
            foreach (LegacyPptRecord child in shapeGroup.Children) {
                shapeGroupChildren.Add(child.CopyRecordBytes());
                if (!inserted && child.Type == OfficeArtSpContainer) {
                    shapeGroupChildren.AddRange(appended);
                    inserted = true;
                }
            }
            if (!inserted) return false;
            byte[] shapeGroupBytes = BuildRecord(shapeGroup.Version,
                shapeGroup.Instance, shapeGroup.Type,
                Concat(shapeGroupChildren));

            var drawingContainerChildren = new List<byte[]>(
                drawingContainer.Children.Count);
            foreach (LegacyPptRecord child in drawingContainer.Children) {
                drawingContainerChildren.Add(ReferenceEquals(child,
                        drawingAtom)
                    ? drawingAtomBytes
                    : ReferenceEquals(child, shapeGroup)
                        ? shapeGroupBytes
                        : child.CopyRecordBytes());
            }
            byte[] drawingContainerBytes = BuildRecord(
                drawingContainer.Version, drawingContainer.Instance,
                drawingContainer.Type, Concat(drawingContainerChildren));

            var drawingChildren = new List<byte[]>(drawing.Children.Count);
            foreach (LegacyPptRecord child in drawing.Children) {
                drawingChildren.Add(ReferenceEquals(child, drawingContainer)
                    ? drawingContainerBytes
                    : child.CopyRecordBytes());
            }
            byte[] drawingBytes = BuildRecord(drawing.Version,
                drawing.Instance, drawing.Type, Concat(drawingChildren));

            var slideChildren = new List<byte[]>(slideRecord.Children.Count);
            foreach (LegacyPptRecord child in slideRecord.Children) {
                slideChildren.Add(ReferenceEquals(child, drawing)
                    ? drawingBytes
                    : child.CopyRecordBytes());
            }
            bytes = BuildRecord(slideRecord.Version, slideRecord.Instance,
                slideRecord.Type, Concat(slideChildren));
            return true;
        }

        private static bool HasLayoutInteraction(PowerPointShape shape) =>
            shape.Element.Descendants<A.HyperlinkOnClick>().Any()
            || shape.Element.Descendants<A.HyperlinkOnHover>().Any()
            || shape.Element.Descendants<A.HyperlinkOnMouseOver>().Any();

        private static bool TryRewriteDocumentDrawingClusters(
            LegacyPptPackage package, byte[]? currentDocumentBytes,
            IReadOnlyDictionary<uint, MaterializedLayoutDrawingUpdate>
                updates,
            out byte[] bytes) {
            bytes = currentDocumentBytes
                ?? package.PersistObjects[package.DocumentPersistId]
                    .RecordBytes;
            if (updates.Count == 0) return true;
            LegacyPptRecord document = LegacyPptRecordReader.ReadSingle(
                bytes, 0, new LegacyPptImportOptions());
            LegacyPptRecord[] drawingGroups = document.DescendantsAndSelf()
                .Where(record => record.Type == OfficeArtDgg).ToArray();
            if (drawingGroups.Length != 1
                || !TryReadDggClusters(drawingGroups[0],
                    out List<KeyValuePair<uint, uint>> clusters)) {
                return false;
            }
            LegacyPptRecord dgg = drawingGroups[0];
            byte[] rewrittenDgg = dgg.CopyRecordBytes();
            uint addedShapeCount = 0;
            foreach (KeyValuePair<uint, MaterializedLayoutDrawingUpdate>
                         update in updates) {
                int index = clusters.FindIndex(pair =>
                    pair.Key == update.Key);
                if (index < 0 || update.Value.NextShapeIndex
                    < clusters[index].Value) {
                    return false;
                }
                addedShapeCount = checked(addedShapeCount
                    + update.Value.AddedShapeCount);
                clusters[index] = new KeyValuePair<uint, uint>(
                    update.Key, update.Value.NextShapeIndex);
                WriteUInt32(rewrittenDgg, checked(28 + index * 8),
                    update.Value.NextShapeIndex);
            }
            WriteUInt32(rewrittenDgg, 8, clusters.Max(pair => checked(
                (pair.Key << 10) + pair.Value)));
            WriteUInt32(rewrittenDgg, 16, checked(
                dgg.ReadUInt32(8) + addedShapeCount));
            if (!TryReplaceDescendant(document, dgg.Offset,
                    rewrittenDgg, out byte[] rewrittenDocument)) {
                return false;
            }
            bytes = rewrittenDocument;
            return true;
        }

        private readonly struct MaterializedLayoutDrawingUpdate {
            internal MaterializedLayoutDrawingUpdate(uint nextShapeIndex,
                uint addedShapeCount) {
                NextShapeIndex = nextShapeIndex;
                AddedShapeCount = addedShapeCount;
            }

            internal uint NextShapeIndex { get; }

            internal uint AddedShapeCount { get; }
        }
    }
}
