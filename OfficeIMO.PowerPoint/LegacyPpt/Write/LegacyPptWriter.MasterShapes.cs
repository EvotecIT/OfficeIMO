using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint.LegacyPpt.Write {
    internal static partial class LegacyPptWriter {
        internal static IReadOnlyList<PowerPointShape> ReadMasterShapesForWrite(
            SlideMasterPart masterPart, out string? unsupportedReason) {
            if (masterPart == null) throw new ArgumentNullException(nameof(masterPart));
            return ReadMasterShapesForWrite(masterPart,
                masterPart.SlideMaster?.CommonSlideData?.ShapeTree,
                "slide master", out unsupportedReason);
        }

        internal static IReadOnlyList<PowerPointShape> ReadMasterShapesForWrite(
            NotesMasterPart masterPart, out string? unsupportedReason) {
            if (masterPart == null) throw new ArgumentNullException(nameof(masterPart));
            return ReadMasterShapesForWrite(masterPart,
                masterPart.NotesMaster?.CommonSlideData?.ShapeTree,
                "notes master", out unsupportedReason);
        }

        internal static IReadOnlyList<PowerPointShape> ReadMasterShapesForWrite(
            HandoutMasterPart masterPart, out string? unsupportedReason) {
            if (masterPart == null) throw new ArgumentNullException(nameof(masterPart));
            return ReadMasterShapesForWrite(masterPart,
                masterPart.HandoutMaster?.CommonSlideData?.ShapeTree,
                "handout master", out unsupportedReason);
        }

        internal static IReadOnlyList<PowerPointShape> ReadMasterShapesForWrite(
            SlideLayoutPart masterPart, out string? unsupportedReason) {
            if (masterPart == null) throw new ArgumentNullException(nameof(masterPart));
            return ReadMasterShapesForWrite(masterPart,
                masterPart.SlideLayout?.CommonSlideData?.ShapeTree,
                "title master layout", out unsupportedReason);
        }

        private static IReadOnlyList<PowerPointShape> ReadMasterShapesForWrite(
            OpenXmlPartContainer ownerPart, P.ShapeTree? tree, string ownerName,
            out string? unsupportedReason) {
            unsupportedReason = null;
            if (tree == null) return Array.Empty<PowerPointShape>();
            var shapes = new List<PowerPointShape>(tree.ChildElements.Count);
            foreach (OpenXmlElement element in tree.ChildElements) {
                if (element is P.NonVisualGroupShapeProperties
                    or P.GroupShapeProperties) continue;
                PowerPointShape? shape = WrapShapeForWrite(element,
                    ownerPart);
                if (shape == null) {
                    unsupportedReason ??=
                        $"The {ownerName} contains '{element.LocalName}' content that is not yet encoded by the native binary writer.";
                    continue;
                }
                shapes.Add(shape);
            }
            return shapes;
        }

        private static byte[] RewriteDrawingId(LegacyPptRecord record,
            uint drawingId) {
            if (record.Version == 0x0F && record.Children.Count > 0) {
                return BuildRecord(record.Version, record.Instance, record.Type,
                    Concat(record.Children.Select(child =>
                        RewriteDrawingId(child, drawingId))));
            }

            byte[] bytes = record.CopyRecordBytes();
            if (record.Type == OfficeArtDg) {
                WriteUInt16(bytes, 0,
                    checked((ushort)((drawingId << 4) | record.Version)));
                uint currentShapeId = ReadUInt32(bytes, 12);
                WriteUInt32(bytes, 12, checked((drawingId << 10)
                    | (currentShapeId & 0x000003FFU)));
            } else if (record.Type == OfficeArtFsp) {
                uint shapeId = ReadUInt32(bytes, 8);
                WriteUInt32(bytes, 8, checked((drawingId << 10)
                    | (shapeId & 0x000003FFU)));
            }
            return bytes;
        }
    }
}
