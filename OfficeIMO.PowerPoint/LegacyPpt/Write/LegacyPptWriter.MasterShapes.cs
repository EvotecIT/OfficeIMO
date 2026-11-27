using OfficeIMO.OpenXml.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using OfficeIMO.PowerPoint.OpenXml;

namespace OfficeIMO.PowerPoint.LegacyPpt.Write {
    internal static partial class LegacyPptWriter {
        /// <summary>
        /// The shapes of a slide master, notes master, handout master or title master layout, wrapped over the live
        /// document's part of the same name.
        /// </summary>
        internal static IReadOnlyList<PowerPointShape> ReadMasterShapesForWrite(
            PowerPointDocument presentation, LegacyPptPartView masterPart,
            out string? unsupportedReason) {
            ArgumentNullException.ThrowIfNull(presentation);
            if (masterPart.IsNone) throw new ArgumentNullException(nameof(masterPart));
            string ownerName = masterPart.IsNotesMaster ? "notes master"
                : masterPart.IsHandoutMaster ? "handout master"
                : masterPart.IsSlideLayout ? "title master layout"
                : "slide master";
            unsupportedReason = null;
            EnsureShapeWrappersCanBind(presentation);
            PowerPointPartRef ownerPart = PowerPointPartRef.Package(presentation, masterPart.Name);
            OfficeXmlNode tree = ownerPart.Root.Path(Pml, "cSld", "spTree");
            if (tree.IsNone) return Array.Empty<PowerPointShape>();
            var shapes = new List<PowerPointShape>(tree.Children().Count);
            foreach (OfficeXmlNode element in tree.Children()) {
                if (IsGroupProperties(element)) continue;
                PowerPointShape? shape = WrapShapeForWrite(presentation,
                    ownerPart, element);
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
