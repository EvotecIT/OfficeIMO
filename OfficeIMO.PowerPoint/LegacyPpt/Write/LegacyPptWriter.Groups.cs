using DocumentFormat.OpenXml;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint.LegacyPpt.Write {
    internal static partial class LegacyPptWriter {
        private const ushort OfficeArtFspgr = 0xF009;

        internal static bool TryReadGroupForWrite(
            PowerPointGroupShape group,
            out IReadOnlyList<PowerPointShape> children,
            out string? reason) {
            if (group == null) throw new ArgumentNullException(nameof(group));
            children = ReadGroupChildrenForWrite(group, out reason);
            if (reason != null) return false;
            P.GroupShapeProperties? properties = group.GroupShape
                .GroupShapeProperties;
            A.TransformGroup? transform = properties?.TransformGroup;
            if (properties == null || transform == null
                || properties.HasAttributes
                || properties.ChildElements.Any(child =>
                    child is not A.TransformGroup)
                || transform.GetAttributes().Any(attribute =>
                    attribute.LocalName is not "rot"
                        and not "flipH" and not "flipV")
                || transform.ChildElements.Any(child =>
                    child is not A.Offset and not A.Extents
                        and not A.ChildOffset and not A.ChildExtents)
                || transform.Offset?.X?.Value is not long left
                || transform.Offset.Y?.Value is not long top
                || transform.Extents?.Cx?.Value is not long width
                || transform.Extents.Cy?.Value is not long height
                || transform.ChildOffset?.X?.Value is not long childLeft
                || transform.ChildOffset.Y?.Value is not long childTop
                || transform.ChildExtents?.Cx?.Value is not long childWidth
                || transform.ChildExtents.Cy?.Value is not long childHeight
                || width < 0L || height < 0L
                || childWidth <= 0L || childHeight <= 0L) {
                reason = "The group requires one complete offset, extent, child-offset, and child-extent transform with no unsupported group properties.";
                return false;
            }
            try {
                _ = checked(ToMasterUnits(left) + ToMasterUnits(width));
                _ = checked(ToMasterUnits(top) + ToMasterUnits(height));
                _ = checked(ToMasterUnits(childLeft)
                    + ToMasterUnits(childWidth));
                _ = checked(ToMasterUnits(childTop)
                    + ToMasterUnits(childHeight));
            } catch (OverflowException) {
                reason = "The group transform exceeds the classic OfficeArt coordinate range.";
                return false;
            }
            if (!TryReadShapeTransform(group, out _, out reason)) {
                return false;
            }
            P.NonVisualGroupShapeDrawingProperties? drawing = group
                .GroupShape.NonVisualGroupShapeProperties?
                .NonVisualGroupShapeDrawingProperties;
            if (drawing is { HasAttributes: true }
                || drawing is { HasChildren: true }) {
                reason = "Group-shape locks and non-visual extensions are not yet encoded by the binary PowerPoint writer.";
                return false;
            }
            reason = null;
            return true;
        }

        internal static int CountDrawingShapes(
            IEnumerable<PowerPointShape> shapes,
            LegacyPptWriterPictureCatalog? pictureCatalog = null) {
            int count = 0;
            foreach (PowerPointShape shape in shapes) {
                if (shape is PowerPointGroupShape group) {
                    if (!TryReadGroupForWrite(group,
                            out IReadOnlyList<PowerPointShape> children,
                            out string? reason)) {
                        throw new NotSupportedException(reason);
                    }
                    count = checked(count + 1
                        + CountDrawingShapes(children, pictureCatalog));
                } else if (shape is PowerPointTable table) {
                    count = checked(count + (pictureCatalog?.Contains(table)
                        == true ? 1 : CountTableDrawingShapes(table)));
                } else {
                    count = checked(count + 1);
                }
            }
            return count;
        }

        private static byte[] BuildGroupRecord(PowerPointGroupShape group,
            ref uint nextShapeId,
            LegacyPptWriterInteractionCatalog interactionCatalog,
            LegacyPptWriterAnimationCatalog animationCatalog,
            LegacyPptWriterShapeContext shapeContext,
            LegacyPptWriterMediaCatalog? mediaCatalog,
            LegacyPptWriterOleObjectCatalog? oleCatalog,
            LegacyPptWriterPictureCatalog? pictureCatalog,
            LegacyPptWriterFontCatalog fonts,
            LegacyPptWriterPictureBulletCatalog? pictureBullets) {
            if (!TryReadGroupForWrite(group,
                    out IReadOnlyList<PowerPointShape> groupChildren,
                    out string? reason)) {
                throw new NotSupportedException(reason);
            }
            uint groupShapeId = nextShapeId++;
            var descriptorChildren = new List<byte[]> {
                BuildGroupCoordinateRecord(group),
                BuildFsp(0, groupShapeId, group, isGroup: true)
            };
            byte[]? formatting = BuildShapeFoptRecord(group);
            if (formatting != null) descriptorChildren.Add(formatting);
            descriptorChildren.Add(BuildAnchor(group));
            byte[]? clientData = BuildClientData(group,
                interactionCatalog.Get(group).ShapeInteractions,
                animationCatalog.Get(group), shapeContext);
            if (clientData != null) descriptorChildren.Add(clientData);

            var children = new List<byte[]>(groupChildren.Count + 1) {
                BuildContainer(OfficeArtSpContainer, instance: 0,
                    descriptorChildren)
            };
            foreach (PowerPointShape child in groupChildren) {
                children.Add(child is PowerPointGroupShape nested
                    ? BuildGroupRecord(nested, ref nextShapeId,
                        interactionCatalog, animationCatalog, shapeContext,
                        mediaCatalog, oleCatalog, pictureCatalog, fonts,
                        pictureBullets)
                    : child is PowerPointTable table
                        && pictureCatalog?.Contains(table) != true
                    ? BuildTableRecord(table, ref nextShapeId,
                        interactionCatalog, animationCatalog, shapeContext,
                        fonts, pictureBullets)
                    : BuildShapeRecord(child, nextShapeId++,
                        interactionCatalog, animationCatalog, shapeContext,
                        mediaCatalog, oleCatalog, pictureCatalog, fonts,
                        pictureBullets));
            }
            return BuildContainer(OfficeArtSpgrContainer, instance: 0,
                children);
        }

        private static byte[] BuildGroupCoordinateRecord(
            PowerPointGroupShape group) {
            GetGroupCoordinateValues(group, out int left, out int top,
                out int right, out int bottom);
            var payload = new byte[16];
            WriteInt32(payload, 0, left);
            WriteInt32(payload, 4, top);
            WriteInt32(payload, 8, right);
            WriteInt32(payload, 12, bottom);
            return BuildRecord(version: 1, instance: 0,
                OfficeArtFspgr, payload);
        }

        internal static byte[] BuildPreservedGroupCoordinateRecord(
            LegacyPptRecord prototype, PowerPointGroupShape group) {
            if (prototype == null) throw new ArgumentNullException(
                nameof(prototype));
            if (group == null) throw new ArgumentNullException(nameof(group));
            if (prototype.Type != OfficeArtFspgr
                || prototype.PayloadLength < 16) {
                throw new InvalidDataException(
                    "The preserved OfficeArt FSPGR atom is truncated.");
            }
            GetGroupCoordinateValues(group, out int left, out int top,
                out int right, out int bottom);
            byte[] bytes = prototype.CopyRecordBytes();
            WriteInt32(bytes, 8, left);
            WriteInt32(bytes, 12, top);
            WriteInt32(bytes, 16, right);
            WriteInt32(bytes, 20, bottom);
            return bytes;
        }

        private static void GetGroupCoordinateValues(
            PowerPointGroupShape group, out int left, out int top,
            out int right, out int bottom) {
            A.TransformGroup transform = group.GroupShape
                .GroupShapeProperties!.TransformGroup!;
            left = ToMasterUnits(transform.ChildOffset!.X!.Value);
            top = ToMasterUnits(transform.ChildOffset.Y!.Value);
            right = checked(left
                + ToMasterUnits(transform.ChildExtents!.Cx!.Value));
            bottom = checked(top
                + ToMasterUnits(transform.ChildExtents.Cy!.Value));
        }
    }
}
