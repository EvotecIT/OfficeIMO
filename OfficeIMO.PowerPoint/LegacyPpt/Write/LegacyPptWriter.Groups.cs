using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using static OfficeIMO.PowerPoint.LegacyPpt.Write.LegacyPptWriterXml;

namespace OfficeIMO.PowerPoint.LegacyPpt.Write {
    internal static partial class LegacyPptWriter {
        private const ushort OfficeArtFspgr = 0xF009;

        internal static bool TryReadGroupForWrite(
            PowerPointGroupShape group,
            out IReadOnlyList<PowerPointShape> children,
            out string? reason) {
            ArgumentNullException.ThrowIfNull(group);
            children = ReadGroupChildrenForWrite(group, out reason);
            if (reason != null) return false;
            OfficeXmlNode groupElement = ShapeXml(group);
            OfficeXmlNode properties = groupElement.Child(LegacyPptXml.P,
                "grpSpPr");
            OfficeXmlNode transform = properties.Child(LegacyPptXml.A, "xfrm");
            if (properties.IsNone || transform.IsNone
                || HasAttributes(properties)
                || AnyChildOtherThan(properties, LegacyPptXml.A, "xfrm")
                || HasAttributeOtherThan(transform, "rot", "flipH", "flipV")
                || HasGroupTransformChildOtherThanPoints(transform)
                || transform.Child(LegacyPptXml.A, "off").Long("x") is not long left
                || transform.Child(LegacyPptXml.A, "off").Long("y") is not long top
                || transform.Child(LegacyPptXml.A, "ext").Long("cx") is not long width
                || transform.Child(LegacyPptXml.A, "ext").Long("cy") is not long height
                || transform.Child(LegacyPptXml.A, "chOff").Long("x") is not long childLeft
                || transform.Child(LegacyPptXml.A, "chOff").Long("y") is not long childTop
                || transform.Child(LegacyPptXml.A, "chExt").Long("cx") is not long childWidth
                || transform.Child(LegacyPptXml.A, "chExt").Long("cy") is not long childHeight
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
            OfficeXmlNode drawing = groupElement
                .Child(LegacyPptXml.P, "nvGrpSpPr")
                .Child(LegacyPptXml.P, "cNvGrpSpPr");
            if (!drawing.IsNone
                && (HasAttributes(drawing) || HasChildren(drawing))) {
                reason = "Group-shape locks and non-visual extensions are not yet encoded by the binary PowerPoint writer.";
                return false;
            }
            reason = null;
            return true;
        }

        private static bool HasGroupTransformChildOtherThanPoints(
            OfficeXmlNode transform) {
            foreach (OfficeXmlNode child in transform.Children()) {
                if (!child.Is(LegacyPptXml.A, "off")
                    && !child.Is(LegacyPptXml.A, "ext")
                    && !child.Is(LegacyPptXml.A, "chOff")
                    && !child.Is(LegacyPptXml.A, "chExt")) return true;
            }
            return false;
        }

        internal static int CountDrawingShapes(
            IEnumerable<PowerPointShape> shapes,
            LegacyPptWriterPictureCatalog? pictureCatalog = null) =>
            CountDrawingShapes(shapes, pictureCatalog, depth: 0);

        private static int CountDrawingShapes(
            IEnumerable<PowerPointShape> shapes,
            LegacyPptWriterPictureCatalog? pictureCatalog, int depth) {
            int count = 0;
            foreach (PowerPointShape shape in shapes) {
                if (shape is PowerPointGroupShape group) {
                    if (depth >= MaximumGroupNestingDepth) {
                        throw new NotSupportedException(
                            $"Binary PowerPoint conversion supports at most {MaximumGroupNestingDepth} nested group levels.");
                    }
                    if (!TryReadGroupForWrite(group,
                            out IReadOnlyList<PowerPointShape> children,
                            out string? reason)) {
                        throw new NotSupportedException(reason);
                    }
                    count = checked(count + 1
                        + CountDrawingShapes(children, pictureCatalog,
                            checked(depth + 1)));
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
            LegacyPptWriterPictureBulletCatalog? pictureBullets,
            int depth = 1) {
            if (depth > MaximumGroupNestingDepth) {
                throw new NotSupportedException(
                    $"Binary PowerPoint conversion supports at most {MaximumGroupNestingDepth} nested group levels.");
            }
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
                        pictureBullets, checked(depth + 1))
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
            ArgumentNullException.ThrowIfNull(prototype);
            ArgumentNullException.ThrowIfNull(group);
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
            OfficeXmlNode transform = ShapeXml(group)
                .Child(LegacyPptXml.P, "grpSpPr")
                .Child(LegacyPptXml.A, "xfrm");
            OfficeXmlNode childOffset = transform.Child(LegacyPptXml.A, "chOff");
            OfficeXmlNode childExtents = transform.Child(LegacyPptXml.A, "chExt");
            left = ToMasterUnits(childOffset.Long("x")
                ?? throw new InvalidOperationException("The group transform has no child offset."));
            top = ToMasterUnits(childOffset.Long("y")
                ?? throw new InvalidOperationException("The group transform has no child offset."));
            right = checked(left
                + ToMasterUnits(childExtents.Long("cx")
                    ?? throw new InvalidOperationException("The group transform has no child extents.")));
            bottom = checked(top
                + ToMasterUnits(childExtents.Long("cy")
                    ?? throw new InvalidOperationException("The group transform has no child extents.")));
        }
    }
}
