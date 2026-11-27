using System.Collections.ObjectModel;
using OfficeIMO.Core.Internal;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using static OfficeIMO.PowerPoint.LegacyPpt.Write.LegacyPptWriterXml;

namespace OfficeIMO.PowerPoint.LegacyPpt.Write {
    internal static partial class LegacyPptWriter {
        private const ushort RecordExternalOleObjectAtom = 0x0FC3;
        private const ushort RecordExternalOleEmbed = 0x0FCC;
        private const ushort RecordExternalOleEmbedAtom = 0x0FCD;
        private const ushort RecordExternalObjectRefAtom = 0x0BC1;

        internal static bool TryReadOleObjects(LegacyPptPackageView view,
            IEnumerable<PowerPointSlide> slides, uint firstObjectId,
            out LegacyPptWriterOleObjectCatalog catalog,
            out string? reason) {
            ArgumentNullException.ThrowIfNull(view);
            catalog = new LegacyPptWriterOleObjectCatalog();
            reason = null;
            uint nextId = firstObjectId;
            foreach (PowerPointSlide slide in slides) {
                LegacyPptPartView slidePart = view.Part(slide.PartRef.Name);
                foreach (PowerPointOleObject shape in slide
                             .EnumerateShapesDeep(slide.Shapes,
                                 includeHidden: true)
                             .OfType<PowerPointOleObject>()) {
                    if (!TryReadOleObject(view, slidePart, shape, nextId,
                            out LegacyPptWriterOleObject? ole, out reason)
                        || ole == null) {
                        catalog = new LegacyPptWriterOleObjectCatalog();
                        return false;
                    }
                    catalog.Add(shape, ole);
                    nextId = checked(nextId + 1U);
                }
            }
            return true;
        }

        // The p:oleObj a graphic frame carries, or none.
        private static OfficeXmlNode OleObjectElement(OfficeXmlNode frame) =>
            frame.Is(LegacyPptXml.P, "graphicFrame")
                ? frame.Child(LegacyPptXml.A, "graphic")
                    .Child(LegacyPptXml.A, "graphicData")
                    .Child(LegacyPptXml.P, "oleObj")
                : default;

        private static bool TryReadOleObject(LegacyPptPackageView view,
            LegacyPptPartView slidePart,
            PowerPointOleObject shape, uint id,
            out LegacyPptWriterOleObject? result, out string? reason) {
            result = null;
            reason = null;
            if (id == 0) {
                reason = "OLE object identifiers cannot use reserved value zero.";
                return false;
            }
            OfficeXmlNode frame = view.Resolve(shape.OwnerPartRef.Name,
                shape.XmlElement);
            if (!frame.Is(LegacyPptXml.P, "graphicFrame")) {
                reason = "The OLE object is not carried by a graphic frame.";
                return false;
            }
            OfficeXmlNode source = OleObjectElement(frame);
            if (source.IsNone
                || source.Children(LegacyPptXml.P, "embed").Count != 1
                || source.Children(LegacyPptXml.P, "link").Count > 0) {
                reason = "Linked or malformed Open XML OLE objects are not encoded as embedded binary objects.";
                return false;
            }
            OfficeXmlNode embed = source.Child(LegacyPptXml.P, "embed");
            if (HasChildren(embed)) {
                reason = "OLE embed extensions have no binary PowerPoint mapping.";
                return false;
            }
            string? relationshipId = source.Attribute(LegacyPptXml.R, "id");
            if (string.IsNullOrWhiteSpace(relationshipId)) {
                reason = "The OLE object has no embedded-part relationship.";
                return false;
            }
            if (!slidePart.TryGetRelationship(relationshipId,
                    out LegacyPptRelationshipView relationship)
                || relationship.IsExternal
                || relationship.TargetPart.IsNone) {
                throw new ArgumentOutOfRangeException(nameof(relationshipId),
                    $"The relationship '{relationshipId}' does not target a part of '{slidePart.Name}'.");
            }
            LegacyPptPartView part = relationship.TargetPart;
            if (!relationship.IsType(LegacyPptPackageView.OleObjectRelationshipType)) {
                reason = "The OLE relationship does not target an EmbeddedObjectPart.";
                return false;
            }
            if (part.Relationships.Count > 0) {
                reason = "Related parts on an embedded OLE storage have no binary PowerPoint mapping.";
                return false;
            }
            if (!TryReadOlePreview(view, slidePart, shape,
                    out LegacyPptWriterOlePreview? preview, out _, out _,
                    out reason)) {
                return false;
            }
            byte[] storageBytes;
            try {
                storageBytes = part.ReadBytes(64 * 1024 * 1024);
            } catch (Exception exception) when (exception is IOException
                                                or InvalidDataException
                                                or UnauthorizedAccessException) {
                reason = $"The embedded OLE storage could not be read: {exception.Message}";
                return false;
            }
            if (!OfficeCompoundFileReader.TryRead(storageBytes,
                    out OfficeCompoundFile? compound, out reason)
                || compound == null) {
                return false;
            }
            string progId = source.Attribute("progId") ?? "Package";
            if (!IsValidOleString(progId)) {
                reason = "The OLE ProgID contains a null character.";
                return false;
            }
            string? name = source.Attribute("name");
            if (!IsValidOleString(name)) {
                reason = "The OLE display name contains a null character.";
                return false;
            }
            string? follow = embed.Attribute("followColorScheme");
            uint colorFollow = follow == "full"
                ? 1U
                : follow == "textAndBackground" ? 2U : 0U;
            result = new LegacyPptWriterOleObject(id,
                source.Bool("showAsIcon") == true ? 4U : 1U,
                MapOleSubType(progId), colorFollow, name, progId,
                storageBytes, preview);
            return true;
        }

        private static bool TryReadOlePreview(LegacyPptPackageView view,
            LegacyPptPartView slidePart,
            PowerPointOleObject shape, out LegacyPptWriterOlePreview? preview,
            out byte[] imageBytes, out string? contentType,
            out string? reason) {
            preview = null;
            imageBytes = Array.Empty<byte>();
            contentType = null;
            reason = null;
            OfficeXmlNode frame = view.Resolve(shape.OwnerPartRef.Name,
                shape.XmlElement);
            if (!frame.Is(LegacyPptXml.P, "graphicFrame")) return true;
            OfficeXmlNode source = OleObjectElement(frame)
                .Child(LegacyPptXml.P, "pic");
            OfficeXmlNode blip = source.Child(LegacyPptXml.P, "blipFill")
                .Child(LegacyPptXml.A, "blip");
            if (source.IsNone || blip.IsNone
                || string.IsNullOrWhiteSpace(blip.Attribute(LegacyPptXml.R, "embed"))
                    && string.IsNullOrWhiteSpace(blip.Attribute(LegacyPptXml.R, "link"))) {
                return true;
            }

            preview = new LegacyPptWriterOlePreview(source, slidePart);
            if (preview.Left != shape.Left.ToEmus() || preview.Top != shape.Top.ToEmus()
                || preview.Width != shape.Width.ToEmus() || preview.Height != shape.Height.ToEmus()
                || preview.Rotation != shape.Rotation
                || preview.HorizontalFlip != shape.HorizontalFlip
                || preview.VerticalFlip != shape.VerticalFlip) {
                preview = null;
                reason = "The OLE preview uses geometry that differs from its owning object frame and cannot be represented by one binary OfficeArt shape.";
                return false;
            }
            if ((!string.IsNullOrEmpty(preview.Name)
                    && !string.Equals(preview.Name, shape.Name,
                        StringComparison.Ordinal))
                || (!string.IsNullOrEmpty(preview.AltText)
                    && !string.Equals(preview.AltText,
                        shape.AltText, StringComparison.Ordinal))
                || !TryReadShapeMetadataForWrite(preview.Title,
                    preview.Decorative, preview.Name, preview.AltText,
                    out _, out reason)) {
                preview = null;
                reason ??= "The OLE preview carries independent accessibility metadata that one binary OfficeArt object shape cannot preserve.";
                return false;
            }
            if (!TryReadShapeVisualStyle(preview.Element, out _, out reason)
                || !TryReadPicture(preview, out imageBytes,
                    out contentType, out reason)) {
                preview = null;
                return false;
            }
            return true;
        }

        internal static byte[] BuildExternalOleObjectRecord(
            LegacyPptWriterOleObject ole) {
            if (ole.PersistId == 0) {
                throw new InvalidOperationException(
                    "The OLE object has no assigned persist identifier.");
            }
            var embedPayload = new byte[8];
            WriteUInt32(embedPayload, 0, ole.ColorFollow);
            var objectPayload = new byte[24];
            WriteUInt32(objectPayload, 0, ole.DrawAspect);
            WriteUInt32(objectPayload, 4, 0);
            WriteUInt32(objectPayload, 8, ole.Id);
            WriteUInt32(objectPayload, 12, ole.SubType);
            WriteUInt32(objectPayload, 16, ole.PersistId);
            var children = new List<byte[]> {
                BuildRecord(version: 0, instance: 0,
                    RecordExternalOleEmbedAtom, embedPayload),
                BuildRecord(version: 1, instance: 0,
                    RecordExternalOleObjectAtom, objectPayload)
            };
            if (!string.IsNullOrEmpty(ole.Name)) {
                children.Add(BuildOleString(1, ole.Name!));
            }
            children.Add(BuildOleString(2, ole.ProgId));
            if (!string.IsNullOrEmpty(ole.Name)) {
                children.Add(BuildOleString(3, ole.Name!));
            }
            return BuildContainer(RecordExternalOleEmbed, instance: 0,
                children);
        }

        internal static byte[] BuildOleObjectStorageRecord(
            LegacyPptWriterOleObject ole) => BuildRecord(version: 0,
                instance: 0, RecordExternalOleObjectStorage,
                ole.StorageBytes);

        internal static byte[] BuildOleObjectStorageRecord(
            byte[] storageBytes) => BuildRecord(version: 0,
                instance: 0, RecordExternalOleObjectStorage,
                storageBytes ?? throw new ArgumentNullException(
                    nameof(storageBytes)));

        internal static byte[] BuildExternalObjectReferenceAtom(uint id) {
            var payload = new byte[4];
            WriteUInt32(payload, 0, id);
            return BuildRecord(version: 0, instance: 0,
                RecordExternalObjectRefAtom, payload);
        }

        internal static byte[] BuildOleString(ushort instance,
            string value) => BuildRecord(version: 0, instance,
                RecordCString,
                System.Text.Encoding.Unicode.GetBytes(value));

        private static bool IsValidOleString(string? value) =>
            value == null || value.IndexOf('\0') < 0;

        internal static uint MapOleSubType(string progId) {
            if (progId.StartsWith("Word.Document",
                    StringComparison.OrdinalIgnoreCase)) return 2;
            if (progId.StartsWith("Excel.Chart",
                    StringComparison.OrdinalIgnoreCase)) return 14;
            if (progId.StartsWith("Excel.Sheet",
                    StringComparison.OrdinalIgnoreCase)) return 3;
            if (progId.StartsWith("MSGraph",
                    StringComparison.OrdinalIgnoreCase)) return 4;
            if (progId.StartsWith("Equation",
                    StringComparison.OrdinalIgnoreCase)) return 6;
            if (progId.StartsWith("MSProject",
                    StringComparison.OrdinalIgnoreCase)) return 12;
            if (progId.StartsWith("Visio.Drawing",
                    StringComparison.OrdinalIgnoreCase)) return 17;
            return 0;
        }

        internal sealed class LegacyPptWriterOleObjectCatalog {
            // Keyed by the OLE shape's write key.
            private readonly Dictionary<object,
                LegacyPptWriterOleObject> _objectsByShape = new();
            private readonly List<LegacyPptWriterOleObject> _objects = new();

            internal IReadOnlyList<LegacyPptWriterOleObject> Objects =>
                new ReadOnlyCollection<LegacyPptWriterOleObject>(_objects);

            internal LegacyPptWriterOleObject? Get(PowerPointShape shape) =>
                _objectsByShape.TryGetValue(ShapeKey(shape),
                    out LegacyPptWriterOleObject? value) ? value : null;

            internal uint AssignPersistIds(uint firstPersistId) {
                uint next = firstPersistId;
                foreach (LegacyPptWriterOleObject ole in _objects) {
                    ole.PersistId = next;
                    next = checked(next + 1U);
                }
                return next;
            }

            internal void Add(PowerPointShape shape,
                LegacyPptWriterOleObject ole) {
                _objectsByShape.Add(ShapeKey(shape), ole);
                _objects.Add(ole);
            }
        }

        /// <summary>
        /// The preview picture (<c>p:pic</c>) of an OLE object, read from the package view as the shape members read a
        /// picture: geometry from its transform, accessibility metadata from its non-visual properties, and its image
        /// from the slide part's relationships.
        /// </summary>
        internal sealed class LegacyPptWriterOlePreview {
            private const string SvgNamespace = "http://schemas.microsoft.com/office/drawing/2016/SVG/main";
            private const string DecorativeNamespace = "http://schemas.microsoft.com/office/drawing/2017/decorative";

            internal LegacyPptWriterOlePreview(OfficeXmlNode element,
                LegacyPptPartView ownerPart) {
                Element = element;
                OwnerPart = ownerPart;
            }

            internal OfficeXmlNode Element { get; }
            internal LegacyPptPartView OwnerPart { get; }

            private OfficeXmlNode Transform => Element
                .Child(LegacyPptXml.P, "spPr").Child(LegacyPptXml.A, "xfrm");

            private OfficeXmlNode Drawing => Element
                .Child(LegacyPptXml.P, "nvPicPr").Child(LegacyPptXml.P, "cNvPr");

            internal long Left => Transform.Child(LegacyPptXml.A, "off").Int64Attribute("x") ?? 0L;
            internal long Top => Transform.Child(LegacyPptXml.A, "off").Int64Attribute("y") ?? 0L;
            internal long Width => Transform.Child(LegacyPptXml.A, "ext").Int64Attribute("cx") ?? 0L;
            internal long Height => Transform.Child(LegacyPptXml.A, "ext").Int64Attribute("cy") ?? 0L;

            internal double? Rotation {
                get {
                    int? rotation = Transform.Int32Attribute("rot");
                    return rotation != null ? rotation.Value / 60000d : null;
                }
            }

            internal bool? HorizontalFlip => Transform.BooleanAttribute("flipH");
            internal bool? VerticalFlip => Transform.BooleanAttribute("flipV");
            internal string? Name => Drawing.Attribute("name");
            internal string? AltText => Drawing.Attribute("descr");
            internal string? Title => Drawing.Attribute("title");

            internal bool Decorative {
                get {
                    OfficeXmlNode drawing = Drawing;
                    if (drawing.IsNone) return false;
                    foreach (OfficeXmlNode element in drawing.EveryDescendant()) {
                        if (element.Is(DecorativeNamespace, "decorative")
                            && OfficeXmlViewValues.ParseBoolean(element.Attribute("val")) != false) {
                            return true;
                        }
                    }
                    return false;
                }
            }

            /// <summary>
            /// The image part: the SVG image for a picture with an SVG extension, the blip's embedded image otherwise.
            /// </summary>
            internal LegacyPptPartView ImagePart {
                get {
                    OfficeXmlNode blip = Element.Child(LegacyPptXml.P, "blipFill")
                        .Child(LegacyPptXml.A, "blip");
                    foreach (OfficeXmlNode vector in blip.EveryDescendant()) {
                        if (!vector.Is(SvgNamespace, "svgBlip")) continue;
                        string? svg = vector.Attribute(LegacyPptXml.R, "embed");
                        if (string.IsNullOrEmpty(svg)) continue;
                        LegacyPptPartView svgPart = ImagePartById(svg);
                        if (!svgPart.IsNone) return svgPart;
                        break;
                    }
                    return ImagePartById(blip.Attribute(LegacyPptXml.R, "embed"));
                }
            }

            private LegacyPptPartView ImagePartById(string? relationshipId) =>
                OwnerPart.TryGetRelationship(relationshipId,
                    out LegacyPptRelationshipView relationship)
                && !relationship.IsExternal
                && relationship.IsType(LegacyPptPackageView.ImageRelationshipType)
                    ? relationship.TargetPart
                    : default;
        }

        internal sealed class LegacyPptWriterOleObject {
            internal LegacyPptWriterOleObject(uint id, uint drawAspect,
                uint subType, uint colorFollow, string? name,
                string progId, byte[] storageBytes,
                LegacyPptWriterOlePreview? preview) {
                Id = id;
                DrawAspect = drawAspect;
                SubType = subType;
                ColorFollow = colorFollow;
                Name = name;
                ProgId = progId;
                StorageBytes = storageBytes;
                Preview = preview;
            }

            internal uint Id { get; }
            internal uint PersistId { get; set; }
            internal uint DrawAspect { get; }
            internal uint SubType { get; }
            internal uint ColorFollow { get; }
            internal string? Name { get; }
            internal string ProgId { get; }
            internal byte[] StorageBytes { get; }
            internal LegacyPptWriterOlePreview? Preview { get; }
        }
    }
}
