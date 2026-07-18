using System.Collections.ObjectModel;
using System.Runtime.CompilerServices;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing.Internal;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint.LegacyPpt.Write {
    internal static partial class LegacyPptWriter {
        private const ushort RecordExternalOleObjectAtom = 0x0FC3;
        private const ushort RecordExternalOleEmbed = 0x0FCC;
        private const ushort RecordExternalOleEmbedAtom = 0x0FCD;
        private const ushort RecordExternalObjectRefAtom = 0x0BC1;

        internal static bool TryReadOleObjects(
            IEnumerable<PowerPointSlide> slides, uint firstObjectId,
            out LegacyPptWriterOleObjectCatalog catalog,
            out string? reason) {
            catalog = new LegacyPptWriterOleObjectCatalog();
            reason = null;
            uint nextId = firstObjectId;
            foreach (PowerPointSlide slide in slides) {
                foreach (PowerPointOleObject shape in slide
                             .EnumerateShapesDeep(slide.Shapes,
                                 includeHidden: true)
                             .OfType<PowerPointOleObject>()) {
                    if (!TryReadOleObject(slide.SlidePart, shape, nextId,
                            out LegacyPptWriterOleObject? ole, out reason)
                        || ole == null) {
                        catalog = new LegacyPptWriterOleObjectCatalog();
                        return false;
                    }
                    catalog.Add(shape.Element, ole);
                    nextId = checked(nextId + 1U);
                }
            }
            return true;
        }

        private static bool TryReadOleObject(SlidePart slidePart,
            PowerPointOleObject shape, uint id,
            out LegacyPptWriterOleObject? result, out string? reason) {
            result = null;
            reason = null;
            if (id == 0) {
                reason = "OLE object identifiers cannot use reserved value zero.";
                return false;
            }
            if (shape.Element is not P.GraphicFrame frame) {
                reason = "The OLE object is not carried by a graphic frame.";
                return false;
            }
            P.OleObject? source = frame.Graphic?.GraphicData?
                .GetFirstChild<P.OleObject>();
            if (source == null
                || source.Elements<P.OleObjectEmbed>().Count() != 1
                || source.Elements<P.OleObjectLink>().Any()) {
                reason = "Linked or malformed Open XML OLE objects are not encoded as embedded binary objects.";
                return false;
            }
            P.OleObjectEmbed embed = source
                .GetFirstChild<P.OleObjectEmbed>()!;
            if (embed.HasChildren) {
                reason = "OLE embed extensions have no binary PowerPoint mapping.";
                return false;
            }
            string? relationshipId = source.Id?.Value;
            if (string.IsNullOrWhiteSpace(relationshipId)) {
                reason = "The OLE object has no embedded-part relationship.";
                return false;
            }
            if (slidePart.GetPartById(relationshipId!) is not
                    EmbeddedObjectPart part) {
                reason = "The OLE relationship does not target an EmbeddedObjectPart.";
                return false;
            }
            if (part.Parts.Any() || part.ExternalRelationships.Any()
                || part.HyperlinkRelationships.Any()) {
                reason = "Related parts on an embedded OLE storage have no binary PowerPoint mapping.";
                return false;
            }
            if (!TryReadOlePreview(slidePart, shape,
                    out PowerPointPicture? preview, out _, out _,
                    out reason)) {
                return false;
            }
            byte[] storageBytes;
            try {
                using Stream stream = part.GetStream(FileMode.Open,
                    FileAccess.Read);
                storageBytes = OfficeStreamReader.ReadAllBytes(stream,
                    64 * 1024 * 1024);
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
            string progId = source.ProgId?.Value ?? "Package";
            if (!IsValidOleString(progId)) {
                reason = "The OLE ProgID contains a null character.";
                return false;
            }
            string? name = source.Name?.Value;
            if (!IsValidOleString(name)) {
                reason = "The OLE display name contains a null character.";
                return false;
            }
            P.OleObjectFollowColorSchemeValues? follow =
                embed.FollowColorScheme?.Value;
            uint colorFollow = follow ==
                    P.OleObjectFollowColorSchemeValues.Full
                ? 1U
                : follow == P.OleObjectFollowColorSchemeValues
                    .TextAndBackground ? 2U : 0U;
            result = new LegacyPptWriterOleObject(id,
                source.ShowAsIcon?.Value == true ? 4U : 1U,
                MapOleSubType(progId), colorFollow, name, progId,
                storageBytes, preview);
            return true;
        }

        private static bool TryReadOlePreview(SlidePart slidePart,
            PowerPointOleObject shape, out PowerPointPicture? preview,
            out byte[] imageBytes, out string? contentType,
            out string? reason) {
            preview = null;
            imageBytes = Array.Empty<byte>();
            contentType = null;
            reason = null;
            if (shape.Element is not P.GraphicFrame frame) return true;
            P.Picture? source = frame.Graphic?.GraphicData?
                .GetFirstChild<P.OleObject>()?.GetFirstChild<P.Picture>();
            A.Blip? blip = source?.BlipFill?.Blip;
            if (source == null || blip == null
                || string.IsNullOrWhiteSpace(blip.Embed?.Value)
                    && string.IsNullOrWhiteSpace(blip.Link?.Value)) {
                return true;
            }

            preview = new PowerPointPicture(source, slidePart);
            if (preview.Left != shape.Left || preview.Top != shape.Top
                || preview.Width != shape.Width || preview.Height != shape.Height
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
                || (!string.IsNullOrEmpty(preview.Description)
                    && !string.Equals(preview.Description,
                        shape.Description, StringComparison.Ordinal))
                || !TryReadShapeMetadataForWrite(preview, out _, out reason)) {
                preview = null;
                reason ??= "The OLE preview carries independent accessibility metadata that one binary OfficeArt object shape cannot preserve.";
                return false;
            }
            if (!TryReadShapeVisualStyle(preview, out _, out reason)
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
            private readonly Dictionary<OpenXmlElement,
                LegacyPptWriterOleObject> _objectsByShape =
                new(ReferenceComparer.Instance);
            private readonly List<LegacyPptWriterOleObject> _objects = new();

            internal IReadOnlyList<LegacyPptWriterOleObject> Objects =>
                new ReadOnlyCollection<LegacyPptWriterOleObject>(_objects);

            internal LegacyPptWriterOleObject? Get(PowerPointShape shape) =>
                _objectsByShape.TryGetValue(shape.Element,
                    out LegacyPptWriterOleObject? value) ? value : null;

            internal uint AssignPersistIds(uint firstPersistId) {
                uint next = firstPersistId;
                foreach (LegacyPptWriterOleObject ole in _objects) {
                    ole.PersistId = next;
                    next = checked(next + 1U);
                }
                return next;
            }

            internal void Add(OpenXmlElement shape,
                LegacyPptWriterOleObject ole) {
                _objectsByShape.Add(shape, ole);
                _objects.Add(ole);
            }
        }

        internal sealed class LegacyPptWriterOleObject {
            internal LegacyPptWriterOleObject(uint id, uint drawAspect,
                uint subType, uint colorFollow, string? name,
                string progId, byte[] storageBytes,
                PowerPointPicture? preview) {
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
            internal PowerPointPicture? Preview { get; }
        }
    }
}
