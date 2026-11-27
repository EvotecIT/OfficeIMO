using OfficeIMO.Drawing.Binary;
using OfficeIMO.Core.Internal;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Word.LegacyDoc.Model;
using OfficeIMO.Drawing;

namespace OfficeIMO.Word.LegacyDoc.Write {
    internal static partial class LegacyDocWriter {
        private const short LegacyPictureFormatShape = 0x0064;
        private const ushort OfficeArtPictureFrameShapeType = 75;
        private const ushort OfficeArtSpContainer = 0xF004;
        private const ushort OfficeArtFsp = 0xF00A;
        private const ushort OfficeArtFopt = 0xF00B;
        private const ushort OfficeArtClientAnchor = 0xF010;
        private const uint OfficeArtPictureShapeFlags = 0x00000A00;
        private const long EmusPerTwip = 635L;
        private const int MinimumPictureTwips = 15;
        private const int MaximumPictureTwips = 31680;

        private sealed class LegacyDocWritablePictures {
            private readonly MemoryStream _data = new();
            private uint _nextShapeId = 1;

            internal LegacyDocWritablePictures(WordDocument document, LegacyDocSourcePackage source) {
                // Pictures resolve their image parts through the owner part passed to AddInlinePicture, which carries
                // the package; the document supplies the imported compound Data stream.
                _ = source;
                bool preserveExistingData = document.LegacyDocCompoundFeatures.Any(feature =>
                    feature.Kind == LegacyDocCompoundFeatureKind.BinaryData);
                if (preserveExistingData
                    && document.LegacyDocSourceCompoundFile?.Streams.TryGetValue("Data", out byte[]? existingData) == true
                    && existingData.Length > 0) {
                    _data.Write(existingData, 0, existingData.Length);
                }
            }

            internal bool HasPictures { get; private set; }

            internal byte[] DataBytes => _data.ToArray();

            internal int AddInlinePicture(LegacyDocElement drawing, LegacyDocPart ownerPart) {
                LegacyDocWritablePicture picture = ReadInlinePicture(drawing, ownerPart);
                int dataOffset = checked((int)_data.Length);
                byte[] pictureData = CreatePicfAndOfficeArtData(picture, _nextShapeId++);
                _data.Write(pictureData, 0, pictureData.Length);
                HasPictures = true;
                return dataOffset;
            }
        }

        private static LegacyDocWritablePicture ReadInlinePicture(LegacyDocElement drawing, LegacyDocPart ownerPart) {
            LegacyDocElement inline = drawing.Child(OfficeXmlNamespaces.WordprocessingDrawing, "inline");
            if (inline.IsNone || drawing.Child(OfficeXmlNamespaces.WordprocessingDrawing, "anchor").Exists) {
                throw new NotSupportedException(
                    "Native DOC saving currently supports embedded inline pictures only. Floating or anchored pictures are not supported yet.");
            }

            LegacyDocElement extent = inline.Child(OfficeXmlNamespaces.WordprocessingDrawing, "extent");
            long? widthEmus = LegacyDocValues.Int64(extent.Attr(OfficeXmlNamespaces.None, "cx"));
            long? heightEmus = LegacyDocValues.Int64(extent.Attr(OfficeXmlNamespaces.None, "cy"));
            int widthTwips = ConvertPictureExtentToTwips(widthEmus, "width");
            int heightTwips = ConvertPictureExtentToTwips(heightEmus, "height");

            LegacyDocElement source = LegacyDocElement.None;
            int pictureCount = 0;
            foreach (LegacyDocElement picture in inline.Descendants(OfficeXmlNamespaces.Picture, "pic")) {
                if (pictureCount == 0) source = picture;
                pictureCount++;
            }
            if (pictureCount != 1) {
                throw new NotSupportedException(
                    "Native DOC saving supports an inline picture only when its DrawingML frame contains exactly one picture.");
            }

            LegacyDocElement blipFill = source.Child(OfficeXmlNamespaces.Picture, "blipFill");
            LegacyDocElement blip = blipFill.Child(OfficeXmlNamespaces.DrawingML, "blip");
            string? relationshipId = blip.Attr(OfficeXmlNamespaces.Relationships, "embed");
            if (blip.IsNone || string.IsNullOrWhiteSpace(relationshipId) || blip.Attr(OfficeXmlNamespaces.Relationships, "link") != null) {
                throw new NotSupportedException(
                    "Native DOC saving supports embedded inline pictures only. Linked or missing image relationships are not supported.");
            }

            LegacyDocElement sourceCrop = blipFill.Child(OfficeXmlNamespaces.DrawingML, "srcRect");
            OfficeImageSourceCrop crop;
            try { crop = OfficeImageSourceCrop.FromStrictFractions((ReadCropEdge(sourceCrop, "l") ?? 0) / 100000D,
                (ReadCropEdge(sourceCrop, "t") ?? 0) / 100000D, (ReadCropEdge(sourceCrop, "r") ?? 0) / 100000D, (ReadCropEdge(sourceCrop, "b") ?? 0) / 100000D); }
            catch (ArgumentOutOfRangeException exception) { throw new NotSupportedException("Native DOC saving requires an inset crop with a visible source area.", exception); }
            if (sourceCrop.HasElements) throw new NotSupportedException("Native DOC saving does not support extended crop elements.");
            if (Math.Round(crop.Left * 65536) + Math.Round(crop.Right * 65536) >= 65536
                || Math.Round(crop.Top * 65536) + Math.Round(crop.Bottom * 65536) >= 65536)
                throw new NotSupportedException("Native DOC saving cannot preserve a crop whose visible area collapses at binary crop precision.");
            LegacyDocElement transform = source.Child(OfficeXmlNamespaces.Picture, "spPr").Child(OfficeXmlNamespaces.DrawingML, "xfrm");
            if (LegacyDocValues.Int32(transform.Attr(OfficeXmlNamespaces.None, "rot")) is int rotation && rotation != 0
                || ReadXmlBoolean(transform.Attr(OfficeXmlNamespaces.None, "flipH")) == true
                || ReadXmlBoolean(transform.Attr(OfficeXmlNamespaces.None, "flipV")) == true)
                throw new NotSupportedException("Native DOC saving does not yet support rotated or mirrored inline pictures.");

            if (blip.Elements().Any(element => !element.Is(OfficeXmlNamespaces.DrawingML, "extLst"))
                || blip.Descendants().Any(element =>
                    element.LocalName is not ("extLst" or "ext" or "useLocalDpi"))) {
                throw new NotSupportedException(
                    "Native DOC saving does not yet support inline picture recoloring, transparency, or other image effects.");
            }

            // The SDK resolved the id among the owner's part relationships: a missing or external id threw
            // ArgumentOutOfRangeException, and a part that isn't an image part gave no ImagePart.
            LegacyDocRelationship? relationship = ownerPart.FindRelationship(relationshipId);
            string? imagePartName = ownerPart.GetPartById(relationshipId);
            if (relationship == null || imagePartName == null) {
                throw new NotSupportedException(
                    $"Native DOC saving cannot resolve inline picture relationship '{relationshipId}'.", new ArgumentOutOfRangeException("id"));
            }

            if (!string.Equals(relationship.Type, LegacyDocSourcePackage.ImageRelationshipType, StringComparison.Ordinal)) {
                throw new NotSupportedException(
                    $"Native DOC saving cannot resolve inline picture relationship '{relationshipId}'.");
            }

            byte[] imageBytes = ownerPart.Source.ReadPartBytes(imagePartName);
            string contentType = ownerPart.Source.GetContentType(imagePartName) ?? string.Empty;

            try {
                _ = OfficeArtBlipStoreEntryWriter.CreateEmbedded(imageBytes, contentType);
            } catch (Exception exception) when (exception is ArgumentException
                                                or IOException
                                                or NotSupportedException
                                                or OverflowException) {
                throw new NotSupportedException(
                    $"Native DOC saving cannot encode the embedded inline picture '{contentType}' as an OfficeArt BLIP. {exception.Message}",
                    exception);
            }

            return new LegacyDocWritablePicture(imageBytes, contentType, widthTwips, heightTwips, crop);
        }

        private static int? ReadCropEdge(LegacyDocElement sourceCrop, string edge) => LegacyDocValues.Int32(sourceCrop.Attr(OfficeXmlNamespaces.None, edge));

        // An xsd:boolean attribute, as the SDK's BooleanValue parsed it.
        private static bool? ReadXmlBoolean(string? value) => value == null ? null : System.Xml.XmlConvert.ToBoolean(value);

        private static int ConvertPictureExtentToTwips(long? emus, string dimensionName) {
            if (emus == null || emus.Value <= 0) {
                throw new NotSupportedException(
                    $"Native DOC saving requires a positive inline picture {dimensionName}.");
            }

            long twips = checked((emus.Value + (EmusPerTwip / 2)) / EmusPerTwip);
            if (twips < MinimumPictureTwips || twips > MaximumPictureTwips) {
                throw new NotSupportedException(
                    $"Native DOC saving supports inline picture {dimensionName}s from {MinimumPictureTwips} through {MaximumPictureTwips} twips.");
            }

            return checked((int)twips);
        }

        private static byte[] CreatePicfAndOfficeArtData(LegacyDocWritablePicture picture, uint shapeId) {
            byte[] inlineShape = CreateInlinePictureShape(shapeId, picture.Crop);
            byte[] blip = OfficeArtBlipStoreEntryWriter.CreateEmbedded(picture.ImageBytes, picture.ContentType);
            int totalLength = checked(68 + inlineShape.Length + blip.Length);
            var result = new byte[totalLength];
            WriteInt32(result, 0, totalLength);
            WriteUInt16(result, 4, 68);
            WriteUInt16(result, 6, unchecked((ushort)LegacyPictureFormatShape));
            WriteUInt16(result, 28, checked((ushort)picture.WidthTwips));
            WriteUInt16(result, 30, checked((ushort)picture.HeightTwips));
            WriteUInt16(result, 32, 1000);
            WriteUInt16(result, 34, 1000);
            Buffer.BlockCopy(inlineShape, 0, result, 68, inlineShape.Length);
            Buffer.BlockCopy(blip, 0, result, 68 + inlineShape.Length, blip.Length);
            return result;
        }

        private static byte[] CreateInlinePictureShape(uint shapeId, OfficeImageSourceCrop crop) {
            byte[] shapeProperties = CreateOfficeArtRecord(
                version: 2,
                instance: OfficeArtPictureFrameShapeType,
                type: OfficeArtFsp,
                payload: CreatePictureFspPayload(shapeId));
            byte[] options = CreateOfficeArtRecord(
                version: 3,
                instance: 15,
                type: OfficeArtFopt,
                payload: CreatePictureFoptPayload(crop));
            var anchorPayload = new byte[4];
            WriteUInt32(anchorPayload, 0, 0x80000000);
            byte[] clientAnchor = CreateOfficeArtRecord(
                version: 0,
                instance: 0,
                type: OfficeArtClientAnchor,
                payload: anchorPayload);
            byte[] payload = ConcatBytes(shapeProperties, options, clientAnchor);
            return CreateOfficeArtRecord(15, 0, OfficeArtSpContainer, payload);
        }

        private static byte[] CreatePictureFspPayload(uint shapeId) {
            var payload = new byte[8];
            WriteUInt32(payload, 0, shapeId);
            WriteUInt32(payload, 4, OfficeArtPictureShapeFlags);
            return payload;
        }

        private static byte[] CreatePictureFoptPayload(OfficeImageSourceCrop crop) {
            var properties = new (ushort Id, uint Value)[] {
                (0x0081, 0),
                (0x0082, 0),
                (0x0083, 0),
                (0x0084, 0),
                (0x0100, (uint)Math.Round(crop.Top * 65536)),
                (0x0101, (uint)Math.Round(crop.Bottom * 65536)),
                (0x0102, (uint)Math.Round(crop.Left * 65536)),
                (0x0103, (uint)Math.Round(crop.Right * 65536)),
                (0x4104, 1),
                (0x0106, 0),
                (0x013F, 0),
                (0x0181, 0x00FFFFFF),
                (0x0183, 0),
                (0x01BF, 0x00100010),
                (0x01FF, 0x00080000)
            };
            var payload = new byte[properties.Length * 6];
            for (int index = 0; index < properties.Length; index++) {
                WriteUInt16(payload, index * 6, properties[index].Id);
                WriteUInt32(payload, (index * 6) + 2, properties[index].Value);
            }

            return payload;
        }

        private static byte[] CreateOfficeArtRecord(byte version, ushort instance, ushort type, byte[] payload) {
            var result = new byte[checked(8 + payload.Length)];
            WriteUInt16(result, 0, unchecked((ushort)((instance << 4) | version)));
            WriteUInt16(result, 2, type);
            WriteUInt32(result, 4, checked((uint)payload.Length));
            Buffer.BlockCopy(payload, 0, result, 8, payload.Length);
            return result;
        }

        private static byte[] ConcatBytes(params byte[][] values) {
            int length = values.Sum(value => value.Length);
            var result = new byte[length];
            int offset = 0;
            foreach (byte[] value in values) {
                Buffer.BlockCopy(value, 0, result, offset, value.Length);
                offset += value.Length;
            }

            return result;
        }

        private readonly struct LegacyDocWritablePicture {
            internal LegacyDocWritablePicture(byte[] imageBytes, string contentType, int widthTwips, int heightTwips, OfficeImageSourceCrop crop) {
                ImageBytes = imageBytes;
                ContentType = contentType;
                WidthTwips = widthTwips;
                HeightTwips = heightTwips;
                Crop = crop;
            }

            internal byte[] ImageBytes { get; }
            internal string ContentType { get; }
            internal int WidthTwips { get; }
            internal int HeightTwips { get; }
            internal OfficeImageSourceCrop Crop { get; }
        }
    }
}
