using System;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing.Internal;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointSlide {
        private const string OleGraphicDataUri =
            "http://schemas.openxmlformats.org/presentationml/2006/ole";

        /// <summary>Adds an embedded OLE compound storage to the slide.</summary>
        public PowerPointOleObject AddOleObject(Stream storage, string progId,
            long left = 0L, long top = 0L, long width = 2743200L,
            long height = 1828800L,
            string contentType = PowerPointOleObject.DefaultContentType) {
            if (storage == null) throw new ArgumentNullException(nameof(storage));
            if (!storage.CanRead) {
                throw new ArgumentException(
                    "OLE storage stream must be readable.", nameof(storage));
            }
            if (string.IsNullOrWhiteSpace(progId)) {
                throw new ArgumentException(
                    "An OLE ProgID is required.", nameof(progId));
            }
            if (width <= 0) throw new ArgumentOutOfRangeException(nameof(width));
            if (height <= 0) throw new ArgumentOutOfRangeException(nameof(height));
            if (string.IsNullOrWhiteSpace(contentType)) {
                throw new ArgumentException(
                    "An embedded-part content type is required.",
                    nameof(contentType));
            }

            byte[] storageBytes = ReadOleStorage(storage);
            if (!PowerPointOleObject.TryValidateStorage(storageBytes,
                    out string? reason)) {
                throw new InvalidDataException(reason
                    ?? "The embedded object is not an OLE compound storage.");
            }

            uint shapeId = AllocateShapeIds(2);
            uint previewShapeId = shapeId + 1U;
            EmbeddedObjectPart part = _slidePart
                .AddEmbeddedObjectPart(contentType);
            using (var source = new MemoryStream(storageBytes,
                       writable: false)) {
                part.FeedData(source);
            }
            string relationshipId = _slidePart.GetIdOfPart(part);
            string name = GenerateUniqueName("Object");

            var picture = new P.Picture(
                new P.NonVisualPictureProperties(
                    new P.NonVisualDrawingProperties {
                        Id = previewShapeId,
                        Name = name
                    },
                    new P.NonVisualPictureDrawingProperties(
                        new A.PictureLocks { NoChangeAspect = true }),
                    new P.ApplicationNonVisualDrawingProperties()),
                new P.BlipFill(),
                new P.ShapeProperties(
                    new A.Transform2D(
                        new A.Offset { X = left, Y = top },
                        new A.Extents { Cx = width, Cy = height }),
                    new A.PresetGeometry(new A.AdjustValueList()) {
                        Preset = A.ShapeTypeValues.Rectangle
                    }));
            var ole = new P.OleObject(
                new P.OleObjectEmbed {
                    FollowColorScheme =
                        P.OleObjectFollowColorSchemeValues.None
                },
                picture) {
                ShapeId = "_x0000_s" + shapeId,
                Name = name,
                Id = relationshipId,
                ProgId = progId,
                ImageWidth = checked((int)Math.Min(width, int.MaxValue)),
                ImageHeight = checked((int)Math.Min(height, int.MaxValue))
            };
            var frame = new P.GraphicFrame(
                new P.NonVisualGraphicFrameProperties(
                    new P.NonVisualDrawingProperties {
                        Id = shapeId,
                        Name = name
                    },
                    new P.NonVisualGraphicFrameDrawingProperties(),
                    new P.ApplicationNonVisualDrawingProperties()),
                new P.Transform(
                    new A.Offset { X = left, Y = top },
                    new A.Extents { Cx = width, Cy = height }),
                new A.Graphic(new A.GraphicData(ole) {
                    Uri = OleGraphicDataUri
                }));

            P.CommonSlideData data = SlideRoot.CommonSlideData ??=
                new P.CommonSlideData(new P.ShapeTree());
            P.ShapeTree tree = data.ShapeTree ??= new P.ShapeTree();
            tree.Append(frame);
            return TrackShape(new PowerPointOleObject(frame, _slidePart));
        }

        private static byte[] ReadOleStorage(Stream storage) {
            if (storage.CanSeek) storage.Position = 0;
            byte[] bytes = OfficeStreamReader.ReadAllBytes(storage,
                PowerPointOleObject.MaximumStorageBytes);
            if (bytes.Length == 0) {
                throw new InvalidDataException(
                    "The embedded OLE storage is empty.");
            }
            return bytes;
        }
    }
}
