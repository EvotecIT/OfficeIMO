using System.Text;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint.LegacyPpt.Write {
    internal static partial class LegacyPptWriter {
        private static byte[] BuildDrawingRecord(LegacyPptRecord slidePrototype,
            IReadOnlyList<PowerPointShape> shapes, uint drawingId,
            LegacyPptWriterInteractionCatalog interactionCatalog,
            LegacyPptWriterAnimationCatalog animationCatalog,
            LegacyPptWriterMediaCatalog mediaCatalog,
            LegacyPptWriterOleObjectCatalog oleCatalog,
            LegacyPptWriterPictureCatalog pictureCatalog,
            LegacyPptWriterFontCatalog fonts,
            LegacyPptWriterPictureBulletCatalog pictureBullets,
            LegacyPptWriterBackground? background = null) {
            LegacyPptRecord baseDrawing = slidePrototype.Children.First(record => record.Type == RecordDrawing);
            return BuildDrawingRecord(baseDrawing, shapes, drawingId,
                interactionCatalog, animationCatalog, fonts, background,
                LegacyPptWriterShapeContext.Slide, mediaCatalog,
                oleCatalog, pictureCatalog, pictureBullets);
        }

        private static byte[] BuildDrawingRecord(LegacyPptRecord baseDrawing,
            IReadOnlyList<PowerPointShape> shapes, uint drawingId,
            LegacyPptWriterInteractionCatalog interactionCatalog,
            LegacyPptWriterAnimationCatalog animationCatalog,
            LegacyPptWriterFontCatalog fonts,
            LegacyPptWriterBackground? background,
            LegacyPptWriterShapeContext shapeContext,
            LegacyPptWriterMediaCatalog? mediaCatalog = null,
            LegacyPptWriterOleObjectCatalog? oleCatalog = null,
            LegacyPptWriterPictureCatalog? pictureCatalog = null,
            LegacyPptWriterPictureBulletCatalog? pictureBullets = null) {
            LegacyPptRecord baseDgContainer = baseDrawing.Children.First(record => record.Type == OfficeArtDgContainer);
            LegacyPptRecord baseSpgr = baseDgContainer.Children.First(record => record.Type == OfficeArtSpgrContainer);
            LegacyPptRecord baseRootShape = baseSpgr.Children.First(record => record.Type == OfficeArtSpContainer);
            LegacyPptRecord baseBackground = baseDgContainer.Children.Last(
                IsBackgroundShapeRecord);

            uint baseShapeId = drawingId << 10;
            var spgrChildren = new List<byte[]> { PatchShapeId(baseRootShape.CopyRecordBytes(), baseShapeId) };
            int contentShapeCount = CountDrawingShapes(shapes);
            uint nextShapeId = checked(baseShapeId + 2U);
            foreach (PowerPointShape shape in shapes) {
                spgrChildren.Add(shape is PowerPointGroupShape group
                    ? BuildGroupRecord(group, ref nextShapeId,
                        interactionCatalog, animationCatalog, shapeContext,
                        mediaCatalog, oleCatalog, pictureCatalog, fonts,
                        pictureBullets)
                    : shape is PowerPointTable table
                    ? BuildTableRecord(table, ref nextShapeId,
                        interactionCatalog, animationCatalog, shapeContext,
                        fonts, pictureBullets)
                    : BuildShapeRecord(shape, nextShapeId++,
                        interactionCatalog, animationCatalog, shapeContext,
                        mediaCatalog, oleCatalog, pictureCatalog, fonts,
                        pictureBullets));
            }

            byte[] backgroundShape = PatchShapeId(background == null
                    ? baseBackground.CopyRecordBytes()
                    : BuildBackgroundShapeRecord(baseBackground, background,
                        pictureCatalog),
                checked(baseShapeId + 1));
            var dgPayload = new byte[8];
            WriteUInt32(dgPayload, 0,
                unchecked((uint)(contentShapeCount + 1)));
            WriteUInt32(dgPayload, 4, checked(baseShapeId
                + unchecked((uint)contentShapeCount) + 1U));
            byte[] dgAtom = BuildRecord(version: 0, unchecked((ushort)drawingId), OfficeArtDg, dgPayload);
            byte[] spgr = BuildContainer(OfficeArtSpgrContainer, instance: 0, spgrChildren);
            var drawingChildren = new List<byte[]>(baseDgContainer.Children.Count);
            foreach (LegacyPptRecord child in baseDgContainer.Children) {
                if (child.Type == OfficeArtDg) {
                    drawingChildren.Add(dgAtom);
                } else if (ReferenceEquals(child, baseSpgr)) {
                    drawingChildren.Add(spgr);
                } else if (ReferenceEquals(child, baseBackground)) {
                    drawingChildren.Add(backgroundShape);
                } else {
                    drawingChildren.Add(child.CopyRecordBytes());
                }
            }
            byte[] dgContainer = BuildContainer(OfficeArtDgContainer,
                baseDgContainer.Instance, drawingChildren);
            var outerChildren = new List<byte[]>(baseDrawing.Children.Count);
            foreach (LegacyPptRecord child in baseDrawing.Children) {
                outerChildren.Add(ReferenceEquals(child, baseDgContainer)
                    ? dgContainer
                    : child.CopyRecordBytes());
            }
            return BuildContainer(RecordDrawing, baseDrawing.Instance, outerChildren);
        }

        internal static IReadOnlyList<byte[]>
            BuildAppendedSlideShapeRecords(
                IReadOnlyList<PowerPointShape> shapes,
                ref uint nextShapeId, LegacyPptWriterFontCatalog fonts,
                LegacyPptWriterPictureBulletCatalog pictureBullets) {
            if (shapes == null) throw new ArgumentNullException(nameof(shapes));
            if (fonts == null) throw new ArgumentNullException(nameof(fonts));
            if (pictureBullets == null) throw new ArgumentNullException(
                nameof(pictureBullets));
            var interactions = new LegacyPptWriterInteractionCatalog();
            var animations = new LegacyPptWriterAnimationCatalog(
                new Dictionary<string, LegacyPptWriterAnimation>(
                    StringComparer.Ordinal));
            var result = new List<byte[]>(shapes.Count);
            foreach (PowerPointShape shape in shapes) {
                result.Add(shape is PowerPointGroupShape group
                    ? BuildGroupRecord(group, ref nextShapeId,
                        interactions, animations,
                        LegacyPptWriterShapeContext.Slide,
                        mediaCatalog: null, oleCatalog: null,
                        pictureCatalog: null, fonts, pictureBullets)
                    : shape is PowerPointTable table
                    ? BuildTableRecord(table, ref nextShapeId,
                        interactions, animations,
                        LegacyPptWriterShapeContext.Slide,
                        fonts, pictureBullets)
                    : BuildShapeRecord(shape, nextShapeId++,
                        interactions, animations,
                        LegacyPptWriterShapeContext.Slide,
                        mediaCatalog: null, oleCatalog: null,
                        pictureCatalog: null, fonts, pictureBullets));
            }
            return result;
        }

        private static byte[] PatchShapeId(byte[] spContainer, uint shapeId) {
            // The template SpContainers begin with FSPGR/FSP or FSP. Locate the FSP record defensively.
            for (int offset = 8; offset <= spContainer.Length - 16;) {
                ushort type = ReadUInt16(spContainer, offset + 2);
                int length = checked((int)ReadUInt32(spContainer, offset + 4));
                if (type == OfficeArtFsp) {
                    WriteUInt32(spContainer, offset + 8, shapeId);
                    return spContainer;
                }
                offset = checked(offset + 8 + length);
            }
            throw new InvalidDataException("The embedded PowerPoint shape template has no FSP atom.");
        }

        private static byte[] BuildShapeRecord(PowerPointShape shape, uint shapeId,
            LegacyPptWriterInteractionCatalog interactionCatalog,
            LegacyPptWriterAnimationCatalog animationCatalog,
            LegacyPptWriterShapeContext shapeContext,
            LegacyPptWriterMediaCatalog? mediaCatalog,
            LegacyPptWriterOleObjectCatalog? oleCatalog,
            LegacyPptWriterPictureCatalog? pictureCatalog,
            LegacyPptWriterFontCatalog fonts,
            LegacyPptWriterPictureBulletCatalog? pictureBullets) {
            ushort shapeType;
            var children = new List<byte[]>();
            LegacyPptWriterShapeInteractions interactions = interactionCatalog.Get(shape);
            LegacyPptWriterAnimation? animation = animationCatalog.Get(shape);
            if (shape is PowerPointTextBox textBox) {
                if (!TryBuildTextBoxContent(textBox, fonts,
                        pictureBullets
                            ?? LegacyPptWriterPictureBulletCatalog.Empty,
                        out LegacyPptWriterTextBoxContent? textContent,
                        out string? textReason)) {
                    throw new NotSupportedException(textReason);
                }
                shapeType = 202;
                children.Add(BuildFsp(shapeType, shapeId, shape));
                byte[]? formatting = BuildShapeFoptRecord(shape);
                if (formatting != null) children.Add(formatting);
                children.Add(BuildAnchor(shape));
                byte[]? clientData = BuildClientData(shape,
                    interactions.ShapeInteractions, animation, shapeContext,
                    style9Record: textContent!.Style9Record);
                if (clientData != null) children.Add(clientData);
                children.Add(BuildTextBox(MapTextType(shape, shapeContext),
                    interactions.TextInteractions, textContent));
            } else if (shape is PowerPointAutoShape autoShape) {
                if (!TryReadOfficeArtShapeType(autoShape,
                        requireConnector: false, out shapeType,
                        out string? reason)) {
                    throw new NotSupportedException(reason);
                }
                children.Add(BuildFsp(shapeType, shapeId, shape));
                byte[]? formatting = BuildShapeFoptRecord(shape, shapeType);
                if (formatting != null) children.Add(formatting);
                children.Add(BuildAnchor(shape));
                byte[]? clientData = BuildClientData(shape,
                    interactions.ShapeInteractions, animation, shapeContext);
                if (clientData != null) children.Add(clientData);
            } else if (shape is PowerPointConnectionShape connector) {
                if (!TryReadOfficeArtShapeType(connector,
                        requireConnector: true, out shapeType,
                        out string? reason)) {
                    throw new NotSupportedException(reason);
                }
                children.Add(BuildFsp(shapeType, shapeId, shape));
                byte[]? formatting = BuildShapeFoptRecord(shape, shapeType);
                if (formatting != null) children.Add(formatting);
                children.Add(BuildAnchor(shape));
                byte[]? clientData = BuildClientData(shape,
                    interactions.ShapeInteractions, animation, shapeContext);
                if (clientData != null) children.Add(clientData);
            } else if (shape is PowerPointMedia) {
                LegacyPptWriterMedia media = mediaCatalog?.Get(shape)
                    ?? throw new InvalidOperationException(
                        "The media shape has no external-object catalog entry.");
                shapeType = 75;
                children.Add(BuildFsp(shapeType, shapeId, shape));
                byte[]? formatting = BuildShapeFoptRecord(shape);
                if (formatting != null) children.Add(formatting);
                children.Add(BuildAnchor(shape));
                byte[]? clientData = BuildClientData(shape,
                    interactions.ShapeInteractions, animation, shapeContext,
                    media.Id);
                if (clientData == null) {
                    throw new InvalidOperationException(
                        "The media shape has no external-object reference.");
                }
                children.Add(clientData);
            } else if (shape is PowerPointPicture picture) {
                LegacyPptWriterPicture catalogPicture = pictureCatalog?.Get(
                        picture)
                    ?? throw new InvalidOperationException(
                        "The picture shape has no BLIP store catalog entry.");
                shapeType = 75;
                children.Add(BuildFsp(shapeType, shapeId, shape));
                children.Add(BuildPictureFoptRecord(picture,
                    catalogPicture.OneBasedStoreIndex));
                byte[]? tertiaryPictureProperties =
                    BuildPictureTertiaryFoptRecord(picture);
                if (tertiaryPictureProperties != null) {
                    children.Add(tertiaryPictureProperties);
                }
                children.Add(BuildAnchor(shape));
                byte[]? clientData = BuildClientData(shape,
                    interactions.ShapeInteractions, animation, shapeContext);
                if (clientData != null) children.Add(clientData);
            } else if (shape is PowerPointChart chart) {
                LegacyPptWriterPicture catalogPicture = pictureCatalog?.Get(
                        chart)
                    ?? throw new InvalidOperationException(
                        "The converted chart has no BLIP store catalog entry.");
                shapeType = 75;
                children.Add(BuildFsp(shapeType, shapeId, shape));
                children.Add(BuildStaticVisualFoptRecord(
                    shape, catalogPicture.OneBasedStoreIndex));
                children.Add(BuildAnchor(shape));
                byte[]? clientData = BuildClientData(shape,
                    interactions.ShapeInteractions, animation, shapeContext);
                if (clientData != null) children.Add(clientData);
            } else if (shape is PowerPointSmartArt smartArt) {
                LegacyPptWriterPicture catalogPicture = pictureCatalog?.Get(
                        smartArt)
                    ?? throw new InvalidOperationException(
                        "The converted SmartArt diagram has no BLIP store catalog entry.");
                shapeType = 75;
                children.Add(BuildFsp(shapeType, shapeId, shape));
                children.Add(BuildStaticVisualFoptRecord(
                    shape, catalogPicture.OneBasedStoreIndex));
                children.Add(BuildAnchor(shape));
                byte[]? clientData = BuildClientData(shape,
                    interactions.ShapeInteractions, animation, shapeContext);
                if (clientData != null) children.Add(clientData);
            } else if (shape is PowerPointOleObject) {
                LegacyPptWriterOleObject ole = oleCatalog?.Get(shape)
                    ?? throw new InvalidOperationException(
                        "The OLE shape has no external-object catalog entry.");
                shapeType = 75;
                children.Add(BuildFsp(shapeType, shapeId, shape));
                byte[]? formatting;
                if (ole.Preview != null) {
                    LegacyPptWriterPicture previewPicture =
                        pictureCatalog?.Get(shape)
                        ?? throw new InvalidOperationException(
                            "The OLE preview has no BLIP store catalog entry.");
                    formatting = BuildOlePreviewFoptRecord(
                        (PowerPointOleObject)shape, ole.Preview,
                        previewPicture.OneBasedStoreIndex);
                } else {
                    formatting = BuildShapeFoptRecord(shape);
                }
                if (formatting != null) children.Add(formatting);
                if (ole.Preview != null) {
                    byte[]? tertiaryPictureProperties =
                        BuildPictureTertiaryFoptRecord(ole.Preview);
                    if (tertiaryPictureProperties != null) {
                        children.Add(tertiaryPictureProperties);
                    }
                }
                children.Add(BuildAnchor(shape));
                byte[]? clientData = BuildClientData(shape,
                    interactions.ShapeInteractions, animation, shapeContext,
                    ole.Id);
                if (clientData == null) {
                    throw new InvalidOperationException(
                        "The OLE shape has no external-object reference.");
                }
                children.Add(clientData);
            } else {
                throw new InvalidOperationException("Preflight admitted an unsupported PowerPoint shape.");
            }
            return BuildContainer(OfficeArtSpContainer, instance: 0, children);
        }

        private static byte[] BuildFsp(ushort shapeType, uint shapeId,
            PowerPointShape shape, bool isGroup = false) {
            var payload = new byte[8];
            WriteUInt32(payload, 0, shapeId);
            uint flags = GetShapeFspFlags(shape);
            if (shape.Element.Parent is P.GroupShape) flags |= 1U << 1;
            if (isGroup) flags |= 1U;
            WriteUInt32(payload, 4, flags);
            return BuildRecord(version: 2, shapeType, OfficeArtFsp, payload);
        }

        private static byte[] BuildAnchor(PowerPointShape shape) {
            int left = ToMasterUnits(shape.Left);
            int top = ToMasterUnits(shape.Top);
            int right = checked(left + ToMasterUnits(shape.Width));
            int bottom = checked(top + ToMasterUnits(shape.Height));
            if (shape.Element.Parent is P.GroupShape) {
                var childPayload = new byte[16];
                WriteInt32(childPayload, 0, left);
                WriteInt32(childPayload, 4, top);
                WriteInt32(childPayload, 8, right);
                WriteInt32(childPayload, 12, bottom);
                return BuildRecord(version: 0, instance: 0,
                    OfficeArtChildAnchor, childPayload);
            }
            if (FitsInt16(left) && FitsInt16(top) && FitsInt16(right) && FitsInt16(bottom)) {
                var payload = new byte[8];
                WriteInt16(payload, 0, unchecked((short)top));
                WriteInt16(payload, 2, unchecked((short)left));
                WriteInt16(payload, 4, unchecked((short)right));
                WriteInt16(payload, 6, unchecked((short)bottom));
                return BuildRecord(version: 0, instance: 0, OfficeArtClientAnchor, payload);
            }
            var largePayload = new byte[16];
            WriteInt32(largePayload, 0, top);
            WriteInt32(largePayload, 4, left);
            WriteInt32(largePayload, 8, right);
            WriteInt32(largePayload, 12, bottom);
            return BuildRecord(version: 0, instance: 0, OfficeArtClientAnchor, largePayload);
        }

        private static byte[]? BuildClientData(PowerPointShape shape,
            IReadOnlyList<LegacyPptWriterInteraction> interactions,
            LegacyPptWriterAnimation? animation,
            LegacyPptWriterShapeContext shapeContext,
            uint? externalObjectId = null,
            byte[]? style9Record = null) {
            var children = new List<byte[]>();
            if (externalObjectId.HasValue) {
                children.Add(BuildExternalObjectReferenceAtom(
                    externalObjectId.Value));
            }
            if (!TryReadPlaceholderForWrite(shape, shapeContext,
                    out LegacyPptWriterPlaceholder? placeholder,
                    out string? placeholderReason)) {
                throw new NotSupportedException(placeholderReason);
            }
            if (placeholder != null) {
                children.Add(BuildPlaceholderAtom(placeholder.Position,
                    placeholder.Type, placeholder.Size));
            }
            if (animation != null) children.Add(BuildAnimationInfoRecord(animation));
            foreach (LegacyPptWriterInteraction interaction in interactions) {
                children.Add(BuildInteractiveInfoRecord(interaction));
            }
            if (style9Record != null) {
                children.Add(BuildShapePpt9ProgrammableTagsRecord(
                    style9Record));
            }
            return children.Count == 0
                ? null
                : BuildContainer(OfficeArtClientData, instance: 0, children);
        }

        internal static byte[] BuildPlaceholderAtom(int position, byte placeholderType,
            byte placeholderSize) {
            var payload = new byte[8];
            WriteInt32(payload, 0, position);
            payload[4] = placeholderType;
            payload[5] = placeholderSize;
            return BuildRecord(version: 0, instance: 0, RecordPlaceholder, payload);
        }

        internal static bool TryReadPlaceholderForWrite(PowerPointShape shape,
            LegacyPptWriterShapeContext shapeContext,
            out LegacyPptWriterPlaceholder? placeholder, out string? reason) {
            if (shape == null) throw new ArgumentNullException(nameof(shape));
            P.PlaceholderShape? source = shape.Element switch {
                P.Shape value => value.NonVisualShapeProperties?
                    .ApplicationNonVisualDrawingProperties?.PlaceholderShape,
                P.ConnectionShape value => value.NonVisualConnectionShapeProperties?
                    .ApplicationNonVisualDrawingProperties?.PlaceholderShape,
                P.Picture value => value.NonVisualPictureProperties?
                    .ApplicationNonVisualDrawingProperties?.PlaceholderShape,
                P.GraphicFrame value => value.NonVisualGraphicFrameProperties?
                    .ApplicationNonVisualDrawingProperties?.PlaceholderShape,
                _ => null
            };
            if (source == null) {
                placeholder = null;
                reason = null;
                return true;
            }
            if (source.HasCustomPrompt?.Value == true || source.HasChildren) {
                placeholder = null;
                reason = "The placeholder uses a custom prompt or extension that has no binary PowerPoint mapping.";
                return false;
            }
            P.PlaceholderValues type = source.Type?.Value
                ?? P.PlaceholderValues.Object;
            bool vertical = source.Orientation?.Value
                == P.DirectionValues.Vertical;
            bool supportsVertical = shapeContext is not (
                    LegacyPptWriterShapeContext.MainMaster
                    or LegacyPptWriterShapeContext.NotesMaster)
                && (type == P.PlaceholderValues.Title
                    || type == P.PlaceholderValues.Body
                    || type == P.PlaceholderValues.Object);
            if (vertical && !supportsVertical) {
                placeholder = null;
                reason = "The placeholder uses a vertical orientation that has no equivalent binary placeholder kind in this shape context.";
                return false;
            }
            byte mappedType = MapPlaceholder(type,
                source.Orientation?.Value, shapeContext);
            uint index = source.Index?.Value ?? 0U;
            if (mappedType == 0 || index > int.MaxValue) {
                placeholder = null;
                reason = "The placeholder type or index cannot be represented by a binary PlaceholderAtom.";
                return false;
            }
            placeholder = new LegacyPptWriterPlaceholder(checked((int)index),
                mappedType, MapPlaceholderSize(source.Size?.Value));
            reason = null;
            return true;
        }

        private static byte[] BuildTextBox(string text, uint textType = 0U,
            IReadOnlyList<LegacyPptWriterTextInteraction>? textInteractions = null) {
            string normalized = (text ?? string.Empty).Replace("\r\n", "\r").Replace("\n", "\r");
            var headerPayload = new byte[4];
            WriteUInt32(headerPayload, 0, textType);
            byte[] header = BuildRecord(version: 0, instance: 0, RecordTextHeader,
                headerPayload);
            byte[] chars = BuildRecord(version: 0, instance: 0, RecordTextChars,
                Encoding.Unicode.GetBytes(normalized));
            var children = new List<byte[]> { header, chars };
            foreach (LegacyPptWriterTextInteraction interaction in textInteractions
                         ?? Array.Empty<LegacyPptWriterTextInteraction>()) {
                children.Add(BuildInteractiveInfoRecord(interaction.Interaction));
                children.Add(BuildTextInteractiveInfoRecord(interaction));
            }
            return BuildContainer(OfficeArtClientTextbox, instance: 0, children);
        }

        private static byte[] BuildTextBox(uint textType,
            IReadOnlyList<LegacyPptWriterTextInteraction> textInteractions,
            LegacyPptWriterTextBoxContent content) {
            var headerPayload = new byte[4];
            WriteUInt32(headerPayload, 0, textType);
            byte[] header = BuildRecord(version: 0, instance: 0,
                RecordTextHeader, headerPayload);
            byte[] chars = BuildRecord(version: 0, instance: 0,
                RecordTextChars, Encoding.Unicode.GetBytes(content.Text));
            var children = new List<byte[]> { header, chars };
            if (content.StyleRecord != null) {
                children.Add(content.StyleRecord);
            }
            children.AddRange(content.FieldRecords);
            if (content.SpecialInfoRecord != null) {
                children.Add(content.SpecialInfoRecord);
            }
            foreach (LegacyPptWriterTextInteraction interaction
                     in textInteractions) {
                children.Add(BuildInteractiveInfoRecord(
                    interaction.Interaction));
                children.Add(BuildTextInteractiveInfoRecord(interaction));
            }
            if (content.RulerRecord != null) {
                children.Add(content.RulerRecord);
            }
            return BuildContainer(OfficeArtClientTextbox, instance: 0,
                children);
        }

        private static byte MapPlaceholder(P.PlaceholderValues? value,
            P.DirectionValues? orientation = null,
            LegacyPptWriterShapeContext shapeContext =
                LegacyPptWriterShapeContext.Slide) {
            if (!value.HasValue) return 0;
            bool vertical = orientation == P.DirectionValues.Vertical;
            if (shapeContext == LegacyPptWriterShapeContext.MainMaster) {
                if (value.Value == P.PlaceholderValues.Title) return 0x01;
                if (value.Value == P.PlaceholderValues.Body) return 0x02;
                if (value.Value == P.PlaceholderValues.CenteredTitle) return 0x03;
                if (value.Value == P.PlaceholderValues.SubTitle) return 0x04;
            }
            if (shapeContext == LegacyPptWriterShapeContext.NotesMaster) {
                if (value.Value == P.PlaceholderValues.SlideImage) return 0x05;
                if (value.Value == P.PlaceholderValues.Body) return 0x06;
            }
            if (value.Value == P.PlaceholderValues.Title) return vertical ? (byte)0x11 : (byte)0x0D;
            if (value.Value == P.PlaceholderValues.CenteredTitle) return 0x0F;
            if (value.Value == P.PlaceholderValues.SubTitle) return 0x10;
            if (value.Value == P.PlaceholderValues.Body) return vertical ? (byte)0x12 : (byte)0x0E;
            if (value.Value == P.PlaceholderValues.Object) return vertical ? (byte)0x19 : (byte)0x13;
            if (value.Value == P.PlaceholderValues.Chart) return 0x14;
            if (value.Value == P.PlaceholderValues.Table) return 0x15;
            if (value.Value == P.PlaceholderValues.ClipArt) return 0x16;
            if (value.Value == P.PlaceholderValues.Diagram) return 0x17;
            if (value.Value == P.PlaceholderValues.Media) return 0x18;
            if (value.Value == P.PlaceholderValues.Picture) return 0x1A;
            if (value.Value == P.PlaceholderValues.SlideImage) return 0x0B;
            if (value.Value == P.PlaceholderValues.DateAndTime) return 0x07;
            if (value.Value == P.PlaceholderValues.SlideNumber) return 0x08;
            if (value.Value == P.PlaceholderValues.Footer) return 0x09;
            if (value.Value == P.PlaceholderValues.Header) return 0x0A;
            return 0;
        }

        private static uint MapTextType(PowerPointShape shape,
            LegacyPptWriterShapeContext shapeContext) {
            if (shapeContext == LegacyPptWriterShapeContext.NotesMaster) {
                return (uint)LegacyPptTextType.Notes;
            }
            P.PlaceholderValues? type = shape.ShapePlaceholderType;
            if (type == P.PlaceholderValues.Title) {
                return (uint)LegacyPptTextType.Title;
            }
            if (type == P.PlaceholderValues.CenteredTitle) {
                return (uint)LegacyPptTextType.CenterTitle;
            }
            if (type == P.PlaceholderValues.Body) {
                return (uint)LegacyPptTextType.Body;
            }
            return (uint)LegacyPptTextType.Other;
        }

        private static byte MapPlaceholderSize(P.PlaceholderSizeValues? size) {
            if (size == P.PlaceholderSizeValues.Half) return 0x01;
            if (size == P.PlaceholderSizeValues.Quarter) return 0x02;
            return 0x00;
        }

        internal enum LegacyPptWriterShapeContext {
            Slide,
            MainMaster,
            NotesMaster,
            HandoutMaster
        }

        internal sealed class LegacyPptWriterPlaceholder {
            internal LegacyPptWriterPlaceholder(int position, byte type,
                byte size) {
                Position = position;
                Type = type;
                Size = size;
            }

            internal int Position { get; }
            internal byte Type { get; }
            internal byte Size { get; }

            internal bool IsEquivalentTo(LegacyPptPlaceholder? source) =>
                source != null && Position == source.Position
                && Type == (byte)source.Kind && Size == (byte)source.Size;
        }
    }
}
