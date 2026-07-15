using System.Reflection;
using DocumentFormat.OpenXml;
using OfficeIMO.Drawing.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint.LegacyPpt.Write {
    internal static partial class LegacyPptWriter {
        private const string BaseDocumentResource = "OfficeIMO.PowerPoint.Resources.legacy-ppt-base-document.bin";
        private static readonly Guid PowerPointClassId = new("64818D10-4F9B-11CF-86EA-00AA00B929E8");
        private static readonly Lazy<LegacyPptWriterTemplate> Template = new(LoadTemplate);

        private const ushort RecordDocument = 0x03E8;
        private const ushort RecordDocumentAtom = 0x03E9;
        private const ushort RecordSlide = 0x03EE;
        private const ushort RecordSlideAtom = 0x03EF;
        private const ushort RecordNotes = 0x03F0;
        private const ushort RecordNotesAtom = 0x03F1;
        private const ushort RecordSlidePersistAtom = 0x03F3;
        private const ushort RecordSlideShowSlideInfoAtom = 0x03F9;
        private const ushort RecordDrawingGroup = 0x040B;
        private const ushort RecordDrawing = 0x040C;
        private const ushort RecordPlaceholder = 0x0BC3;
        private const ushort RecordTextHeader = 0x0F9F;
        private const ushort RecordTextChars = 0x0FA0;
        private const ushort RecordSlideListWithText = 0x0FF0;
        private const ushort RecordCurrentUser = 0x0FF6;
        private const ushort RecordUserEdit = 0x0FF5;
        private const ushort RecordPersistDirectory = 0x1772;
        private const ushort OfficeArtDggContainer = 0xF000;
        private const ushort OfficeArtDgg = 0xF006;
        private const ushort OfficeArtDgContainer = 0xF002;
        private const ushort OfficeArtSpgrContainer = 0xF003;
        private const ushort OfficeArtSpContainer = 0xF004;
        private const ushort OfficeArtDg = 0xF008;
        private const ushort OfficeArtFsp = 0xF00A;
        private const ushort OfficeArtClientTextbox = 0xF00D;
        private const ushort OfficeArtClientAnchor = 0xF010;
        private const ushort OfficeArtClientData = 0xF011;

        internal static byte[] WritePresentation(PowerPointPresentation presentation,
            PowerPointSaveOptions? options = null) {
            if (presentation == null) throw new ArgumentNullException(nameof(presentation));
            LegacyPptWritePreflightReport preflight = LegacyPptWritePreflight.Analyze(presentation);
            LegacyPptWritePreflight.ThrowIfBlocked(preflight, options);
            if (!TryReadAllClassicComments(presentation,
                    out IReadOnlyDictionary<string, IReadOnlyList<LegacyPptWriterComment>> commentsBySlide,
                    out _)) {
                commentsBySlide = new Dictionary<string, IReadOnlyList<LegacyPptWriterComment>>(
                    StringComparer.Ordinal);
            }
            if (!TryReadInteractions(presentation,
                    out LegacyPptWriterInteractionCatalog interactionCatalog,
                    out string? interactionReason)) {
                throw new NotSupportedException(interactionReason);
            }

            LegacyPptWriterTemplate template = Template.Value;
            var notes = new List<LegacyPptWriterNote>();
            for (int slideIndex = 0; slideIndex < presentation.Slides.Count; slideIndex++) {
                PowerPointSlide slide = presentation.Slides[slideIndex];
                if (!slide.Notes.TryGetText(out string noteText)
                    || string.IsNullOrWhiteSpace(noteText)) continue;
                int noteIndex = notes.Count;
                notes.Add(new LegacyPptWriterNote(slideIndex, noteText,
                    unchecked((uint)(256 + slideIndex)),
                    checked((uint)(14 + presentation.Slides.Count + noteIndex)),
                    checked((uint)(13 + presentation.Slides.Count + noteIndex))));
            }
            if (presentation.Slides.Count + notes.Count > 4082) {
                throw new NotSupportedException(
                    "Native binary PowerPoint saving supports at most 4,082 combined slide and notes persist objects.");
            }
            IReadOnlyDictionary<int, LegacyPptWriterNote> notesBySlide = notes.ToDictionary(
                note => note.SlideIndex);
            var slideRecords = new List<byte[]>(presentation.Slides.Count);
            var slideShapeCounts = new List<int>(presentation.Slides.Count);
            for (int index = 0; index < presentation.Slides.Count; index++) {
                PowerPointSlide slide = presentation.Slides[index];
                IReadOnlyList<PowerPointShape> supportedShapes = slide.Shapes.Where(IsSupportedShape).ToArray();
                slideShapeCounts.Add(supportedShapes.Count);
                uint? notesId = notesBySlide.TryGetValue(index, out LegacyPptWriterNote? note)
                    ? note.NotesId
                    : null;
                commentsBySlide.TryGetValue(slide.SlidePart.Uri.ToString(),
                    out IReadOnlyList<LegacyPptWriterComment>? comments);
                slideRecords.Add(BuildSlideRecord(template.SlidePrototype, slide, supportedShapes,
                    unchecked((uint)(13 + index)), masterIdRef: null, notesId,
                    comments ?? Array.Empty<LegacyPptWriterComment>(), interactionCatalog));
            }
            var notesRecords = notes.Select(note => BuildNotesRecord(template.NotesPrototype,
                note.Text, unchecked((uint)(256 + note.SlideIndex)), note.DrawingId)).ToArray();

            var persistObjects = new List<byte[]>(13 + slideRecords.Count + notesRecords.Length) {
                BuildDocumentRecord(template.Document, presentation, slideShapeCounts, notes,
                    interactionCatalog)
            };
            persistObjects.AddRange(template.SharedPersistObjects);
            persistObjects.AddRange(slideRecords);
            persistObjects.AddRange(notesRecords);

            byte[] documentStream = BuildDocumentStream(persistObjects, presentation.Slides.Count);
            byte[] currentUserStream = BuildCurrentUserStream(FindUserEditOffset(documentStream));
            var streams = new[] {
                new OfficeCompoundStream("Current User", currentUserStream),
                new OfficeCompoundStream("PowerPoint Document", documentStream)
            };
            return OfficeCompoundFileWriter.Write(streams, PowerPointClassId);
        }

        private static bool IsSupportedShape(PowerPointShape shape) {
            if (shape is PowerPointTextBox) return true;
            return shape is PowerPointAutoShape autoShape
                && (autoShape.ShapeType == A.ShapeTypeValues.Rectangle
                    || autoShape.ShapeType == A.ShapeTypeValues.Ellipse
                    || autoShape.ShapeType == A.ShapeTypeValues.Line);
        }

        internal static byte[] BuildIncrementalSlideRecord(PowerPointSlide slide, uint drawingId,
            uint masterIdRef) {
            if (slide == null) throw new ArgumentNullException(nameof(slide));
            IReadOnlyList<PowerPointShape> shapes = slide.Shapes.Where(IsSupportedShape).ToArray();
            if (shapes.Count != slide.Shapes.Count) {
                throw new InvalidOperationException("The incremental slide contains an unsupported shape.");
            }
            if (!TryReadInteractions(new[] { slide },
                    out LegacyPptWriterInteractionCatalog interactionCatalog,
                    out string? reason)) throw new NotSupportedException(reason);
            return BuildSlideRecord(Template.Value.SlidePrototype, slide, shapes, drawingId,
                masterIdRef, notesIdRef: null, ReadClassicCommentsForSlide(slide),
                interactionCatalog);
        }

        internal static byte[] BuildIncrementalSlideRecord(PowerPointSlide slide,
            uint drawingId, uint masterIdRef,
            LegacyPptWriterInteractionCatalog interactionCatalog) {
            if (slide == null) throw new ArgumentNullException(nameof(slide));
            IReadOnlyList<PowerPointShape> shapes = slide.Shapes.Where(IsSupportedShape).ToArray();
            if (shapes.Count != slide.Shapes.Count) {
                throw new InvalidOperationException("The incremental slide contains an unsupported shape.");
            }
            return BuildSlideRecord(Template.Value.SlidePrototype, slide, shapes, drawingId,
                masterIdRef, notesIdRef: null, ReadClassicCommentsForSlide(slide),
                interactionCatalog);
        }

        private static byte[] BuildDocumentRecord(LegacyPptRecord document, PowerPointPresentation presentation,
            IReadOnlyList<int> slideShapeCounts, IReadOnlyList<LegacyPptWriterNote> notes,
            LegacyPptWriterInteractionCatalog interactionCatalog) {
            var children = new List<byte[]>();
            foreach (LegacyPptRecord child in document.Children) {
                if (child.Type == RecordDocumentAtom) {
                    byte[] atom = child.CopyRecordBytes();
                    WriteInt32(atom, 8, ToMasterUnits(presentation.SlideSize.WidthEmus));
                    WriteInt32(atom, 12, ToMasterUnits(presentation.SlideSize.HeightEmus));
                    PatchDocumentSettings(atom, presentation);
                    children.Add(atom);
                    byte[] externalObjects = BuildExternalObjectListRecord(interactionCatalog);
                    if (externalObjects.Length > 0) children.Add(externalObjects);
                } else if (child.Type == RecordExternalObjectList) {
                    continue;
                } else if (child.Type == RecordDrawingGroup) {
                    children.Add(BuildDrawingGroupRecord(child, slideShapeCounts, notes.Count));
                } else if (child.Type == RecordHeadersFooters
                           && (child.Instance == 3 || child.Instance == 4)) {
                    children.Add(BuildDocumentHeaderFooterRecord(presentation,
                        child.Instance));
                } else if (child.Type == RecordSlideListWithText && child.Instance == 0) {
                    children.Add(BuildSlideList(presentation.Slides.Count));
                } else if (child.Type == RecordSlideListWithText && child.Instance == 2) {
                    if (notes.Count > 0) children.Add(BuildNotesList(notes));
                } else {
                    children.Add(child.CopyRecordBytes());
                }
            }
            byte[] rebuilt = BuildContainer(RecordDocument, instance: 0, children);
            LegacyPptRecord rebuiltRecord = LegacyPptRecordReader.ReadSingle(rebuilt, 0,
                new LegacyPptImportOptions());
            if (!TryRewriteDocumentHyperlinkExtensions(rebuiltRecord,
                    interactionCatalog.Hyperlinks, replaceExisting: true,
                    out byte[] withExtensions)) {
                throw new InvalidDataException(
                    "The embedded document template has malformed hyperlink extension records.");
            }
            return withExtensions;
        }

        private static void PatchDocumentSettings(byte[] atom,
            PowerPointPresentation presentation) {
            if (atom.Length < 48) {
                throw new InvalidDataException("The embedded PowerPoint DocumentAtom template is truncated.");
            }
            P.Presentation root = presentation.OpenXmlDocument.PresentationPart?.Presentation
                ?? throw new InvalidDataException("The Open XML presentation root is missing.");
            P.NotesSize? notesSize = root.NotesSize;
            if (notesSize?.Cx?.Value > 0 && notesSize.Cy?.Value > 0) {
                WriteInt32(atom, 16, ToMasterUnits(notesSize.Cx.Value));
                WriteInt32(atom, 20, ToMasterUnits(notesSize.Cy.Value));
            }
            int serverZoom = root.ServerZoom?.Value ?? 50000;
            if (serverZoom <= 0) serverZoom = 50000;
            int divisor = GreatestCommonDivisor(serverZoom, 100000);
            WriteInt32(atom, 24, serverZoom / divisor);
            WriteInt32(atom, 28, 100000 / divisor);
            int firstSlideNumber = root.FirstSlideNum?.Value ?? 1;
            WriteUInt16(atom, 40, checked((ushort)Math.Min(
                Math.Max(firstSlideNumber, 0), 9999)));
            WriteUInt16(atom, 42, MapSlideSizeType(presentation.SlideSize.Type));
            atom[44] = root.EmbedTrueTypeFonts?.Value == true ? (byte)1 : (byte)0;
            atom[45] = root.ShowSpecialPlaceholderOnTitleSlide?.Value == false
                ? (byte)1 : (byte)0;
            atom[46] = root.RightToLeft?.Value == true ? (byte)1 : (byte)0;
            atom[47] = presentation.OpenXmlDocument.PresentationPart?
                .ViewPropertiesPart?.ViewProperties?.ShowComments?.Value == true
                ? (byte)1 : (byte)0;
        }

        private static int GreatestCommonDivisor(int left, int right) {
            while (right != 0) {
                int remainder = left % right;
                left = right;
                right = remainder;
            }
            return left;
        }

        private static ushort MapSlideSizeType(P.SlideSizeValues? type) {
            if (type == P.SlideSizeValues.Letter) return 1;
            if (type == P.SlideSizeValues.A4) return 2;
            if (type == P.SlideSizeValues.Film35mm) return 3;
            if (type == P.SlideSizeValues.Overhead) return 4;
            if (type == P.SlideSizeValues.Banner) return 5;
            if (type == P.SlideSizeValues.Custom) return 6;
            return 0;
        }

        private static byte[] BuildDrawingGroupRecord(LegacyPptRecord drawingGroup,
            IReadOnlyList<int> slideShapeCounts, int notesCount) {
            var drawingGroupChildren = new List<byte[]>();
            foreach (LegacyPptRecord child in drawingGroup.Children) {
                if (child.Type != OfficeArtDggContainer) {
                    drawingGroupChildren.Add(child.CopyRecordBytes());
                    continue;
                }

                var dggChildren = new List<byte[]>();
                foreach (LegacyPptRecord dggChild in child.Children) {
                    dggChildren.Add(dggChild.Type == OfficeArtDgg
                        ? BuildDggAtom(dggChild, slideShapeCounts, notesCount)
                        : dggChild.CopyRecordBytes());
                }
                drawingGroupChildren.Add(BuildContainer(OfficeArtDggContainer, child.Instance, dggChildren));
            }
            return BuildContainer(RecordDrawingGroup, drawingGroup.Instance, drawingGroupChildren);
        }

        private static byte[] BuildDggAtom(LegacyPptRecord baseAtom,
            IReadOnlyList<int> slideShapeCounts, int notesCount) {
            var clusters = new List<KeyValuePair<uint, uint>>();
            int baseClusterCount = Math.Max(0, unchecked((int)baseAtom.ReadUInt32(4)) - 1);
            for (int index = 0; index < baseClusterCount && 16 + index * 8 + 8 <= baseAtom.PayloadLength; index++) {
                uint drawingId = baseAtom.ReadUInt32(16 + index * 8);
                uint nextShapeIndex = baseAtom.ReadUInt32(20 + index * 8);
                if (drawingId <= 12) clusters.Add(new KeyValuePair<uint, uint>(drawingId, nextShapeIndex));
            }
            for (int index = 0; index < slideShapeCounts.Count; index++) {
                clusters.Add(new KeyValuePair<uint, uint>(unchecked((uint)(13 + index)),
                    unchecked((uint)(slideShapeCounts[index] + 2))));
            }
            for (int index = 0; index < notesCount; index++) {
                clusters.Add(new KeyValuePair<uint, uint>(
                    unchecked((uint)(13 + slideShapeCounts.Count + index)), 4U));
            }

            uint maxDrawingId = clusters.Count == 0 ? 1U : clusters.Max(cluster => cluster.Key);
            uint shapeCount = unchecked((uint)clusters.Sum(cluster => checked((int)cluster.Value - 1)));
            uint lastNextShapeIndex = clusters.Count == 0 ? 1U : clusters[clusters.Count - 1].Value;
            var payload = new byte[checked(16 + clusters.Count * 8)];
            WriteUInt32(payload, 0, checked((maxDrawingId << 10) + lastNextShapeIndex));
            WriteUInt32(payload, 4, unchecked((uint)(clusters.Count + 1)));
            WriteUInt32(payload, 8, shapeCount);
            WriteUInt32(payload, 12, unchecked((uint)clusters.Count));
            for (int index = 0; index < clusters.Count; index++) {
                WriteUInt32(payload, 16 + index * 8, clusters[index].Key);
                WriteUInt32(payload, 20 + index * 8, clusters[index].Value);
            }
            return BuildRecord(version: 0, baseAtom.Instance, OfficeArtDgg, payload);
        }

        private static byte[] BuildSlideList(int slideCount) {
            var children = new List<byte[]>(slideCount);
            for (int index = 0; index < slideCount; index++) {
                var payload = new byte[20];
                WriteUInt32(payload, 0, unchecked((uint)(14 + index)));
                WriteUInt32(payload, 4, 4);
                WriteUInt32(payload, 12, unchecked((uint)(256 + index)));
                children.Add(BuildRecord(version: 0, instance: 0, RecordSlidePersistAtom, payload));
            }
            return BuildContainer(RecordSlideListWithText, instance: 0, children);
        }

        private static byte[] BuildDrawingRecord(LegacyPptRecord slidePrototype,
            IReadOnlyList<PowerPointShape> shapes, uint drawingId,
            LegacyPptWriterInteractionCatalog interactionCatalog) {
            LegacyPptRecord baseDrawing = slidePrototype.Children.First(record => record.Type == RecordDrawing);
            LegacyPptRecord baseDgContainer = baseDrawing.Children.First(record => record.Type == OfficeArtDgContainer);
            LegacyPptRecord baseSpgr = baseDgContainer.Children.First(record => record.Type == OfficeArtSpgrContainer);
            LegacyPptRecord baseRootShape = baseSpgr.Children.First(record => record.Type == OfficeArtSpContainer);
            LegacyPptRecord baseBackground = baseDgContainer.Children.Last(record => record.Type == OfficeArtSpContainer);

            uint baseShapeId = drawingId << 10;
            var spgrChildren = new List<byte[]> { PatchShapeId(baseRootShape.CopyRecordBytes(), baseShapeId) };
            for (int index = 0; index < shapes.Count; index++) {
                spgrChildren.Add(BuildShapeRecord(shapes[index],
                    checked(baseShapeId + unchecked((uint)index) + 2U), index,
                    interactionCatalog));
            }

            byte[] background = PatchShapeId(baseBackground.CopyRecordBytes(), checked(baseShapeId + 1));
            var dgPayload = new byte[8];
            WriteUInt32(dgPayload, 0, unchecked((uint)(shapes.Count + 1)));
            WriteUInt32(dgPayload, 4, checked(baseShapeId + unchecked((uint)shapes.Count) + 1U));
            byte[] dgAtom = BuildRecord(version: 0, unchecked((ushort)drawingId), OfficeArtDg, dgPayload);
            byte[] spgr = BuildContainer(OfficeArtSpgrContainer, instance: 0, spgrChildren);
            byte[] dgContainer = BuildContainer(OfficeArtDgContainer, instance: 0,
                new[] { dgAtom, spgr, background });
            return BuildContainer(RecordDrawing, instance: 0, new[] { dgContainer });
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
            int shapeIndex, LegacyPptWriterInteractionCatalog interactionCatalog) {
            ushort shapeType;
            var children = new List<byte[]>();
            LegacyPptWriterShapeInteractions interactions = interactionCatalog.Get(shape);
            if (shape is PowerPointTextBox textBox) {
                shapeType = 202;
                children.Add(BuildFsp(shapeType, shapeId));
                children.Add(BuildAnchor(shape));
                byte[]? clientData = BuildClientData(shape, shapeIndex,
                    interactions.ShapeInteractions);
                if (clientData != null) children.Add(clientData);
                children.Add(BuildTextBox(textBox.Text, textInteractions:
                    interactions.TextInteractions));
            } else if (shape is PowerPointAutoShape autoShape) {
                shapeType = autoShape.ShapeType == A.ShapeTypeValues.Ellipse ? (ushort)3
                    : autoShape.ShapeType == A.ShapeTypeValues.Line ? (ushort)20
                    : (ushort)1;
                children.Add(BuildFsp(shapeType, shapeId));
                children.Add(BuildAnchor(shape));
                byte[]? clientData = BuildClientData(shape, shapeIndex,
                    interactions.ShapeInteractions);
                if (clientData != null) children.Add(clientData);
            } else {
                throw new InvalidOperationException("Preflight admitted an unsupported PowerPoint shape.");
            }
            return BuildContainer(OfficeArtSpContainer, instance: 0, children);
        }

        private static byte[] BuildFsp(ushort shapeType, uint shapeId) {
            var payload = new byte[8];
            WriteUInt32(payload, 0, shapeId);
            WriteUInt32(payload, 4, 0x00000A00);
            return BuildRecord(version: 2, shapeType, OfficeArtFsp, payload);
        }

        private static byte[] BuildAnchor(PowerPointShape shape) {
            int left = ToMasterUnits(shape.Left);
            int top = ToMasterUnits(shape.Top);
            int right = checked(left + ToMasterUnits(shape.Width));
            int bottom = checked(top + ToMasterUnits(shape.Height));
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

        private static byte[]? BuildClientData(PowerPointShape shape, int shapeIndex,
            IReadOnlyList<LegacyPptWriterInteraction> interactions) {
            var children = new List<byte[]>();
            byte placeholderType = MapPlaceholder(shape.ShapePlaceholderType,
                shape.ShapePlaceholderOrientation);
            if (placeholderType != 0) {
                int position = checked((int)(shape.ShapePlaceholderIndex
                    ?? unchecked((uint)shapeIndex)));
                children.Add(BuildPlaceholderAtom(position, placeholderType,
                    MapPlaceholderSize(shape.ShapePlaceholderSize)));
            }
            foreach (LegacyPptWriterInteraction interaction in interactions) {
                children.Add(BuildInteractiveInfoRecord(interaction));
            }
            return children.Count == 0
                ? null
                : BuildContainer(OfficeArtClientData, instance: 0, children);
        }

        private static byte[] BuildPlaceholderAtom(int position, byte placeholderType,
            byte placeholderSize) {
            var payload = new byte[8];
            WriteInt32(payload, 0, position);
            payload[4] = placeholderType;
            payload[5] = placeholderSize;
            return BuildRecord(version: 0, instance: 0, RecordPlaceholder, payload);
        }

        private static byte[] BuildTextBox(string text, uint textType = 0U,
            IReadOnlyList<LegacyPptWriterTextInteraction>? textInteractions = null) {
            string normalized = (text ?? string.Empty).Replace("\r\n", "\r").Replace("\n", "\r");
            if (!normalized.EndsWith("\r", StringComparison.Ordinal)) normalized += "\r";
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

        private static byte MapPlaceholder(P.PlaceholderValues? value,
            P.DirectionValues? orientation = null) {
            if (!value.HasValue) return 0;
            bool vertical = orientation == P.DirectionValues.Vertical;
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

        private static byte MapPlaceholderSize(P.PlaceholderSizeValues? size) {
            if (size == P.PlaceholderSizeValues.Half) return 0x01;
            if (size == P.PlaceholderSizeValues.Quarter) return 0x02;
            return 0x00;
        }

        private static byte[] BuildLayoutPlaceholderTypes(PowerPointSlide slide,
            IReadOnlyList<PowerPointShape> shapes) {
            var result = new byte[8];
            P.ShapeTree? layoutTree = slide.SlidePart.SlideLayoutPart?.SlideLayout?
                .CommonSlideData?.ShapeTree;
            if (layoutTree != null) {
                foreach (OpenXmlElement element in layoutTree.ChildElements) {
                    P.PlaceholderShape? placeholder = element switch {
                        P.Shape item => item.NonVisualShapeProperties?
                            .ApplicationNonVisualDrawingProperties?.PlaceholderShape,
                        P.Picture item => item.NonVisualPictureProperties?
                            .ApplicationNonVisualDrawingProperties?.PlaceholderShape,
                        P.GraphicFrame item => item.NonVisualGraphicFrameProperties?
                            .ApplicationNonVisualDrawingProperties?.PlaceholderShape,
                        _ => null
                    };
                    if (placeholder == null) continue;
                    AddLayoutPlaceholderType(result, placeholder.Type?.Value,
                        placeholder.Orientation?.Value, placeholder.Index?.Value,
                        replaceExisting: false);
                }
            }
            foreach (PowerPointShape shape in shapes) {
                AddLayoutPlaceholderType(result, shape.ShapePlaceholderType,
                    shape.ShapePlaceholderOrientation, shape.ShapePlaceholderIndex,
                    replaceExisting: true);
            }
            return result;
        }

        private static void AddLayoutPlaceholderType(byte[] result,
            P.PlaceholderValues? type, P.DirectionValues? orientation, uint? requestedIndex,
            bool replaceExisting) {
            byte mapped = MapPlaceholder(type, orientation);
            if (mapped == 0) return;
            int index = requestedIndex.HasValue && requestedIndex.Value < result.Length
                ? checked((int)requestedIndex.Value)
                : Array.FindIndex(result, value => value == 0);
            if (index >= 0 && (replaceExisting || result[index] == 0)) result[index] = mapped;
        }

        private static LegacyPptSlideLayoutType MapSlideLayout(PowerPointSlide slide,
            IReadOnlyList<PowerPointShape> shapes) {
            P.SlideLayoutValues? type = slide.SlidePart.SlideLayoutPart?.SlideLayout?.Type?.Value;
            if (type == P.SlideLayoutValues.Title) return LegacyPptSlideLayoutType.TitleSlide;
            if (type == P.SlideLayoutValues.Text || type == P.SlideLayoutValues.Table
                || type == P.SlideLayoutValues.Chart || type == P.SlideLayoutValues.TextAndChart
                || type == P.SlideLayoutValues.ChartAndText || type == P.SlideLayoutValues.TextAndClipArt
                || type == P.SlideLayoutValues.ClipArtAndText || type == P.SlideLayoutValues.TextAndObject
                || type == P.SlideLayoutValues.ObjectAndText || type == P.SlideLayoutValues.TextAndMedia
                || type == P.SlideLayoutValues.MidiaAndText) {
                return LegacyPptSlideLayoutType.TitleBody;
            }
            if (type == P.SlideLayoutValues.TitleOnly) return LegacyPptSlideLayoutType.TitleOnly;
            if (type == P.SlideLayoutValues.TwoColumnText) return LegacyPptSlideLayoutType.TwoColumns;
            if (type == P.SlideLayoutValues.TwoObjects) return LegacyPptSlideLayoutType.TwoRows;
            if (type == P.SlideLayoutValues.ObjectAndTwoObjects) {
                return LegacyPptSlideLayoutType.ColumnTwoRows;
            }
            if (type == P.SlideLayoutValues.TwoObjectsAndObject) {
                return LegacyPptSlideLayoutType.TwoRowsColumn;
            }
            if (type == P.SlideLayoutValues.TwoObjectsOverText) {
                return LegacyPptSlideLayoutType.TwoColumnsRow;
            }
            if (type == P.SlideLayoutValues.FourObjects) return LegacyPptSlideLayoutType.FourObjects;
            if (type == P.SlideLayoutValues.ObjectOnly || type == P.SlideLayoutValues.Object) {
                return LegacyPptSlideLayoutType.BigObject;
            }
            if (type == P.SlideLayoutValues.VerticalTitleAndText) {
                return LegacyPptSlideLayoutType.VerticalTitleBody;
            }
            if (type == P.SlideLayoutValues.VerticalTitleAndTextOverChart) {
                return LegacyPptSlideLayoutType.VerticalTwoRows;
            }
            if (type == P.SlideLayoutValues.Blank) return LegacyPptSlideLayoutType.Blank;

            PowerPointShape[] placeholders = shapes
                .Where(shape => shape.ShapePlaceholderType.HasValue)
                .ToArray();
            if (placeholders.Length == 0) return LegacyPptSlideLayoutType.Blank;
            if (placeholders.Any(shape =>
                    shape.ShapePlaceholderType == P.PlaceholderValues.CenteredTitle)) {
                return LegacyPptSlideLayoutType.TitleSlide;
            }
            if (placeholders.Length == 1
                && placeholders[0].ShapePlaceholderType == P.PlaceholderValues.Title) {
                return LegacyPptSlideLayoutType.TitleOnly;
            }
            return LegacyPptSlideLayoutType.TitleBody;
        }

        private static byte[] BuildDocumentStream(IReadOnlyList<byte[]> persistObjects, int slideCount) {
            using var output = new MemoryStream();
            var offsets = new uint[persistObjects.Count];
            for (int index = 0; index < persistObjects.Count; index++) {
                offsets[index] = checked((uint)output.Position);
                output.Write(persistObjects[index], 0, persistObjects[index].Length);
            }

            uint directoryOffset = checked((uint)output.Position);
            var directoryPayload = new byte[checked(4 + persistObjects.Count * 4)];
            WriteUInt32(directoryPayload, 0, checked((unchecked((uint)persistObjects.Count) << 20) | 1U));
            for (int index = 0; index < offsets.Length; index++) WriteUInt32(directoryPayload, 4 + index * 4, offsets[index]);
            byte[] directory = BuildRecord(version: 0, instance: 0, RecordPersistDirectory, directoryPayload);
            output.Write(directory, 0, directory.Length);

            uint userEditOffset = checked((uint)output.Position);
            var editPayload = new byte[28];
            WriteUInt32(editPayload, 0, slideCount == 0 ? 0U : unchecked((uint)(255 + slideCount)));
            editPayload[4] = 0xBC;
            editPayload[5] = 0x0D;
            editPayload[7] = 0x03;
            WriteUInt32(editPayload, 12, directoryOffset);
            WriteUInt32(editPayload, 16, 1);
            WriteUInt32(editPayload, 20, unchecked((uint)persistObjects.Count));
            WriteUInt32(editPayload, 24, 0x00120001);
            byte[] edit = BuildRecord(version: 0, instance: 0, RecordUserEdit, editPayload);
            output.Write(edit, 0, edit.Length);
            if (userEditOffset != FindUserEditOffset(output.ToArray())) {
                throw new InvalidOperationException("The generated UserEditAtom offset is inconsistent.");
            }
            return output.ToArray();
        }

        private static uint FindUserEditOffset(byte[] documentStream) {
            int position = 0;
            int lastUserEdit = -1;
            while (position <= documentStream.Length - 8) {
                ushort type = ReadUInt16(documentStream, position + 2);
                int length = checked((int)ReadUInt32(documentStream, position + 4));
                if (type == RecordUserEdit) lastUserEdit = position;
                position = checked(position + 8 + length);
            }
            if (lastUserEdit < 0) throw new InvalidDataException("The generated PowerPoint Document stream has no UserEditAtom.");
            return unchecked((uint)lastUserEdit);
        }

        private static byte[] BuildCurrentUserStream(uint userEditOffset) {
            var payload = new byte[36];
            WriteUInt32(payload, 0, 20);
            WriteUInt32(payload, 4, 0xE391C05F);
            WriteUInt32(payload, 8, userEditOffset);
            WriteUInt16(payload, 12, 12);
            payload[14] = 0xF4;
            payload[15] = 0x03;
            payload[16] = 0x03;
            Encoding.ASCII.GetBytes("Current User", 0, 12, payload, 20);
            WriteUInt32(payload, 32, 8);
            return BuildRecord(version: 0, instance: 0, RecordCurrentUser, payload);
        }

        private static LegacyPptWriterTemplate LoadTemplate() {
            using Stream stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(BaseDocumentResource)
                ?? throw new InvalidOperationException($"Embedded resource '{BaseDocumentResource}' is missing.");
            using var memory = new MemoryStream();
            stream.CopyTo(memory);
            byte[] bytes = memory.ToArray();
            var options = new LegacyPptImportOptions();
            IReadOnlyList<LegacyPptRecord> topLevel = LegacyPptRecordReader.ReadSequence(bytes, 0, bytes.Length, options);
            LegacyPptRecord directory = topLevel.Last(record => record.Type == RecordPersistDirectory);
            var offsets = new Dictionary<uint, uint>();
            int position = 0;
            while (position < directory.PayloadLength) {
                uint packed = directory.ReadUInt32(position);
                position += 4;
                uint id = packed & 0x000FFFFF;
                int count = unchecked((int)(packed >> 20));
                for (int index = 0; index < count; index++) {
                    offsets[id + unchecked((uint)index)] = directory.ReadUInt32(position);
                    position += 4;
                }
            }
            LegacyPptRecord ReadPersist(uint id) => LegacyPptRecordReader.ReadSingle(bytes,
                checked((int)offsets[id]), options);
            LegacyPptRecord document = ReadPersist(1);
            var shared = new List<byte[]>(12);
            for (uint id = 2; id <= 13; id++) shared.Add(ReadPersist(id).CopyRecordBytes());
            return new LegacyPptWriterTemplate(document, shared, ReadPersist(14),
                ReadPersist(15));
        }

        private static byte[] BuildContainer(ushort type, ushort instance, IEnumerable<byte[]> children) =>
            BuildRecord(version: 0x0F, instance, type, Concat(children));

        private static byte[] BuildRecord(byte version, ushort instance, ushort type, byte[] payload) {
            var bytes = new byte[checked(8 + payload.Length)];
            WriteUInt16(bytes, 0, unchecked((ushort)((instance << 4) | version)));
            WriteUInt16(bytes, 2, type);
            WriteUInt32(bytes, 4, unchecked((uint)payload.Length));
            Buffer.BlockCopy(payload, 0, bytes, 8, payload.Length);
            return bytes;
        }

        private static byte[] Concat(IEnumerable<byte[]> records) {
            byte[][] values = records.ToArray();
            int length = values.Sum(record => record.Length);
            var result = new byte[length];
            int offset = 0;
            foreach (byte[] record in values) {
                Buffer.BlockCopy(record, 0, result, offset, record.Length);
                offset += record.Length;
            }
            return result;
        }

        private static int ToMasterUnits(long emus) => checked((int)Math.Round(
            emus / 1587.5d, MidpointRounding.AwayFromZero));

        private static bool FitsInt16(int value) => value >= short.MinValue && value <= short.MaxValue;

        private static ushort ReadUInt16(byte[] bytes, int offset) => unchecked((ushort)(bytes[offset] | bytes[offset + 1] << 8));

        private static uint ReadUInt32(byte[] bytes, int offset) => unchecked((uint)(bytes[offset]
            | bytes[offset + 1] << 8 | bytes[offset + 2] << 16 | bytes[offset + 3] << 24));

        private static void WriteInt16(byte[] bytes, int offset, short value) => WriteUInt16(bytes, offset, unchecked((ushort)value));

        private static void WriteUInt16(byte[] bytes, int offset, ushort value) {
            bytes[offset] = unchecked((byte)value);
            bytes[offset + 1] = unchecked((byte)(value >> 8));
        }

        private static void WriteInt32(byte[] bytes, int offset, int value) => WriteUInt32(bytes, offset, unchecked((uint)value));

        private static void WriteUInt32(byte[] bytes, int offset, uint value) {
            bytes[offset] = unchecked((byte)value);
            bytes[offset + 1] = unchecked((byte)(value >> 8));
            bytes[offset + 2] = unchecked((byte)(value >> 16));
            bytes[offset + 3] = unchecked((byte)(value >> 24));
        }

        private sealed class LegacyPptWriterTemplate {
            internal LegacyPptWriterTemplate(LegacyPptRecord document, IReadOnlyList<byte[]> sharedPersistObjects,
                LegacyPptRecord slidePrototype, LegacyPptRecord notesPrototype) {
                Document = document;
                SharedPersistObjects = sharedPersistObjects;
                SlidePrototype = slidePrototype;
                NotesPrototype = notesPrototype;
            }

            internal LegacyPptRecord Document { get; }
            internal IReadOnlyList<byte[]> SharedPersistObjects { get; }
            internal LegacyPptRecord SlidePrototype { get; }
            internal LegacyPptRecord NotesPrototype { get; }
        }

    }
}
