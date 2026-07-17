using System.Reflection;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
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
        private const ushort RecordColorSchemeAtom = 0x07F0;
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
        private const ushort OfficeArtChildAnchor = 0xF00F;
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
            if (!TryReadCustomShows(presentation,
                    out LegacyPptWriterCustomShowCatalog customShows,
                    out string? customShowReason)) {
                throw new NotSupportedException(customShowReason);
            }
            var soundCatalog = new LegacyPptWriterSoundCatalog();
            if (!TryReadInteractions(presentation.Slides, soundCatalog,
                    out LegacyPptWriterInteractionCatalog interactionCatalog,
                    out string? interactionReason)) {
                throw new NotSupportedException(interactionReason);
            }
            uint firstExternalObjectId = checked((uint)
                interactionCatalog.Hyperlinks.Count + 1U);
            if (!TryReadMedia(presentation.Slides, firstExternalObjectId,
                    soundCatalog, out LegacyPptWriterMediaCatalog mediaCatalog,
                    out string? mediaReason)) {
                throw new NotSupportedException(mediaReason);
            }
            if (!TryReadOleObjects(presentation.Slides,
                    mediaCatalog.NextObjectId,
                    out LegacyPptWriterOleObjectCatalog oleCatalog,
                    out string? oleReason)) {
                throw new NotSupportedException(oleReason);
            }
            if (!TryReadPictureBulletCatalog(presentation,
                    out LegacyPptWriterPictureBulletCatalog pictureBullets,
                    out string? pictureBulletReason)) {
                throw new NotSupportedException(pictureBulletReason);
            }
            if (!TryReadPictureCatalog(presentation,
                    CreateFontCatalogForWrite(), pictureBullets,
                    convertUnsupportedTables: options?.LossPolicy
                        == PowerPointConversionLossPolicy.Allow,
                    out LegacyPptWriterPictureCatalog pictureCatalog,
                    out _, out string? pictureReason)) {
                throw new NotSupportedException(pictureReason);
            }
            if (!TryReadClassicAnimations(presentation.Slides, soundCatalog,
                    out LegacyPptWriterAnimationCatalog animationCatalog,
                    out string? animationReason)) {
                throw new NotSupportedException(animationReason);
            }
            if (!TryReadVbaProject(presentation, out byte[]? vbaProjectBytes,
                    out string? vbaReason)) {
                throw new NotSupportedException(vbaReason);
            }

            LegacyPptWriterTemplate template = Template.Value;
            var noteSources = new List<(int SlideIndex, string Text,
                NotesSlidePart? SourcePart)>();
            for (int slideIndex = 0; slideIndex < presentation.Slides.Count; slideIndex++) {
                PowerPointSlide slide = presentation.Slides[slideIndex];
                if (!ShouldWriteNotesPage(slide, out string noteText)) continue;
                noteSources.Add((slideIndex, noteText,
                    slide.SlidePart.NotesSlidePart));
            }
            bool hasHandoutMaster = presentation.OpenXmlDocument.PresentationPart?
                .HandoutMasterPart != null;
            int masterCount = presentation.OpenXmlDocument.PresentationPart?
                .SlideMasterParts.Count() ?? 0;
            var topology = new LegacyPptWriterTopology(masterCount,
                presentation.Slides.Count, noteSources.Count,
                hasHandoutMaster);
            topology.EnsurePersistObjectCapacity(checked(oleCatalog.Objects.Count
                + (vbaProjectBytes == null ? 0 : 1)));
            var notes = noteSources.Select((source, noteIndex) =>
                new LegacyPptWriterNote(source.SlideIndex, source.Text,
                    unchecked((uint)(256 + source.SlideIndex)),
                    topology.GetNotesPersistId(noteIndex),
                    topology.GetNotesDrawingId(noteIndex), source.SourcePart))
                .ToList();
            LegacyPptWriterMasterCatalog masters = ReadMasterCatalog(presentation,
                template.Document, template.MainMasterPrototypes,
                template.NotesMasterPrototype, topology,
                pictureBullets, pictureCatalog);
            uint nextAdditionalPersistId = oleCatalog.AssignPersistIds(
                topology.FirstAdditionalPersistId);
            uint? vbaProjectPersistId = vbaProjectBytes == null
                ? null
                : nextAdditionalPersistId;
            IReadOnlyDictionary<int, LegacyPptWriterNote> notesBySlide = notes.ToDictionary(
                note => note.SlideIndex);
            var slideRecords = new List<byte[]>(presentation.Slides.Count);
            var slideShapeCounts = new List<int>(presentation.Slides.Count);
            for (int index = 0; index < presentation.Slides.Count; index++) {
                PowerPointSlide slide = presentation.Slides[index];
                IReadOnlyList<PowerPointShape> supportedShapes = ReadSlideShapesForWrite(
                    slide, out _).Where(shape => IsSupportedShape(shape,
                        includeOleObjects: true,
                        includeMedia: true,
                        includePictures: true,
                        includeCharts: true,
                        includeSmartArt: true)).ToArray();
                slideShapeCounts.Add(CountDrawingShapes(supportedShapes,
                    pictureCatalog));
                uint? notesId = notesBySlide.TryGetValue(index, out LegacyPptWriterNote? note)
                    ? note.NotesId
                    : null;
                commentsBySlide.TryGetValue(slide.SlidePart.Uri.ToString(),
                    out IReadOnlyList<LegacyPptWriterComment>? comments);
                slideRecords.Add(BuildSlideRecord(
                    template.SlidePrototype, slide, supportedShapes,
                    topology.GetSlideDrawingId(index), masters.GetMasterId(slide), notesId,
                    comments ?? Array.Empty<LegacyPptWriterComment>(),
                    interactionCatalog, animationCatalog, mediaCatalog,
                    oleCatalog, pictureCatalog, masters.Fonts,
                    pictureBullets));
            }
            var notesRecords = notes.Select(note => BuildNotesRecord(template.NotesPrototype,
                note.Text, unchecked((uint)(256 + note.SlideIndex)), note.DrawingId,
                note.SourcePart, pictureCatalog)).ToArray();
            uint handoutMasterPersistId = masters.HandoutMasterPersistObject == null
                ? 0U
                : topology.HandoutMasterPersistId;
            var persistObjects = new List<byte[]>(topology.BasePersistObjectCount
                + oleCatalog.Objects.Count + (vbaProjectBytes == null ? 0 : 1)) {
                BuildDocumentRecord(template.Document, presentation, slideShapeCounts, notes,
                    interactionCatalog, customShows, soundCatalog, masters.Count,
                    masters.DrawingShapeCounts, masters.Fonts,
                    topology, handoutMasterPersistId, vbaProjectPersistId,
                    mediaCatalog, oleCatalog, pictureCatalog,
                    pictureBullets)
            };
            persistObjects.AddRange(masters.PersistObjects);
            persistObjects.Add(masters.NotesMasterPersistObject);
            persistObjects.AddRange(slideRecords);
            persistObjects.AddRange(notesRecords);
            if (masters.HandoutMasterPersistObject != null) {
                persistObjects.Add(masters.HandoutMasterPersistObject);
            }
            persistObjects.AddRange(oleCatalog.Objects.Select(
                BuildOleObjectStorageRecord));
            if (vbaProjectBytes != null) {
                persistObjects.Add(BuildVbaProjectStorageRecord(
                    vbaProjectBytes));
            }

            byte[] documentStream = BuildDocumentStream(persistObjects, presentation.Slides.Count);
            byte[] currentUserStream = BuildCurrentUserStream(FindUserEditOffset(documentStream));
            var streams = new List<OfficeCompoundStream> {
                new OfficeCompoundStream("Current User", currentUserStream),
                new OfficeCompoundStream("PowerPoint Document", documentStream)
            };
            if (pictureCatalog.Entries.Count > 0) {
                streams.Add(new OfficeCompoundStream("Pictures",
                    pictureCatalog.BuildPicturesStream()));
            }
            if (!LegacyPptPropertySetCodec.TryCreateFreshStreams(presentation,
                    out IReadOnlyList<OfficeCompoundStream> propertyStreams,
                    out string? propertyReason)) {
                throw new NotSupportedException(propertyReason);
            }
            streams.AddRange(propertyStreams);
            return OfficeCompoundFileWriter.Write(streams, PowerPointClassId);
        }

        private static bool IsSupportedShape(PowerPointShape shape) =>
            IsSupportedShape(shape, includeOleObjects: false);

        private static bool IsSupportedShape(PowerPointShape shape,
            bool includeOleObjects, bool includeMedia = false,
            bool includePictures = false, bool includeCharts = false,
            bool includeSmartArt = false) {
            if (shape is PowerPointTextBox) return true;
            if (shape is PowerPointTable) return true;
            if (includeMedia && shape is PowerPointMedia) return true;
            if (includePictures && shape is PowerPointPicture) return true;
            if (includeCharts && shape is PowerPointChart) return true;
            if (includeSmartArt && shape is PowerPointSmartArt) return true;
            if (includeOleObjects && shape is PowerPointOleObject) return true;
            if (shape is PowerPointGroupShape group) {
                return TryReadGroupForWrite(group, out _, out _);
            }
            if (shape is PowerPointConnectionShape connector) {
                return TryReadOfficeArtShapeType(connector,
                    requireConnector: true, out _, out _);
            }
            return shape is PowerPointAutoShape autoShape
                && TryReadOfficeArtShapeType(autoShape,
                    requireConnector: false, out _, out _);
        }

        internal static byte[] BuildIncrementalSlideRecord(PowerPointSlide slide, uint drawingId,
            uint masterIdRef) {
            if (slide == null) throw new ArgumentNullException(nameof(slide));
            IReadOnlyList<PowerPointShape> sourceShapes = ReadSlideShapesForWrite(slide,
                out string? layoutReason);
            IReadOnlyList<PowerPointShape> shapes = sourceShapes.Where(IsSupportedShape).ToArray();
            if (layoutReason != null || shapes.Count != sourceShapes.Count) {
                throw new InvalidOperationException("The incremental slide contains an unsupported shape.");
            }
            if (!TryReadInteractions(new[] { slide },
                    out LegacyPptWriterInteractionCatalog interactionCatalog,
                    out string? reason)) throw new NotSupportedException(reason);
            if (!TryReadClassicAnimations(new[] { slide }, interactionCatalog.Sounds,
                    out LegacyPptWriterAnimationCatalog animationCatalog,
                    out string? animationReason)) {
                throw new NotSupportedException(animationReason);
            }
            return BuildSlideRecord(Template.Value.SlidePrototype, slide, shapes, drawingId,
                masterIdRef, notesIdRef: null, ReadClassicCommentsForSlide(slide),
                interactionCatalog, animationCatalog,
                new LegacyPptWriterMediaCatalog(),
                new LegacyPptWriterOleObjectCatalog(),
                new LegacyPptWriterPictureCatalog(),
                CreateFontCatalogForWrite(),
                LegacyPptWriterPictureBulletCatalog.Empty);
        }

        internal static byte[] BuildIncrementalSlideRecord(PowerPointSlide slide,
            uint drawingId, uint masterIdRef,
            LegacyPptWriterInteractionCatalog interactionCatalog,
            LegacyPptWriterFontCatalog fonts,
            LegacyPptWriterPictureBulletCatalog pictureBullets,
            bool layoutIsIndependentMaster = false) {
            if (slide == null) throw new ArgumentNullException(nameof(slide));
            IReadOnlyList<PowerPointShape> sourceShapes = ReadSlideShapesForWrite(slide,
                out string? layoutReason);
            IReadOnlyList<PowerPointShape> shapes = sourceShapes.Where(IsSupportedShape).ToArray();
            if (layoutReason != null || shapes.Count != sourceShapes.Count) {
                throw new InvalidOperationException("The incremental slide contains an unsupported shape.");
            }
            if (!TryReadClassicAnimations(new[] { slide }, interactionCatalog.Sounds,
                    out LegacyPptWriterAnimationCatalog animationCatalog,
                    out string? animationReason)) {
                throw new NotSupportedException(animationReason);
            }
            return BuildSlideRecord(Template.Value.SlidePrototype, slide, shapes, drawingId,
                masterIdRef, notesIdRef: null, ReadClassicCommentsForSlide(slide),
                interactionCatalog, animationCatalog,
                new LegacyPptWriterMediaCatalog(),
                new LegacyPptWriterOleObjectCatalog(),
                new LegacyPptWriterPictureCatalog(),
                fonts,
                pictureBullets,
                layoutIsIndependentMaster);
        }

        private static byte[] BuildDocumentRecord(LegacyPptRecord document, PowerPointPresentation presentation,
            IReadOnlyList<int> slideShapeCounts, IReadOnlyList<LegacyPptWriterNote> notes,
            LegacyPptWriterInteractionCatalog interactionCatalog,
            LegacyPptWriterCustomShowCatalog customShows,
            LegacyPptWriterSoundCatalog soundCatalog,
            int masterCount,
            IReadOnlyDictionary<uint, int> masterDrawingShapeCounts,
            LegacyPptWriterFontCatalog fonts,
            LegacyPptWriterTopology topology,
            uint handoutMasterPersistId,
            uint? vbaProjectPersistId,
            LegacyPptWriterMediaCatalog mediaCatalog,
            LegacyPptWriterOleObjectCatalog oleCatalog,
            LegacyPptWriterPictureCatalog pictureCatalog,
            LegacyPptWriterPictureBulletCatalog pictureBullets) {
            var children = new List<byte[]>();
            bool wroteSounds = false;
            bool wroteFonts = false;
            foreach (LegacyPptRecord child in document.Children) {
                if (!wroteFonts && fonts.HasAddedFonts
                    && fonts.TryRewriteCollection(child,
                        out byte[] rewrittenFontOwner)) {
                    children.Add(rewrittenFontOwner);
                    wroteFonts = true;
                } else if (child.Type == RecordDocumentAtom) {
                    byte[] atom = child.CopyRecordBytes();
                    WriteInt32(atom, 8, ToMasterUnits(presentation.SlideSize.WidthEmus));
                    WriteInt32(atom, 12, ToMasterUnits(presentation.SlideSize.HeightEmus));
                    PatchDocumentSettings(atom, presentation,
                        topology.NotesMasterPersistId,
                        handoutMasterPersistId);
                    children.Add(atom);
                    byte[] externalObjects = BuildExternalObjectListRecord(
                        interactionCatalog, mediaCatalog, oleCatalog);
                    if (externalObjects.Length > 0) children.Add(externalObjects);
                } else if (child.Type == RecordExternalObjectList) {
                    continue;
                } else if (child.Type == RecordSoundCollection) {
                    if (soundCatalog.Sounds.Count > 0) {
                        children.Add(BuildSoundCollectionRecord(soundCatalog));
                        wroteSounds = true;
                    }
                } else if (child.Type == RecordDrawingGroup) {
                    if (!wroteFonts && fonts.HasAddedFonts
                        && !fonts.HasPrototype) {
                        children.Add(fonts.BuildCollection());
                        wroteFonts = true;
                    }
                    if (!wroteSounds && soundCatalog.Sounds.Count > 0) {
                        children.Add(BuildSoundCollectionRecord(soundCatalog));
                        wroteSounds = true;
                    }
                    children.Add(BuildDrawingGroupRecord(child,
                        masterDrawingShapeCounts, slideShapeCounts, notes.Count,
                        topology, pictureCatalog));
                } else if (child.Type == RecordHeadersFooters
                           && (child.Instance == 3 || child.Instance == 4)) {
                    children.Add(BuildDocumentHeaderFooterRecord(presentation,
                        child.Instance));
                } else if (child.Type == RecordSlideListWithText && child.Instance == 1) {
                    children.Add(BuildMasterList(masterCount));
                } else if (child.Type == RecordSlideListWithText && child.Instance == 0) {
                    children.Add(BuildSlideList(topology));
                } else if (child.Type == RecordSlideListWithText && child.Instance == 2) {
                    if (notes.Count > 0) children.Add(BuildNotesList(notes));
                } else {
                    children.Add(child.CopyRecordBytes());
                }
            }
            if (!wroteSounds && soundCatalog.Sounds.Count > 0) {
                children.Add(BuildSoundCollectionRecord(soundCatalog));
            }
            if (!wroteFonts && fonts.HasAddedFonts) {
                if (fonts.HasPrototype) {
                    throw new InvalidDataException(
                        "The embedded document template font collection could not be extended in place.");
                }
                children.Add(fonts.BuildCollection());
            }
            byte[] rebuilt = BuildContainer(RecordDocument, instance: 0, children);
            LegacyPptRecord rebuiltRecord = LegacyPptRecordReader.ReadSingle(rebuilt, 0,
                new LegacyPptImportOptions());
            var slideIdsByPartUri = presentation.Slides.Select((slide, index) =>
                    new KeyValuePair<string, uint>(slide.SlidePart.Uri.ToString(),
                        checked(unchecked((uint)index) + 256U)))
                .ToDictionary(pair => pair.Key, pair => pair.Value,
                    StringComparer.Ordinal);
            if (!TryBuildNamedShowsRecord(customShows, partUri =>
                    slideIdsByPartUri.TryGetValue(partUri, out uint slideId)
                        ? slideId
                        : (uint?)null,
                    out byte[] namedShows)
                || !TryRewriteDocumentNamedShows(rebuiltRecord, namedShows,
                    out byte[] withNamedShows)) {
                throw new InvalidDataException(
                    "The presentation custom-show list cannot be mapped to binary slide identifiers.");
            }
            rebuiltRecord = LegacyPptRecordReader.ReadSingle(withNamedShows, 0,
                new LegacyPptImportOptions());
            if (!TryRewriteDocumentHyperlinkExtensions(rebuiltRecord,
                    interactionCatalog.Hyperlinks, replaceExisting: true,
                    out byte[] withExtensions)) {
                throw new InvalidDataException(
                    "The embedded document template has malformed hyperlink extension records.");
            }
            rebuiltRecord = LegacyPptRecordReader.ReadSingle(withExtensions, 0,
                new LegacyPptImportOptions());
            if (!TryRewriteDocumentPictureBullets(rebuiltRecord,
                    pictureBullets, replaceExisting: true,
                    out byte[] withPictureBullets)) {
                throw new InvalidDataException(
                    "The embedded document template has malformed picture-bullet extension records.");
            }
            rebuiltRecord = LegacyPptRecordReader.ReadSingle(
                withPictureBullets, 0, new LegacyPptImportOptions());
            return RewriteDocumentVbaInfo(rebuiltRecord, vbaProjectPersistId);
        }

        private static void PatchDocumentSettings(byte[] atom,
            PowerPointPresentation presentation,
            uint notesMasterPersistId,
            uint handoutMasterPersistId) {
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
            WriteUInt32(atom, 32, notesMasterPersistId);
            WriteUInt32(atom, 36, handoutMasterPersistId);
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
            IReadOnlyDictionary<uint, int> masterDrawingShapeCounts,
            IReadOnlyList<int> slideShapeCounts, int notesCount,
            LegacyPptWriterTopology topology,
            LegacyPptWriterPictureCatalog pictureCatalog) {
            var drawingGroupChildren = new List<byte[]>();
            foreach (LegacyPptRecord child in drawingGroup.Children) {
                if (child.Type != OfficeArtDggContainer) {
                    drawingGroupChildren.Add(child.CopyRecordBytes());
                    continue;
                }

                var dggChildren = new List<byte[]>();
                bool wrotePictureStore = false;
                foreach (LegacyPptRecord dggChild in child.Children) {
                    if (dggChild.Type == OfficeArtDgg) {
                        dggChildren.Add(BuildDggAtom(dggChild,
                            masterDrawingShapeCounts, slideShapeCounts,
                            notesCount, topology));
                        if (pictureCatalog.Entries.Count > 0
                            && !child.Children.Any(candidate =>
                                candidate.Type == OfficeArtBStoreContainer)) {
                            dggChildren.Add(pictureCatalog.BuildStore());
                            wrotePictureStore = true;
                        }
                    } else if (dggChild.Type == OfficeArtBStoreContainer) {
                        if (pictureCatalog.Entries.Count > 0) {
                            dggChildren.Add(pictureCatalog.BuildStore());
                            wrotePictureStore = true;
                        }
                    } else {
                        dggChildren.Add(dggChild.CopyRecordBytes());
                    }
                }
                if (pictureCatalog.Entries.Count > 0 && !wrotePictureStore) {
                    dggChildren.Add(pictureCatalog.BuildStore());
                }
                drawingGroupChildren.Add(BuildContainer(OfficeArtDggContainer, child.Instance, dggChildren));
            }
            return BuildContainer(RecordDrawingGroup, drawingGroup.Instance, drawingGroupChildren);
        }

        private static byte[] BuildDggAtom(LegacyPptRecord baseAtom,
            IReadOnlyDictionary<uint, int> masterDrawingShapeCounts,
            IReadOnlyList<int> slideShapeCounts, int notesCount,
            LegacyPptWriterTopology topology) {
            var clusters = new List<KeyValuePair<uint, uint>>();
            var clusteredDrawingIds = new HashSet<uint>();
            uint templateNotesNextShapeIndex = 2U;
            int baseClusterCount = Math.Max(0, unchecked((int)baseAtom.ReadUInt32(4)) - 1);
            for (int index = 0; index < baseClusterCount && 16 + index * 8 + 8 <= baseAtom.PayloadLength; index++) {
                uint drawingId = baseAtom.ReadUInt32(16 + index * 8);
                uint nextShapeIndex = baseAtom.ReadUInt32(20 + index * 8);
                if (drawingId == 12U) {
                    templateNotesNextShapeIndex = nextShapeIndex;
                }
                if (drawingId <= topology.NotesMasterDrawingId) {
                    if (masterDrawingShapeCounts.TryGetValue(drawingId,
                            out int masterShapeCount)) {
                        nextShapeIndex = checked(
                            unchecked((uint)masterShapeCount) + 2U);
                    }
                    clusters.Add(new KeyValuePair<uint, uint>(drawingId,
                        nextShapeIndex));
                    clusteredDrawingIds.Add(drawingId);
                }
            }
            if (!clusteredDrawingIds.Contains(topology.NotesMasterDrawingId)
                && !masterDrawingShapeCounts.ContainsKey(
                    topology.NotesMasterDrawingId)) {
                clusters.Add(new KeyValuePair<uint, uint>(
                    topology.NotesMasterDrawingId,
                    templateNotesNextShapeIndex));
                clusteredDrawingIds.Add(topology.NotesMasterDrawingId);
            }
            for (int index = 0; index < slideShapeCounts.Count; index++) {
                uint drawingId = topology.GetSlideDrawingId(index);
                clusters.Add(new KeyValuePair<uint, uint>(drawingId,
                    unchecked((uint)(slideShapeCounts[index] + 2))));
                clusteredDrawingIds.Add(drawingId);
            }
            for (int index = 0; index < notesCount; index++) {
                uint drawingId = topology.GetNotesDrawingId(index);
                clusters.Add(new KeyValuePair<uint, uint>(drawingId, 4U));
                clusteredDrawingIds.Add(drawingId);
            }
            foreach (KeyValuePair<uint, int> pair in masterDrawingShapeCounts
                         .Where(pair => !clusteredDrawingIds.Contains(pair.Key))
                         .OrderBy(pair => pair.Key)) {
                clusters.Add(new KeyValuePair<uint, uint>(pair.Key,
                    checked(unchecked((uint)pair.Value) + 2U)));
            }
            clusters.Sort((left, right) => left.Key.CompareTo(right.Key));

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

        private static byte[] BuildSlideList(LegacyPptWriterTopology topology) {
            var children = new List<byte[]>(topology.SlideCount);
            for (int index = 0; index < topology.SlideCount; index++) {
                var payload = new byte[20];
                WriteUInt32(payload, 0, topology.GetSlidePersistId(index));
                WriteUInt32(payload, 4, 4);
                WriteUInt32(payload, 12, unchecked((uint)(256 + index)));
                children.Add(BuildRecord(version: 0, instance: 0, RecordSlidePersistAtom, payload));
            }
            return BuildContainer(RecordSlideListWithText, instance: 0, children);
        }

        internal static byte[] BuildLayoutPlaceholderTypes(PowerPointSlide slide,
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

        internal static byte[] BuildShapePlaceholderTypes(
            IReadOnlyList<PowerPointShape> shapes,
            LegacyPptWriterShapeContext shapeContext) {
            if (shapes == null) throw new ArgumentNullException(nameof(shapes));
            var result = new byte[8];
            foreach (PowerPointShape shape in shapes) {
                if (!TryReadPlaceholderForWrite(shape, shapeContext,
                        out LegacyPptWriterPlaceholder? placeholder,
                        out string? reason)) {
                    throw new NotSupportedException(reason);
                }
                if (placeholder == null) continue;
                int index = placeholder.Position >= 0
                    && placeholder.Position < result.Length
                    ? placeholder.Position
                    : Array.FindIndex(result, value => value == 0);
                if (index >= 0) result[index] = placeholder.Type;
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

        internal static LegacyPptSlideLayoutType MapSlideLayout(PowerPointSlide slide,
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
            if (persistObjects.Count > 0x0FFF) {
                throw new NotSupportedException(
                    $"The presentation requires {persistObjects.Count} persist objects, but a binary PowerPoint persist-directory run supports at most 4095.");
            }
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
            var mainMasters = new List<LegacyPptRecord>(11);
            for (uint id = 2; id <= 12; id++) mainMasters.Add(ReadPersist(id));
            return new LegacyPptWriterTemplate(document, mainMasters, ReadPersist(13),
                ReadPersist(14), ReadPersist(15));
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
            internal LegacyPptWriterTemplate(LegacyPptRecord document,
                IReadOnlyList<LegacyPptRecord> mainMasterPrototypes,
                LegacyPptRecord notesMasterPrototype, LegacyPptRecord slidePrototype,
                LegacyPptRecord notesPrototype) {
                Document = document;
                MainMasterPrototypes = mainMasterPrototypes;
                NotesMasterPrototype = notesMasterPrototype;
                SlidePrototype = slidePrototype;
                NotesPrototype = notesPrototype;
            }

            internal LegacyPptRecord Document { get; }
            internal IReadOnlyList<LegacyPptRecord> MainMasterPrototypes { get; }
            internal LegacyPptRecord NotesMasterPrototype { get; }
            internal LegacyPptRecord SlidePrototype { get; }
            internal LegacyPptRecord NotesPrototype { get; }
        }

    }
}
