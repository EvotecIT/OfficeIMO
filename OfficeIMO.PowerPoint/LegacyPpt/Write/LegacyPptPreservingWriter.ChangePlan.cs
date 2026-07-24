using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using System.Text;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint.LegacyPpt.Write {
    internal static partial class LegacyPptPreservingWriter {
        private static bool TryBuildModifiedPersistObjects(PowerPointPresentation presentation,
            out IReadOnlyDictionary<uint, byte[]> modifiedPersistObjects,
            out IReadOnlyList<uint> currentSlideIds,
            out byte[]? replacementPicturesStream) {
            var rewritten = new Dictionary<uint, byte[]>();
            var slideIds = new List<uint>(presentation.Slides.Count);
            modifiedPersistObjects = rewritten;
            currentSlideIds = slideIds;
            replacementPicturesStream = null;
            LegacyPptPackage? package = presentation.LegacyPptPackage;
            LegacyPptProjectionMap? projectionMap = presentation.LegacyPptProjectionMap;
            if (package == null || projectionMap == null || !presentation.HasOnlyLegacyPptPreservableChanges
                || presentation.Slides.Count > 4082) {
                return false;
            }
            if (LegacyPptWriter.HasModernComments(presentation)
                || !LegacyPptWriter.TryReadAllClassicComments(presentation,
                    out IReadOnlyDictionary<string,
                        IReadOnlyList<LegacyPptWriter.LegacyPptWriterComment>> commentsBySlide,
                    out _)) {
                return false;
            }
            var soundCatalog = new LegacyPptWriter.LegacyPptWriterSoundCatalog(
                projectionMap.Sounds, projectionMap.SoundIdSeed);
            if (!LegacyPptWriter.TryReadInteractions(presentation.Slides,
                    soundCatalog,
                    out LegacyPptWriter.LegacyPptWriterInteractionCatalog interactionCatalog,
                    out _)
                || !LegacyPptWriter.TryReadClassicAnimations(presentation.Slides,
                    soundCatalog,
                    out LegacyPptWriter.LegacyPptWriterAnimationCatalog animationCatalog,
                    out _)
                || !LegacyPptWriter.TryReadCustomShows(presentation,
                    out LegacyPptWriter.LegacyPptWriterCustomShowCatalog customShows,
                    out _)
                || !TryCreateInteractionContext(presentation, package, projectionMap,
                    interactionCatalog, out PreservingInteractionContext interactionContext)) {
                return false;
            }
            if (!TryCreateTextFontCatalog(package,
                    out LegacyPptWriter.LegacyPptWriterFontCatalog
                        textFonts)) return false;
            if (presentation.LegacyPptImportDiagnostics.Any(diagnostic =>
                    diagnostic.Code.StartsWith("PPT-PICTURE-BULLET",
                        StringComparison.Ordinal))
                || !LegacyPptWriter.TryReadPictureBulletCatalog(presentation,
                    out LegacyPptWriter
                        .LegacyPptWriterPictureBulletCatalog pictureBullets,
                    out _)) return false;
            bool customShowsChanged = !CustomShowsEqual(projectionMap, customShows);
            if (customShowsChanged && !projectionMap.CanEditCustomShows) return false;

            try {
                var oleObjectEdits = new List<LegacyPptOleObjectEdit>();
                var pictureStore = new PreservingPictureStoreUpdate(package);
                if (!TryBuildModifiedMasterPersistObjects(presentation, package,
                        projectionMap, rewritten, textFonts,
                        pictureBullets, pictureStore)) {
                    return false;
                }
                if (!TryBuildModifiedTitleMasterPersistObjects(presentation,
                        package, projectionMap, rewritten, textFonts,
                        pictureBullets, pictureStore)) {
                    return false;
                }
                if (!TryBuildModifiedSpecialMasterPersistObjects(presentation,
                        package, projectionMap, rewritten, textFonts,
                        pictureBullets, pictureStore)) {
                    return false;
                }

                var currentSlideOrder = new List<LegacyPptSlideProjection>(presentation.Slides.Count);
                var addedSlides = new List<PowerPointSlide>();
                var materializedLayoutDrawingUpdates = new Dictionary<uint,
                    MaterializedLayoutDrawingUpdate>();
                bool encounteredAddedSlide = false;
                foreach (PowerPointSlide slide in presentation.Slides) {
                    if (!projectionMap.TryGetSlide(slide, out LegacyPptSlideProjection? slideProjection)
                        || slideProjection == null) {
                        if (!LegacyPptWritePreflight.CanWriteSlideLosslessly(slide)) return false;
                        encounteredAddedSlide = true;
                        addedSlides.Add(slide);
                        continue;
                    }
                    if (encounteredAddedSlide
                        || !package.PersistObjects.TryGetValue(slideProjection.PersistId,
                            out LegacyPptPersistObject? persistObject)
                        || persistObject == null) {
                        return false;
                    }
                    currentSlideOrder.Add(slideProjection);
                    slideIds.Add(slideProjection.SlideId);

                    string currentNotes = slide.Notes.TryGetText(out string noteText)
                        ? NormalizeLogicalText(noteText)
                        : string.Empty;
                    if (slideProjection.Notes == null) {
                        if (currentNotes.Length > 0) return false;
                    } else {
                        NotesSlidePart notesPart = slide.SlidePart.NotesSlidePart
                            ?? throw new InvalidDataException(
                                "The projected binary notes page has no notes-slide part.");
                        bool notesTextChanged = !string.Equals(currentNotes,
                            NormalizeLogicalText(slideProjection.Notes.Text),
                            StringComparison.Ordinal);
                        bool notesMasterObjectsChanged = !slideProjection.Notes
                            .MasterObjectsMatch(notesPart);
                        bool notesThemeChanged = !slideProjection.Notes
                            .ThemeMatches(notesPart);
                        if (notesThemeChanged && notesPart.ThemeOverridePart?
                                .ThemeOverride == null) {
                            return false;
                        }
                        bool notesBackgroundChanged = !slideProjection.Notes
                            .BackgroundMatches(notesPart);
                        LegacyPptWriter.LegacyPptWriterBackground?
                            currentNotesBackground = null;
                        if (notesBackgroundChanged
                            && (!LegacyPptWriter.TryReadBackground(notesPart,
                                    out currentNotesBackground, out _)
                                || currentNotesBackground == null)) {
                            return false;
                        }
                        if (notesTextChanged || notesMasterObjectsChanged
                            || notesThemeChanged
                            || notesBackgroundChanged) {
                            if (!package.PersistObjects.TryGetValue(
                                    slideProjection.Notes.PersistId,
                                    out LegacyPptPersistObject? notesPersistObject)
                                || notesPersistObject == null) {
                                return false;
                            }
                            LegacyPptRecord notesRecord = LegacyPptRecordReader
                                .ReadSingle(notesPersistObject.RecordBytes, 0,
                                    new LegacyPptImportOptions());
                            byte[] notesBytes = notesRecord.CopyRecordBytes();
                            if (notesTextChanged
                                && !TryRewriteNotesRecord(notesRecord,
                                    slideProjection.Notes.Text, currentNotes,
                                    out notesBytes)) {
                                return false;
                            }
                            if (notesMasterObjectsChanged) {
                                LegacyPptRecord inheritanceRecord =
                                    LegacyPptRecordReader.ReadSingle(notesBytes,
                                        0, new LegacyPptImportOptions());
                                notesBytes = LegacyPptWriter
                                    .BuildPreservedMasterObjectInheritanceRecord(
                                        inheritanceRecord,
                                        notesPart.NotesSlide?
                                            .ShowMasterShapes?.Value != false);
                            }
                            if (notesBackgroundChanged) {
                                LegacyPptRecord backgroundRecord =
                                    LegacyPptRecordReader.ReadSingle(notesBytes,
                                        0, new LegacyPptImportOptions());
                                if (!pictureStore.TryPrepareBackground(
                                        backgroundRecord,
                                        currentNotesBackground!)) return false;
                                notesBytes = LegacyPptWriter
                                    .BuildPreservedBackgroundRecord(
                                        backgroundRecord,
                                        currentNotesBackground!,
                                        pictureStore.Catalog);
                            }
                            if (notesThemeChanged) {
                                LegacyPptRecord themedRecord =
                                    LegacyPptRecordReader.ReadSingle(notesBytes,
                                        0, new LegacyPptImportOptions());
                                notesBytes = LegacyPptWriter
                                    .BuildPreservedThemeRecord(themedRecord,
                                        notesPart, slideProjection.Notes
                                            .GetChangedClassicColorSlots(
                                                notesPart));
                            }
                            rewritten.Add(slideProjection.Notes.PersistId,
                                notesBytes);
                        }
                    }

                    SlideLayoutPart? ordinaryLayoutPart = slide.SlidePart
                        .SlideLayoutPart;
                    bool ordinaryLayoutShapesChanged = ordinaryLayoutPart != null
                        && projectionMap
                            .IsEditableProjectedOrdinaryLayoutPart(
                                ordinaryLayoutPart.Uri.ToString())
                        && !projectionMap.OrdinaryLayoutMatches(
                            ordinaryLayoutPart);
                    bool ordinaryLayoutTypeChanged = ordinaryLayoutPart != null
                        && projectionMap
                            .IsEditableProjectedOrdinaryLayoutPart(
                                ordinaryLayoutPart.Uri.ToString())
                        && !projectionMap.OrdinaryLayoutTypeMatches(
                            ordinaryLayoutPart);
                    IReadOnlyList<PowerPointShape> writableShapes =
                        LegacyPptWriter.ReadSlideShapesForWrite(slide,
                            out string? layoutShapeReason);
                    if (layoutShapeReason != null) return false;
                    IReadOnlyList<uint> addedLayoutShapeIds =
                        Array.Empty<uint>();
                    if (ordinaryLayoutShapesChanged
                        && !projectionMap
                            .TryGetOrdinaryLayoutAddedShapeIds(
                                ordinaryLayoutPart!,
                                out addedLayoutShapeIds)) return false;
                    var addedLayoutShapeIdSet = new HashSet<uint>(
                        addedLayoutShapeIds);
                    IReadOnlyList<PowerPointShape> materializedLayoutShapes =
                        writableShapes.Where(shape =>
                            LegacyPptWriter.IsLayoutShape(shape)
                            && shape.Id.HasValue
                            && addedLayoutShapeIdSet.Contains(
                                shape.Id.Value)).ToArray();
                    if (materializedLayoutShapes.Count
                        != addedLayoutShapeIdSet.Count) return false;
                    uint currentLayoutType = unchecked((uint)
                        LegacyPptWriter.MapSlideLayout(slide,
                            writableShapes));
                    byte[] currentLayoutPlaceholderTypes = LegacyPptWriter
                        .BuildLayoutPlaceholderTypes(slide, writableShapes);
                    bool layoutContractChanged = ordinaryLayoutTypeChanged
                        || ordinaryLayoutShapesChanged
                        && !slideProjection.LayoutContractMatches(
                            currentLayoutType,
                            currentLayoutPlaceholderTypes);

                    PowerPointShape[] shapes = LegacyPptWriter
                        .FlattenShapeTreeForWrite(slide.Shapes,
                            out string? groupShapeReason)
                        .ToArray();
                    if (groupShapeReason != null) return false;
                    if (shapes.Length != slideProjection.Shapes.Count) return false;
                    var editsByOfficeArtId = new Dictionary<uint, ProjectedShapeEdit>();
                    foreach (PowerPointShape shape in shapes) {
                        uint? openXmlShapeId = shape.Id;
                        if (!openXmlShapeId.HasValue
                            || !slideProjection.TryGetShape(openXmlShapeId.Value,
                                out LegacyPptShapeProjection? shapeProjection)
                            || shapeProjection == null
                            || shapeProjection.Kind == LegacyPptShapeKind.Table
                            || !MatchesProjectedKind(shape, shapeProjection.Kind)) {
                            return false;
                        }
                        if (shape is PowerPointOleObject oleShape) {
                            if (shapeProjection.OleObject == null
                                || !shapeProjection.OleObject.TryGetChange(
                                    oleShape,
                                    out LegacyPptOleObjectEdit? oleEdit)) {
                                return false;
                            }
                            if (oleEdit != null) {
                                if (oleEdit.StorageChanged) {
                                    rewritten[oleEdit.Projection.Source
                                        .PersistId] = LegacyPptWriter
                                        .BuildOleObjectStorageRecord(
                                            oleEdit.StorageBytes);
                                }
                                if (oleEdit.MetadataChanged) {
                                    oleObjectEdits.Add(oleEdit);
                                }
                            }
                        }
                        LegacyPptBounds bounds = GetBounds(shape);
                        LegacyPptBounds? changedBounds = BoundsEqual(bounds, shapeProjection.Bounds)
                            ? null
                            : bounds;
                        if (!LegacyPptWriter.TryReadPlaceholderForWrite(shape,
                                LegacyPptWriter.LegacyPptWriterShapeContext.Slide,
                                out LegacyPptWriter.LegacyPptWriterPlaceholder?
                                    currentPlaceholder, out _)) {
                            return false;
                        }
                        bool placeholderChanged = !shapeProjection
                            .PlaceholderMatches(currentPlaceholder);
                        PowerPointShape? changedShapeTransform = null;
                        if (shapeProjection.CanEditShapeTransform
                            && !shapeProjection.ShapeTransformMatches(shape)) {
                            if (!LegacyPptWriter.TryReadShapeTransform(shape,
                                    out _, out _)) {
                                return false;
                            }
                            changedShapeTransform = shape;
                        }
                        PowerPointShape? changedShapeGeometry = null;
                        if (shapeProjection.CanEditShapeGeometry
                            && !shapeProjection.ShapeGeometryMatches(shape)) {
                            if (LegacyPptShapeProjection
                                    .CreateShapeGeometryFingerprint(shape)
                                == null) {
                                return false;
                            }
                            changedShapeGeometry = shape;
                        }
                        PowerPointGroupShape? changedGroupCoordinate = null;
                        if (shapeProjection.CanEditGroupCoordinate
                            && !shapeProjection.GroupCoordinateMatches(shape)) {
                            if (shape is not PowerPointGroupShape group
                                || LegacyPptShapeProjection
                                    .CreateGroupCoordinateFingerprint(group)
                                == null) {
                                return false;
                            }
                            changedGroupCoordinate = group;
                        }
                        PowerPointShape? changedShapeVisualStyle = null;
                        if (shapeProjection.CanEditShapeVisualStyle
                            && !shapeProjection.ShapeVisualStyleMatches(
                                shape)) {
                            if (!LegacyPptWriter.TryReadShapeVisualStyle(shape,
                                    out _, out _)) {
                                return false;
                            }
                            changedShapeVisualStyle = shape;
                        }
                        PowerPointShape? changedShapeVisibility = null;
                        if (shapeProjection.CanEditShapeVisibility
                            && !shapeProjection.ShapeVisibilityMatches(shape)) {
                            if (!LegacyPptWriter
                                    .TryReadShapeVisibilityForWrite(shape,
                                        out _, out _)) {
                                return false;
                            }
                            changedShapeVisibility = shape;
                        }
                        PowerPointShape? changedShapeMetadata = null;
                        if (!shapeProjection.ShapeMetadataMatches(shape)) {
                            if (!shapeProjection.CanEditShapeMetadata
                                || !LegacyPptWriter
                                    .TryReadShapeMetadataForWrite(shape,
                                        out _, out _)) {
                                return false;
                            }
                            changedShapeMetadata = shape;
                        }
                        PowerPointPicture? changedPictureFormatting = null;
                        if (shape is PowerPointPicture picture
                            && picture is not PowerPointMedia
                            && shapeProjection.CanEditPictureFormatting
                            && !shapeProjection.PictureFormattingMatches(
                                picture)) {
                            if (!LegacyPptWriter.TryValidatePictureForWrite(
                                    picture, out _)) {
                                return false;
                            }
                            changedPictureFormatting = picture;
                        }
                        string? changedText = null;
                        PowerPointTextBox? changedTextFormatting = null;
                        PowerPointTextBox? changedTextFrame = null;
                        if (shape is PowerPointTextBox textBox) {
                            if (!shapeProjection.TextFrameMatches(textBox)) {
                                if (!shapeProjection.CanEditTextFrame
                                    || !LegacyPptWriter
                                        .TryReadTextFrameForWrite(textBox,
                                            out _, out _)) return false;
                                changedTextFrame = textBox;
                            }
                            bool formattingMatches =
                                MatchesProjectedTextFormatting(textBox,
                                    shapeProjection, slide.SlidePart);
                            if (!formattingMatches) {
                                if (!shapeProjection.CanEditTextFormatting
                                    || !LegacyPptWriter
                                        .TryReadTextBoxForWrite(textBox,
                                            textFonts, pictureBullets,
                                            out _)) return false;
                                changedTextFormatting = textBox;
                            }
                            string currentText = NormalizeLogicalText(
                                LegacyPptWriter.ReadLogicalTextForWrite(
                                    textBox));
                            if (!string.Equals(currentText, NormalizeLogicalText(shapeProjection.Text),
                                    StringComparison.Ordinal)) {
                                changedText = currentText;
                                if (shapeProjection
                                        .TextFormattingFingerprint != null) {
                                    if (!shapeProjection
                                            .CanEditTextFormatting) {
                                        return false;
                                    }
                                    if (!LegacyPptWriter
                                            .TryReadTextBoxForWrite(textBox,
                                                textFonts, pictureBullets,
                                                out _)) return false;
                                    changedTextFormatting = textBox;
                                }
                            }
                        }
                        LegacyPptWriter.LegacyPptWriterShapeInteractions currentInteractions =
                            interactionCatalog.Get(shape);
                        if (!shapeProjection.CanEditInteractions
                            && currentInteractions.HasInteractions) return false;
                        bool shapeInteractionsChanged = shapeProjection.CanEditInteractions
                            && !ShapeInteractionsEqual(shapeProjection, currentInteractions,
                                interactionCatalog, projectionMap);
                        bool textInteractionsChanged = shapeProjection.CanEditInteractions
                            && !TextInteractionsEqual(shapeProjection, currentInteractions,
                                interactionCatalog, projectionMap);
                        ProjectedInteractionEdit? interactionEdit =
                            shapeInteractionsChanged || textInteractionsChanged
                                ? new ProjectedInteractionEdit(
                                    interactionContext.Remap(currentInteractions),
                                    shapeInteractionsChanged, textInteractionsChanged)
                                : null;
                        LegacyPptWriter.LegacyPptWriterAnimation? currentAnimation =
                            animationCatalog.Get(shape);
                        bool animationChanged = !AnimationsEqual(
                            shapeProjection.Animation, currentAnimation);
                        if (animationChanged && !shapeProjection.CanEditAnimation) {
                            return false;
                        }
                        if (changedBounds.HasValue || changedText != null
                            || changedTextFormatting != null
                            || changedTextFrame != null
                            || placeholderChanged
                            || changedShapeTransform != null
                            || changedShapeGeometry != null
                            || changedGroupCoordinate != null
                            || changedShapeVisualStyle != null
                            || changedShapeVisibility != null
                            || changedShapeMetadata != null
                            || changedPictureFormatting != null
                            || interactionEdit != null || animationChanged) {
                            editsByOfficeArtId.Add(shapeProjection.OfficeArtShapeId,
                                new ProjectedShapeEdit(changedBounds, shapeProjection.Text,
                                    changedText, interactionEdit,
                                    animationChanged, currentAnimation,
                                    placeholderChanged, currentPlaceholder));
                            editsByOfficeArtId[shapeProjection.OfficeArtShapeId]
                                .ShapeTransform = changedShapeTransform;
                            editsByOfficeArtId[shapeProjection.OfficeArtShapeId]
                                .ShapeGeometry = changedShapeGeometry;
                            editsByOfficeArtId[shapeProjection.OfficeArtShapeId]
                                .GroupCoordinate = changedGroupCoordinate;
                            editsByOfficeArtId[shapeProjection.OfficeArtShapeId]
                                .ShapeVisualStyle = changedShapeVisualStyle;
                            editsByOfficeArtId[shapeProjection.OfficeArtShapeId]
                                .ShapeVisibility = changedShapeVisibility;
                            editsByOfficeArtId[shapeProjection.OfficeArtShapeId]
                                .ShapeMetadata = changedShapeMetadata;
                            editsByOfficeArtId[shapeProjection.OfficeArtShapeId]
                                .TextFormatting = changedTextFormatting;
                            editsByOfficeArtId[shapeProjection.OfficeArtShapeId]
                                .TextFrame = changedTextFrame;
                            editsByOfficeArtId[shapeProjection.OfficeArtShapeId]
                                .TextFonts = changedTextFormatting == null
                                    ? null
                                    : textFonts;
                            editsByOfficeArtId[shapeProjection.OfficeArtShapeId]
                                .PictureBullets = changedTextFormatting == null
                                    ? null
                                    : pictureBullets;
                            editsByOfficeArtId[shapeProjection.OfficeArtShapeId]
                                .PictureFormatting = changedPictureFormatting;
                        }
                    }
                    bool? hidden = slide.Hidden == slideProjection.Hidden ? null : slide.Hidden;
                    bool currentFollowsMasterObjects = slide.SlidePart.Slide?
                        .ShowMasterShapes?.Value != false;
                    bool? followsMasterObjects = slideProjection
                        .MasterObjectsMatch(slide)
                        ? null
                        : currentFollowsMasterObjects;
                    LegacyPptWriter.LegacyPptWriterHeaderFooter? currentHeaderFooter =
                        LegacyPptWriter.ReadSlideHeaderFooter(slide);
                    LegacyPptWriter.LegacyPptWriterHeaderFooter? originalHeaderFooter =
                        LegacyPptWriter.LegacyPptWriterHeaderFooter.FromLegacy(
                            slideProjection.HeaderFooter);
                    bool headerFooterChanged = originalHeaderFooter == null
                        ? currentHeaderFooter != null
                        : !originalHeaderFooter.IsEquivalentTo(currentHeaderFooter);
                    if (!LegacyPptWriter.TryReadTransition(slide, soundCatalog,
                            out LegacyPptWriter.LegacyPptWriterTransition? currentTransition,
                            out _)) return false;
                    LegacyPptWriter.LegacyPptWriterTransition? originalTransition =
                        LegacyPptWriter.LegacyPptWriterTransition.FromLegacyProjection(
                            slideProjection.Transition);
                    bool transitionChanged = originalTransition == null
                        ? currentTransition != null
                        : !originalTransition.IsEquivalentTo(currentTransition);
                    IReadOnlyList<LegacyPptWriter.LegacyPptWriterComment> currentComments =
                        commentsBySlide[slide.SlidePart.Uri.ToString()];
                    bool commentsChanged = !CommentsEqual(slideProjection.Comments,
                        currentComments);
                    bool themeChanged = !slideProjection.ThemeMatches(slide);
                    if (themeChanged && (slide.SlidePart.ThemeOverridePart
                            ?? slide.SlidePart.SlideLayoutPart?
                                .ThemeOverridePart)?.ThemeOverride == null) {
                        return false;
                    }
                    bool backgroundChanged = !slideProjection.BackgroundMatches(slide);
                    if (backgroundChanged && !slideProjection.HasExplicitBackground
                        && slide.SlidePart.Slide?.CommonSlideData?.Background == null
                        && slide.SlidePart.SlideLayoutPart is SlideLayoutPart layoutPart
                        && projectionMap.TryGetTitleMaster(layoutPart, out _)) {
                        backgroundChanged = false;
                    }
                    LegacyPptWriter.LegacyPptWriterBackground? currentBackground = null;
                    if (backgroundChanged
                        && (!LegacyPptWriter.TryReadBackground(slide,
                                out currentBackground, out _)
                            || currentBackground == null)) {
                        return false;
                    }
                    bool hasSlideRecordChanges = editsByOfficeArtId.Count > 0
                        || hidden.HasValue || followsMasterObjects.HasValue
                        || headerFooterChanged
                        || transitionChanged || commentsChanged
                        || layoutContractChanged;
                    if (!hasSlideRecordChanges && !backgroundChanged
                        && !themeChanged
                        && materializedLayoutShapes.Count == 0) continue;

                    LegacyPptRecord slideRecord = LegacyPptRecordReader.ReadSingle(persistObject.RecordBytes, 0,
                        new LegacyPptImportOptions());
                    byte[] slideBytes = slideRecord.CopyRecordBytes();
                    if (hasSlideRecordChanges) {
                        if (!TryRewriteSlide(slide, slideRecord, editsByOfficeArtId,
                                hidden, followsMasterObjects,
                                layoutContractChanged
                                    ? currentLayoutType
                                    : null,
                                layoutContractChanged
                                    ? currentLayoutPlaceholderTypes
                                    : null,
                                transitionChanged, currentTransition,
                                soundCatalog,
                                headerFooterChanged, currentHeaderFooter,
                                commentsChanged, currentComments,
                                out RecordRewrite result)
                            || !result.Changed
                            || result.PatchedShapeCount != editsByOfficeArtId.Count) {
                            return false;
                        }
                        slideBytes = result.Bytes;
                    }
                    if (materializedLayoutShapes.Count > 0) {
                        LegacyPptRecord layoutRecord = LegacyPptRecordReader
                            .ReadSingle(slideBytes, 0,
                                new LegacyPptImportOptions());
                        if (!TryAppendMaterializedLayoutShapes(layoutRecord,
                                materializedLayoutShapes, textFonts,
                                pictureBullets, out slideBytes,
                                out uint drawingId,
                                out MaterializedLayoutDrawingUpdate update)
                            || materializedLayoutDrawingUpdates.ContainsKey(
                                drawingId)) {
                            return false;
                        }
                        materializedLayoutDrawingUpdates.Add(drawingId,
                            update);
                    }
                    if (backgroundChanged) {
                        LegacyPptRecord backgroundRecord = LegacyPptRecordReader
                            .ReadSingle(slideBytes, 0,
                                new LegacyPptImportOptions());
                        if (!pictureStore.TryPrepareBackground(backgroundRecord,
                                currentBackground!)) return false;
                        slideBytes = LegacyPptWriter
                            .BuildPreservedBackgroundRecord(
                                backgroundRecord, currentBackground!,
                                pictureStore.Catalog);
                    }
                    if (themeChanged) {
                        LegacyPptRecord themedRecord = LegacyPptRecordReader
                            .ReadSingle(slideBytes, 0,
                                new LegacyPptImportOptions());
                        slideBytes = LegacyPptWriter.BuildPreservedThemeRecord(
                            themedRecord, slide.SlidePart,
                            slideProjection.GetChangedClassicColorSlots(slide));
                    }
                    rewritten.Add(slideProjection.PersistId, slideBytes);
                }
                bool originalTopologyChanged = !currentSlideOrder.Select(slide => slide.PersistId)
                    .SequenceEqual(projectionMap.Slides.Select(slide => slide.PersistId));
                if (addedSlides.Count > 0 || originalTopologyChanged) return false;
                if (materializedLayoutDrawingUpdates.Count > 0) {
                    rewritten.TryGetValue(package.DocumentPersistId,
                        out byte[]? currentDocumentBytes);
                    if (!TryRewriteDocumentDrawingClusters(package,
                            currentDocumentBytes,
                            materializedLayoutDrawingUpdates,
                            out byte[] documentWithLayoutShapes)) {
                        return false;
                    }
                    rewritten[package.DocumentPersistId] =
                        documentWithLayoutShapes;
                }
                if (customShowsChanged) {
                    rewritten.TryGetValue(package.DocumentPersistId,
                        out byte[]? currentDocumentBytes);
                    if (!TryRewriteCustomShows(package, currentDocumentBytes,
                            customShows, interactionContext,
                            out byte[] documentWithCustomShows)) {
                        return false;
                    }
                    rewritten[package.DocumentPersistId] =
                        documentWithCustomShows;
                }
                if (oleObjectEdits.Count > 0) {
                    rewritten.TryGetValue(package.DocumentPersistId,
                        out byte[]? currentDocumentBytes);
                    if (!TryRewriteOleObjectMetadata(package,
                            currentDocumentBytes, oleObjectEdits,
                            out byte[] documentWithOleObjects)) {
                        return false;
                    }
                    rewritten[package.DocumentPersistId] =
                        documentWithOleObjects;
                }
                if (interactionContext.NewHyperlinks.Count > 0) {
                    rewritten.TryGetValue(package.DocumentPersistId,
                        out byte[]? currentDocumentBytes);
                    if (!TryAppendNewHyperlinks(package, currentDocumentBytes,
                            interactionContext.NewHyperlinks, out byte[] documentWithHyperlinks)) {
                        return false;
                    }
                    rewritten[package.DocumentPersistId] = documentWithHyperlinks;
                }
                if (soundCatalog.NewSounds.Count > 0) {
                    rewritten.TryGetValue(package.DocumentPersistId,
                        out byte[]? currentDocumentBytes);
                    if (!TryAppendNewSounds(package, currentDocumentBytes,
                            soundCatalog.NewSounds,
                            out byte[] documentWithSounds)) {
                        return false;
                    }
                    rewritten[package.DocumentPersistId] = documentWithSounds;
                }
                if (textFonts.HasAddedFonts) {
                    rewritten.TryGetValue(package.DocumentPersistId,
                        out byte[]? currentDocumentBytes);
                    if (!TryRewriteTextFontCollection(package,
                            currentDocumentBytes, textFonts,
                            out byte[] documentWithFonts)) return false;
                    rewritten[package.DocumentPersistId] = documentWithFonts;
                }
                rewritten.TryGetValue(package.DocumentPersistId,
                    out byte[]? currentPictureBulletDocument);
                LegacyPptPersistObject sourceDocument = package
                    .PersistObjects[package.DocumentPersistId];
                byte[] pictureBulletSource = currentPictureBulletDocument
                    ?? sourceDocument.RecordBytes;
                LegacyPptRecord pictureBulletDocument =
                    LegacyPptRecordReader.ReadSingle(pictureBulletSource, 0,
                        new LegacyPptImportOptions());
                if (!LegacyPptWriter.TryRewriteDocumentPictureBullets(
                        pictureBulletDocument, pictureBullets,
                        replaceExisting: true,
                        out byte[] documentWithPictureBullets)) return false;
                if (!documentWithPictureBullets.SequenceEqual(
                        pictureBulletSource)) {
                    rewritten[package.DocumentPersistId] =
                        documentWithPictureBullets;
                }
                if (!TryRewriteVbaProject(presentation, package,
                        projectionMap, rewritten)) {
                    return false;
                }
                if (pictureStore.HasChanges) {
                    rewritten.TryGetValue(package.DocumentPersistId,
                        out byte[]? currentDocumentBytes);
                    if (!pictureStore.TryBuildDocumentAndPictures(
                            currentDocumentBytes,
                            out byte[] documentWithPictures,
                            out replacementPicturesStream)) return false;
                    rewritten[package.DocumentPersistId] =
                        documentWithPictures;
                }
                return true;
            } catch (Exception exception) when (exception is InvalidDataException
                                                || exception is OverflowException
                                                || exception is ArgumentException) {
                rewritten.Clear();
                return false;
            }
        }
    }
}
