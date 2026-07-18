using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Model;

namespace OfficeIMO.PowerPoint.LegacyPpt.Write {
    internal static partial class LegacyPptPreservingWriter {
        private static bool TryBuildModifiedMasterPersistObjects(
            PowerPointPresentation presentation, LegacyPptPackage package,
            LegacyPptProjectionMap projectionMap,
            IDictionary<uint, byte[]> rewritten,
            LegacyPptWriter.LegacyPptWriterFontCatalog fonts,
            LegacyPptWriter.LegacyPptWriterPictureBulletCatalog
                pictureBullets,
            PreservingPictureStoreUpdate pictureStore) {
            SlideMasterPart[] masterParts = presentation.OpenXmlDocument
                .PresentationPart?.SlideMasterParts.ToArray()
                ?? Array.Empty<SlideMasterPart>();
            if (masterParts.Length != projectionMap.Masters.Count) return false;

            foreach (SlideMasterPart masterPart in masterParts) {
                if (!projectionMap.TryGetMaster(masterPart,
                        out LegacyPptMasterProjection? projection)
                    || projection == null) {
                    return false;
                }
                bool themeChanged = !projection.ThemeMatches(masterPart);
                bool backgroundChanged = !projection.BackgroundMatches(masterPart);
                bool textStylesChanged = !projection.TextStylesMatch(
                    masterPart);
                if (textStylesChanged && !projection.CanEditTextStyles) {
                    return false;
                }
                LegacyPptWriter.LegacyPptWriterBackground? background = null;
                if (backgroundChanged
                    && (!LegacyPptWriter.TryReadBackground(masterPart,
                            out background, out _)
                        || background == null)) {
                    return false;
                }
                if (!TryBuildMasterShapeEdits(masterPart, projection, fonts,
                        pictureBullets,
                        out IReadOnlyDictionary<uint, ProjectedShapeEdit>
                            shapeEdits)) {
                    return false;
                }
                if (!themeChanged && !backgroundChanged
                    && !textStylesChanged
                    && shapeEdits.Count == 0) continue;
                if (!package.PersistObjects.TryGetValue(projection.PersistId,
                        out LegacyPptPersistObject? persistObject)
                    || persistObject == null) {
                    return false;
                }
                LegacyPptRecord masterRecord = LegacyPptRecordReader.ReadSingle(
                    persistObject.RecordBytes, 0,
                    new LegacyPptImportOptions());
                byte[] masterBytes = masterRecord.CopyRecordBytes();
                if (shapeEdits.Count > 0) {
                    RecordRewrite shapeRewrite = RewriteRecord(masterRecord,
                        shapeEdits);
                    if (!shapeRewrite.Changed
                        || shapeRewrite.PatchedShapeCount != shapeEdits.Count) {
                        return false;
                    }
                    masterBytes = shapeRewrite.Bytes;
                    if (shapeEdits.Values.Any(edit =>
                            edit.RewritePlaceholder)) {
                        LegacyPptRecord placeholderRecord =
                            LegacyPptRecordReader.ReadSingle(masterBytes, 0,
                                new LegacyPptImportOptions());
                        masterBytes = LegacyPptWriter
                            .BuildPreservedPlaceholderSignatureRecord(
                                placeholderRecord,
                                LegacyPptWriter.ReadMasterShapesForWrite(
                                    masterPart, out _),
                                LegacyPptWriter
                                    .LegacyPptWriterShapeContext.MainMaster);
                    }
                }
                if (backgroundChanged) {
                    LegacyPptRecord backgroundRecord = LegacyPptRecordReader
                        .ReadSingle(masterBytes, 0,
                            new LegacyPptImportOptions());
                    if (!pictureStore.TryPrepareBackground(backgroundRecord,
                            background!)) return false;
                    masterBytes = LegacyPptWriter
                        .BuildPreservedBackgroundRecord(backgroundRecord,
                            background!, pictureStore.Catalog);
                }
                if (themeChanged) {
                    LegacyPptRecord themedRecord = LegacyPptRecordReader
                        .ReadSingle(masterBytes, 0,
                            new LegacyPptImportOptions());
                    masterBytes = LegacyPptWriter
                        .BuildPreservedMasterThemeRecord(themedRecord,
                            masterPart,
                            projection.GetChangedClassicColorSlots(masterPart));
                }
                if (textStylesChanged) {
                    LegacyPptRecord textStyleRecord = LegacyPptRecordReader
                        .ReadSingle(masterBytes, 0,
                            new LegacyPptImportOptions());
                    if (!LegacyPptWriter.TryRewriteMasterTextStyleRecords(
                            textStyleRecord,
                            masterPart.SlideMaster?.TextStyles, fonts,
                            pictureBullets,
                            out masterBytes, out _)) return false;
                }
                rewritten.Add(projection.PersistId, masterBytes);
            }
            return true;
        }

        private static bool TryBuildMasterShapeEdits(SlideMasterPart masterPart,
            LegacyPptMasterProjection projection,
            LegacyPptWriter.LegacyPptWriterFontCatalog fonts,
            LegacyPptWriter.LegacyPptWriterPictureBulletCatalog
                pictureBullets,
            out IReadOnlyDictionary<uint, ProjectedShapeEdit> edits) {
            IReadOnlyList<PowerPointShape> shapes = LegacyPptWriter
                .ReadMasterShapesForWrite(masterPart,
                    out string? unsupportedReason);
            if (unsupportedReason != null) {
                edits = new Dictionary<uint, ProjectedShapeEdit>();
                return false;
            }
            return TryBuildMasterShapeEdits(shapes, projection,
                LegacyPptWriter.LegacyPptWriterShapeContext.MainMaster,
                fonts, pictureBullets, masterPart, out edits);
        }

        private static bool TryBuildMasterShapeEdits(
            IReadOnlyList<PowerPointShape> shapes,
            LegacyPptMasterProjection projection,
            LegacyPptWriter.LegacyPptWriterShapeContext shapeContext,
            LegacyPptWriter.LegacyPptWriterFontCatalog fonts,
            LegacyPptWriter.LegacyPptWriterPictureBulletCatalog
                pictureBullets,
            OpenXmlPart ownerPart,
            out IReadOnlyDictionary<uint, ProjectedShapeEdit> edits) {
            var result = new Dictionary<uint, ProjectedShapeEdit>();
            edits = result;
            shapes = LegacyPptWriter.FlattenShapeTreeForWrite(shapes,
                out string? groupReason);
            if (groupReason != null) return false;
            if (shapes.Count != projection.Shapes.Count) return false;

            foreach (PowerPointShape shape in shapes) {
                uint? openXmlShapeId = shape.Id;
                if (!openXmlShapeId.HasValue
                    || !projection.TryGetShape(openXmlShapeId.Value,
                        out LegacyPptShapeProjection? shapeProjection)
                    || shapeProjection == null
                    || shapeProjection.Kind == LegacyPptShapeKind.Table
                    || !MatchesProjectedKind(shape, shapeProjection.Kind)) {
                    return false;
                }
                LegacyPptBounds bounds = GetBounds(shape);
                LegacyPptBounds? changedBounds = BoundsEqual(
                    bounds, shapeProjection.Bounds) ? null : bounds;
                if (!LegacyPptWriter.TryReadPlaceholderForWrite(shape,
                        shapeContext,
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
                            .CreateShapeGeometryFingerprint(shape) == null) {
                        return false;
                    }
                    changedShapeGeometry = shape;
                }
                PowerPointGroupShape? changedGroupCoordinate = null;
                if (shapeProjection.CanEditGroupCoordinate
                    && !shapeProjection.GroupCoordinateMatches(shape)) {
                    if (shape is not PowerPointGroupShape group
                        || LegacyPptShapeProjection
                            .CreateGroupCoordinateFingerprint(group) == null) {
                        return false;
                    }
                    changedGroupCoordinate = group;
                }
                PowerPointShape? changedShapeVisualStyle = null;
                if (shapeProjection.CanEditShapeVisualStyle
                    && !shapeProjection.ShapeVisualStyleMatches(shape)) {
                    if (!LegacyPptWriter.TryReadShapeVisualStyle(shape,
                            out _, out _)) {
                        return false;
                    }
                    changedShapeVisualStyle = shape;
                }
                PowerPointShape? changedShapeVisibility = null;
                if (shapeProjection.CanEditShapeVisibility
                    && !shapeProjection.ShapeVisibilityMatches(shape)) {
                    if (!LegacyPptWriter.TryReadShapeVisibilityForWrite(
                            shape, out _, out _)) {
                        return false;
                    }
                    changedShapeVisibility = shape;
                }
                PowerPointShape? changedShapeMetadata = null;
                if (!shapeProjection.ShapeMetadataMatches(shape)) {
                    if (!shapeProjection.CanEditShapeMetadata
                        || !LegacyPptWriter.TryReadShapeMetadataForWrite(
                            shape, out _, out _)) {
                        return false;
                    }
                    changedShapeMetadata = shape;
                }
                PowerPointPicture? changedPictureFormatting = null;
                if (shape is PowerPointPicture picture
                    && picture is not PowerPointMedia
                    && shapeProjection.CanEditPictureFormatting
                    && !shapeProjection.PictureFormattingMatches(picture)) {
                    if (!LegacyPptWriter.TryValidatePictureForWrite(picture,
                            out _)) {
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
                            || !LegacyPptWriter.TryReadTextFrameForWrite(
                                textBox, out _, out _)) return false;
                        changedTextFrame = textBox;
                    }
                    bool formattingMatches = MatchesProjectedTextFormatting(
                        textBox, shapeProjection, ownerPart);
                    if (!formattingMatches) {
                        if (!shapeProjection.CanEditTextFormatting
                            || !LegacyPptWriter.TryReadTextBoxForWrite(
                                textBox, fonts, pictureBullets,
                                out _)) return false;
                        changedTextFormatting = textBox;
                    }
                    string currentText = NormalizeLogicalText(
                        LegacyPptWriter.ReadLogicalTextForWrite(textBox));
                    if (!string.Equals(currentText,
                            NormalizeLogicalText(shapeProjection.Text),
                            StringComparison.Ordinal)) {
                        changedText = currentText;
                        if (shapeProjection.TextFormattingFingerprint
                                != null) {
                            if (!shapeProjection.CanEditTextFormatting) {
                                return false;
                            }
                            if (!LegacyPptWriter.TryReadTextBoxForWrite(
                                    textBox, fonts, pictureBullets,
                                    out _)) return false;
                            changedTextFormatting = textBox;
                        }
                    }
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
                    || changedPictureFormatting != null) {
                    result.Add(shapeProjection.OfficeArtShapeId,
                        new ProjectedShapeEdit(changedBounds,
                            shapeProjection.Text, changedText,
                            interactions: null, rewriteAnimation: false,
                            animation: null,
                            rewritePlaceholder: placeholderChanged,
                            placeholder: currentPlaceholder));
                    result[shapeProjection.OfficeArtShapeId]
                        .ShapeTransform = changedShapeTransform;
                    result[shapeProjection.OfficeArtShapeId]
                        .ShapeGeometry = changedShapeGeometry;
                    result[shapeProjection.OfficeArtShapeId]
                        .GroupCoordinate = changedGroupCoordinate;
                    result[shapeProjection.OfficeArtShapeId]
                        .ShapeVisualStyle = changedShapeVisualStyle;
                    result[shapeProjection.OfficeArtShapeId]
                        .ShapeVisibility = changedShapeVisibility;
                    result[shapeProjection.OfficeArtShapeId]
                        .ShapeMetadata = changedShapeMetadata;
                    result[shapeProjection.OfficeArtShapeId]
                        .TextFormatting = changedTextFormatting;
                    result[shapeProjection.OfficeArtShapeId]
                        .TextFrame = changedTextFrame;
                    result[shapeProjection.OfficeArtShapeId]
                        .TextFonts = changedTextFormatting == null
                            ? null
                            : fonts;
                    result[shapeProjection.OfficeArtShapeId]
                        .PictureBullets = changedTextFormatting == null
                            ? null
                            : pictureBullets;
                    result[shapeProjection.OfficeArtShapeId]
                        .PictureFormatting = changedPictureFormatting;
                }
            }
            return true;
        }

        private static bool TryBuildModifiedSpecialMasterPersistObjects(
            PowerPointPresentation presentation, LegacyPptPackage package,
            LegacyPptProjectionMap projectionMap,
            IDictionary<uint, byte[]> rewritten,
            LegacyPptWriter.LegacyPptWriterFontCatalog fonts,
            LegacyPptWriter.LegacyPptWriterPictureBulletCatalog
                pictureBullets,
            PreservingPictureStoreUpdate pictureStore) {
            PresentationPart? presentationPart = presentation.OpenXmlDocument
                .PresentationPart;
            int processed = 0;
            if (presentationPart?.NotesMasterPart is NotesMasterPart notesPart) {
                if (projectionMap.TryGetSpecialMaster(notesPart,
                        out LegacyPptMasterProjection? projection)
                    && projection != null) {
                    IReadOnlyList<PowerPointShape> shapes = LegacyPptWriter
                        .ReadMasterShapesForWrite(notesPart,
                            out string? unsupportedReason);
                    if (unsupportedReason != null
                        || !TryRewriteSpecialMaster(package, projection, shapes,
                            LegacyPptWriter.LegacyPptWriterShapeContext.NotesMaster,
                            !projection.ThemeMatches(notesPart),
                            !projection.BackgroundMatches(notesPart),
                            () => LegacyPptWriter.TryReadBackground(notesPart,
                                out LegacyPptWriter.LegacyPptWriterBackground?
                                    background, out _)
                                ? background
                                : null,
                            record => LegacyPptWriter.BuildPreservedMasterThemeRecord(
                                record, notesPart,
                                projection.GetChangedClassicColorSlots(notesPart)),
                            rewritten, fonts, pictureBullets, pictureStore,
                            notesPart)) {
                        return false;
                    }
                    processed++;
                }
            }
            if (presentationPart?.HandoutMasterPart
                    is HandoutMasterPart handoutPart) {
                if (projectionMap.TryGetSpecialMaster(handoutPart,
                        out LegacyPptMasterProjection? projection)
                    && projection != null) {
                    IReadOnlyList<PowerPointShape> shapes = LegacyPptWriter
                        .ReadMasterShapesForWrite(handoutPart,
                            out string? unsupportedReason);
                    if (unsupportedReason != null
                        || !TryRewriteSpecialMaster(package, projection, shapes,
                            LegacyPptWriter.LegacyPptWriterShapeContext.HandoutMaster,
                            !projection.ThemeMatches(handoutPart),
                            !projection.BackgroundMatches(handoutPart),
                            () => LegacyPptWriter.TryReadBackground(handoutPart,
                                out LegacyPptWriter.LegacyPptWriterBackground?
                                    background, out _)
                                ? background
                                : null,
                            record => LegacyPptWriter.BuildPreservedMasterThemeRecord(
                                record, handoutPart,
                                projection.GetChangedClassicColorSlots(handoutPart)),
                            rewritten, fonts, pictureBullets, pictureStore,
                            handoutPart)) {
                        return false;
                    }
                    processed++;
                }
            }
            return processed == projectionMap.SpecialMasters.Count;
        }

        private static bool TryBuildModifiedTitleMasterPersistObjects(
            PowerPointPresentation presentation, LegacyPptPackage package,
            LegacyPptProjectionMap projectionMap,
            IDictionary<uint, byte[]> rewritten,
            LegacyPptWriter.LegacyPptWriterFontCatalog fonts,
            LegacyPptWriter.LegacyPptWriterPictureBulletCatalog
                pictureBullets,
            PreservingPictureStoreUpdate pictureStore) {
            SlideLayoutPart[] layouts = presentation.OpenXmlDocument
                .PresentationPart?.SlideMasterParts
                .SelectMany(master => master.SlideLayoutParts).ToArray()
                ?? Array.Empty<SlideLayoutPart>();
            int processed = 0;
            foreach (SlideLayoutPart part in layouts) {
                if (!projectionMap.TryGetTitleMaster(part,
                        out LegacyPptMasterProjection? projection)
                    || projection == null) continue;
                IReadOnlyList<PowerPointShape> shapes = LegacyPptWriter
                    .ReadMasterShapesForWrite(part,
                        out string? unsupportedReason);
                if (unsupportedReason != null
                    || !TryRewriteSpecialMaster(package, projection, shapes,
                        LegacyPptWriter.LegacyPptWriterShapeContext.Slide,
                        !projection.ThemeMatches(part),
                        !projection.BackgroundMatches(part),
                        () => LegacyPptWriter.TryReadBackground(part,
                            out LegacyPptWriter.LegacyPptWriterBackground?
                                background, out _)
                            ? background
                            : null,
                        record => LegacyPptWriter.BuildPreservedMasterThemeRecord(
                            record, part,
                            projection.GetChangedClassicColorSlots(part)),
                        rewritten, fonts, pictureBullets, pictureStore, part,
                        masterObjectsChanged:
                            !projection.MasterObjectsMatch(part),
                        followsMasterObjects: part.SlideLayout?
                            .ShowMasterShapes?.Value != false)) {
                    return false;
                }
                processed++;
            }
            return processed == projectionMap.TitleMasters.Count;
        }

        private static bool TryRewriteSpecialMaster(LegacyPptPackage package,
            LegacyPptMasterProjection projection,
            IReadOnlyList<PowerPointShape> shapes,
            LegacyPptWriter.LegacyPptWriterShapeContext shapeContext,
            bool themeChanged,
            bool backgroundChanged,
            Func<LegacyPptWriter.LegacyPptWriterBackground?> readBackground,
            Func<LegacyPptRecord, byte[]> rewriteTheme,
            IDictionary<uint, byte[]> rewritten,
            LegacyPptWriter.LegacyPptWriterFontCatalog fonts,
            LegacyPptWriter.LegacyPptWriterPictureBulletCatalog
                pictureBullets,
            PreservingPictureStoreUpdate pictureStore,
            OpenXmlPart ownerPart,
            bool masterObjectsChanged = false,
            bool followsMasterObjects = true) {
            if (!TryBuildMasterShapeEdits(shapes, projection, shapeContext,
                    fonts, pictureBullets, ownerPart,
                    out IReadOnlyDictionary<uint, ProjectedShapeEdit>
                        shapeEdits)) {
                return false;
            }
            LegacyPptWriter.LegacyPptWriterBackground? background = null;
            if (backgroundChanged && (background = readBackground()) == null) {
                return false;
            }
            if (!themeChanged && !backgroundChanged
                && !masterObjectsChanged
                && shapeEdits.Count == 0) return true;
            if (!package.PersistObjects.TryGetValue(projection.PersistId,
                    out LegacyPptPersistObject? persistObject)
                || persistObject == null) {
                return false;
            }
            LegacyPptRecord masterRecord = LegacyPptRecordReader.ReadSingle(
                persistObject.RecordBytes, 0, new LegacyPptImportOptions());
            byte[] bytes = masterRecord.CopyRecordBytes();
            if (shapeEdits.Count > 0) {
                RecordRewrite shapeRewrite = RewriteRecord(masterRecord,
                    shapeEdits);
                if (!shapeRewrite.Changed
                    || shapeRewrite.PatchedShapeCount != shapeEdits.Count) {
                    return false;
                }
                bytes = shapeRewrite.Bytes;
                if (shapeEdits.Values.Any(edit =>
                        edit.RewritePlaceholder)) {
                    LegacyPptRecord placeholderRecord =
                        LegacyPptRecordReader.ReadSingle(bytes, 0,
                            new LegacyPptImportOptions());
                    bytes = LegacyPptWriter
                        .BuildPreservedPlaceholderSignatureRecord(
                            placeholderRecord, shapes, shapeContext);
                }
            }
            if (backgroundChanged) {
                LegacyPptRecord backgroundRecord = LegacyPptRecordReader
                    .ReadSingle(bytes, 0, new LegacyPptImportOptions());
                if (!pictureStore.TryPrepareBackground(backgroundRecord,
                        background!)) return false;
                bytes = LegacyPptWriter.BuildPreservedBackgroundRecord(
                    backgroundRecord, background!, pictureStore.Catalog);
            }
            if (themeChanged) {
                LegacyPptRecord themedRecord = LegacyPptRecordReader.ReadSingle(
                    bytes, 0, new LegacyPptImportOptions());
                bytes = rewriteTheme(themedRecord);
            }
            if (masterObjectsChanged) {
                LegacyPptRecord inheritanceRecord = LegacyPptRecordReader
                    .ReadSingle(bytes, 0, new LegacyPptImportOptions());
                bytes = LegacyPptWriter
                    .BuildPreservedMasterObjectInheritanceRecord(
                        inheritanceRecord, followsMasterObjects);
            }
            rewritten.Add(projection.PersistId, bytes);
            return true;
        }
    }
}
