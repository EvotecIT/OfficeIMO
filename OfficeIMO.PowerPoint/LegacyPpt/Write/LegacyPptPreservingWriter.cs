using OfficeIMO.Drawing.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using System.Text;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint.LegacyPpt.Write {
    /// <summary>
    /// Appends a binary PowerPoint incremental edit for changes that can be represented without rebuilding
    /// or discarding the source persist graph. The original document stream remains an exact prefix.
    /// </summary>
    internal static partial class LegacyPptPreservingWriter {
        private const ushort RecordPersistDirectory = 0x1772;
        private const ushort RecordSlidePersistAtom = 0x03F3;
        private const ushort RecordSlideAtom = 0x03EF;
        private const ushort RecordSlideShowSlideInfoAtom = 0x03F9;
        private const ushort RecordSlideListWithText = 0x0FF0;
        private const ushort RecordTextHeader = 0x0F9F;
        private const ushort RecordTextChars = 0x0FA0;
        private const ushort RecordTextBytes = 0x0FA8;
        private const ushort RecordPlaceholder = 0x0BC3;
        private const ushort RecordHeadersFooters = 0x0FD9;
        private const ushort OfficeArtSpContainer = 0xF004;
        private const ushort OfficeArtDgg = 0xF006;
        private const ushort OfficeArtFsp = 0xF00A;
        private const ushort OfficeArtClientTextbox = 0xF00D;
        private const ushort OfficeArtChildAnchor = 0xF00F;
        private const ushort OfficeArtClientAnchor = 0xF010;

        internal static bool CanWritePresentation(PowerPointPresentation presentation) {
            if (presentation == null) throw new ArgumentNullException(nameof(presentation));
            return TryBuildModifiedPersistObjects(presentation, out _, out _);
        }

        internal static bool TryWritePresentation(PowerPointPresentation presentation, out byte[] bytes) {
            if (presentation == null) throw new ArgumentNullException(nameof(presentation));
            bytes = Array.Empty<byte>();
            if (!TryBuildModifiedPersistObjects(presentation,
                    out IReadOnlyDictionary<uint, byte[]> modifiedPersistObjects,
                    out IReadOnlyList<uint> currentSlideIds)) {
                return false;
            }

            LegacyPptPackage package = presentation.LegacyPptPackage!;
            if (modifiedPersistObjects.Count == 0) {
                bytes = package.CopyOriginalBytes();
                return true;
            }

            byte[] documentStream = AppendIncrementalEdit(package, modifiedPersistObjects, currentSlideIds,
                out uint editOffset);
            byte[] currentUserStream = PatchCurrentEditOffset(package.CurrentUserStream, editOffset);
            bytes = package.RewriteCompoundStreams(
                new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase) {
                    ["PowerPoint Document"] = documentStream,
                    ["Current User"] = currentUserStream
                });
            return true;
        }

        private static bool TryBuildModifiedPersistObjects(PowerPointPresentation presentation,
            out IReadOnlyDictionary<uint, byte[]> modifiedPersistObjects,
            out IReadOnlyList<uint> currentSlideIds) {
            var rewritten = new Dictionary<uint, byte[]>();
            var slideIds = new List<uint>(presentation.Slides.Count);
            modifiedPersistObjects = rewritten;
            currentSlideIds = slideIds;
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
            bool customShowsChanged = !CustomShowsEqual(projectionMap, customShows);
            if (customShowsChanged && !projectionMap.CanEditCustomShows) return false;

            try {
                if (!TryBuildModifiedMasterPersistObjects(presentation, package,
                        projectionMap, rewritten)) {
                    return false;
                }

                var currentSlideOrder = new List<LegacyPptSlideProjection>(presentation.Slides.Count);
                var addedSlides = new List<PowerPointSlide>();
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
                    } else if (!string.Equals(currentNotes,
                                   NormalizeLogicalText(slideProjection.Notes.Text),
                                   StringComparison.Ordinal)) {
                        if (!package.PersistObjects.TryGetValue(
                                slideProjection.Notes.PersistId,
                                out LegacyPptPersistObject? notesPersistObject)
                            || notesPersistObject == null) {
                            return false;
                        }
                        LegacyPptRecord notesRecord = LegacyPptRecordReader.ReadSingle(
                            notesPersistObject.RecordBytes, 0, new LegacyPptImportOptions());
                        if (!TryRewriteNotesRecord(notesRecord,
                                slideProjection.Notes.Text, currentNotes,
                                out byte[] rewrittenNotes)) {
                            return false;
                        }
                        rewritten.Add(slideProjection.Notes.PersistId, rewrittenNotes);
                    }

                    PowerPointShape[] shapes = slide.Shapes.ToArray();
                    if (shapes.Length != slideProjection.Shapes.Count) return false;
                    var editsByOfficeArtId = new Dictionary<uint, ProjectedShapeEdit>();
                    foreach (PowerPointShape shape in shapes) {
                        uint? openXmlShapeId = shape.Id;
                        if (!openXmlShapeId.HasValue
                            || !slideProjection.TryGetShape(openXmlShapeId.Value,
                                out LegacyPptShapeProjection? shapeProjection)
                            || shapeProjection == null
                            || !MatchesProjectedKind(shape, shapeProjection.Kind)) {
                            return false;
                        }
                        LegacyPptBounds bounds = GetBounds(shape);
                        LegacyPptBounds? changedBounds = BoundsEqual(bounds, shapeProjection.Bounds)
                            ? null
                            : bounds;
                        string? changedText = null;
                        if (shape is PowerPointTextBox textBox) {
                            if (!MatchesProjectedTextFormatting(textBox, shapeProjection)) return false;
                            string currentText = NormalizeLogicalText(textBox.Text);
                            if (!string.Equals(currentText, NormalizeLogicalText(shapeProjection.Text),
                                    StringComparison.Ordinal)) {
                                changedText = currentText;
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
                            || interactionEdit != null || animationChanged) {
                            editsByOfficeArtId.Add(shapeProjection.OfficeArtShapeId,
                                new ProjectedShapeEdit(changedBounds, shapeProjection.Text,
                                    changedText, interactionEdit,
                                    animationChanged, currentAnimation));
                        }
                    }
                    bool? hidden = slide.Hidden == slideProjection.Hidden ? null : slide.Hidden;
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
                    if (editsByOfficeArtId.Count == 0 && !hidden.HasValue
                        && !headerFooterChanged && !transitionChanged && !commentsChanged) continue;

                    LegacyPptRecord slideRecord = LegacyPptRecordReader.ReadSingle(persistObject.RecordBytes, 0,
                        new LegacyPptImportOptions());
                    if (!TryRewriteSlide(slide, slideRecord, editsByOfficeArtId, hidden,
                            transitionChanged, currentTransition,
                            soundCatalog,
                            headerFooterChanged, currentHeaderFooter,
                            commentsChanged, currentComments,
                            out RecordRewrite result)
                        || !result.Changed || result.PatchedShapeCount != editsByOfficeArtId.Count) return false;
                    rewritten.Add(slideProjection.PersistId, result.Bytes);
                }
                bool originalTopologyChanged = !currentSlideOrder.Select(slide => slide.PersistId)
                    .SequenceEqual(projectionMap.Slides.Select(slide => slide.PersistId));
                if (addedSlides.Count > 0) {
                    if (originalTopologyChanged || currentSlideOrder.Count != projectionMap.Slides.Count
                        || !TryAppendNewSlides(package, projectionMap, addedSlides, rewritten,
                            interactionCatalog, interactionContext,
                            out IReadOnlyList<uint> addedSlideIds)) {
                        return false;
                    }
                    slideIds.AddRange(addedSlideIds);
                } else if (originalTopologyChanged) {
                    if (!TryRewriteDocumentSlideOrder(package, projectionMap, currentSlideOrder,
                            out byte[] documentRecord)) {
                        return false;
                    }
                    rewritten.Add(package.DocumentPersistId, documentRecord);
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
                return true;
            } catch (Exception exception) when (exception is InvalidDataException
                                                || exception is OverflowException
                                                || exception is ArgumentException) {
                rewritten.Clear();
                return false;
            }
        }

        private static RecordRewrite RewriteRecord(LegacyPptRecord record,
            IReadOnlyDictionary<uint, ProjectedShapeEdit> editsByOfficeArtId) {
            if (record.Type == OfficeArtSpContainer) {
                LegacyPptRecord? fsp = record.Children.FirstOrDefault(child => child.Type == OfficeArtFsp);
                if (fsp != null && fsp.PayloadLength >= 4
                    && editsByOfficeArtId.TryGetValue(fsp.ReadUInt32(0), out ProjectedShapeEdit? edit)
                    && edit != null) {
                    return TryRewriteShapeContainer(record, edit, out byte[] rewrittenShape)
                        ? new RecordRewrite(rewrittenShape, changed: true, patchedShapeCount: 1)
                        : new RecordRewrite(record.CopyRecordBytes(), changed: false, patchedShapeCount: 0);
                }
            }
            if (record.Version != 0x0F || record.Children.Count == 0) {
                return new RecordRewrite(record.CopyRecordBytes(), changed: false, patchedShapeCount: 0);
            }

            var children = new List<byte[]>(record.Children.Count);
            bool changed = false;
            int patchedShapeCount = 0;
            foreach (LegacyPptRecord child in record.Children) {
                RecordRewrite childResult = RewriteRecord(child, editsByOfficeArtId);
                children.Add(childResult.Bytes);
                changed |= childResult.Changed;
                patchedShapeCount = checked(patchedShapeCount + childResult.PatchedShapeCount);
            }
            return changed
                ? new RecordRewrite(BuildRecord(record.Version, record.Instance, record.Type, Concat(children)),
                    changed: true, patchedShapeCount)
                : new RecordRewrite(record.CopyRecordBytes(), changed: false, patchedShapeCount: 0);
        }

        private static bool TryRewriteShapeContainer(LegacyPptRecord shapeContainer, ProjectedShapeEdit edit,
            out byte[] bytes) {
            var children = new List<byte[]>(shapeContainer.Children.Count + 1);
            bool patchedAnchor = !edit.Bounds.HasValue;
            bool patchedText = edit.Text == null
                && edit.Interactions?.RewriteTextInteractions != true;
            bool patchedShapeInteractions = edit.Interactions?.RewriteShapeInteractions != true;
            bool patchedAnimation = !edit.RewriteAnimation;
            bool appendedShapeInteractions = false;
            bool appendedAnimation = false;
            bool sawClientData = false;
            foreach (LegacyPptRecord child in shapeContainer.Children) {
                if (!patchedAnchor && edit.Bounds.HasValue
                    && (child.Type == OfficeArtClientAnchor || child.Type == OfficeArtChildAnchor)) {
                    children.Add(BuildAnchor(child.Type, child.Instance, edit.Bounds.Value));
                    patchedAnchor = true;
                } else if (!patchedText && child.Type == OfficeArtClientTextbox) {
                    bool rewritten = edit.Interactions?.RewriteTextInteractions == true
                        ? TryRewriteTextInteractions(child, edit.OriginalText,
                            edit.Text, edit.Interactions.Interactions.TextInteractions,
                            out byte[] textbox)
                        : TryRewriteTextBox(child, edit.OriginalText, edit.Text!,
                            out textbox);
                    if (!rewritten) {
                        bytes = shapeContainer.CopyRecordBytes();
                        return false;
                    }
                    children.Add(textbox);
                    patchedText = true;
                } else if (child.Type == OfficeArtClientData
                           && (edit.Interactions?.RewriteShapeInteractions == true
                               || edit.RewriteAnimation)) {
                    sawClientData = true;
                    byte[] clientData = child.CopyRecordBytes();
                    if (edit.Interactions?.RewriteShapeInteractions == true
                        && !TryRewriteClientDataInteractions(child,
                            edit.Interactions.Interactions.ShapeInteractions,
                            append: !appendedShapeInteractions,
                            out clientData)) {
                        bytes = shapeContainer.CopyRecordBytes();
                        return false;
                    }
                    LegacyPptRecord rewrittenClientData =
                        LegacyPptRecordReader.ReadSingle(clientData, 0,
                            new LegacyPptImportOptions());
                    if (edit.RewriteAnimation
                        && !TryRewriteClientDataAnimation(rewrittenClientData,
                            append: !appendedAnimation ? edit.Animation : null,
                            out clientData)) {
                        bytes = shapeContainer.CopyRecordBytes();
                        return false;
                    }
                    children.Add(clientData);
                    if (edit.Interactions?.RewriteShapeInteractions == true) {
                        appendedShapeInteractions = true;
                        patchedShapeInteractions = true;
                    }
                    if (edit.RewriteAnimation) {
                        appendedAnimation |= edit.Animation != null;
                        patchedAnimation = true;
                    }
                } else {
                    children.Add(child.CopyRecordBytes());
                }
            }
            if (!sawClientData
                && (edit.Interactions?.RewriteShapeInteractions == true
                    || edit.RewriteAnimation)) {
                var clientChildren = new List<byte[]>();
                if (edit.Animation != null) {
                    clientChildren.Add(LegacyPptWriter.BuildAnimationInfoRecord(
                        edit.Animation));
                }
                if (edit.Interactions?.RewriteShapeInteractions == true) {
                    clientChildren.AddRange(edit.Interactions.Interactions
                        .ShapeInteractions.Select(
                            LegacyPptWriter.BuildInteractiveInfoRecord));
                    patchedShapeInteractions = true;
                }
                patchedAnimation = true;
                if (clientChildren.Count > 0) {
                    children.Add(BuildRecord(version: 0x0F, instance: 0,
                        OfficeArtClientData, Concat(clientChildren)));
                }
            }
            if (!patchedAnchor || !patchedText || !patchedShapeInteractions
                || !patchedAnimation) {
                bytes = shapeContainer.CopyRecordBytes();
                return false;
            }
            bytes = BuildRecord(shapeContainer.Version, shapeContainer.Instance,
                shapeContainer.Type, Concat(children));
            return true;
        }

        private static bool TryRewriteTextBox(LegacyPptRecord textbox, string originalText, string replacementText,
            out byte[] bytes) {
            LegacyPptRecord[] textRecords = textbox.DescendantsAndSelf().Where(record =>
                record.Type == RecordTextChars || record.Type == RecordTextBytes).ToArray();
            if (textRecords.Length != 1
                || !TryBuildTextRecord(textbox, textRecords[0], originalText, replacementText,
                    out byte[] replacementRecord)) {
                bytes = textbox.CopyRecordBytes();
                return false;
            }
            return TryReplaceDescendant(textbox, textRecords[0].Offset, replacementRecord, out bytes);
        }

        private static bool TryRewriteNotesRecord(LegacyPptRecord record,
            string originalText, string replacementText, out byte[] bytes) {
            if (record.Type == OfficeArtSpContainer && IsNotesBodyShape(record)) {
                var children = new List<byte[]>(record.Children.Count);
                bool replaced = false;
                foreach (LegacyPptRecord child in record.Children) {
                    if (!replaced && child.Type == OfficeArtClientTextbox
                        && TryRewriteTextBox(child, originalText, replacementText,
                            out byte[] textbox)) {
                        children.Add(textbox);
                        replaced = true;
                    } else {
                        children.Add(child.CopyRecordBytes());
                    }
                }
                bytes = replaced
                    ? BuildRecord(record.Version, record.Instance, record.Type,
                        Concat(children))
                    : record.CopyRecordBytes();
                return replaced;
            }
            if (record.Version != 0x0F || record.Children.Count == 0) {
                bytes = record.CopyRecordBytes();
                return false;
            }

            var rewrittenChildren = new List<byte[]>(record.Children.Count);
            bool changed = false;
            foreach (LegacyPptRecord child in record.Children) {
                if (!changed && TryRewriteNotesRecord(child, originalText, replacementText,
                        out byte[] rewrittenChild)) {
                    rewrittenChildren.Add(rewrittenChild);
                    changed = true;
                } else {
                    rewrittenChildren.Add(child.CopyRecordBytes());
                }
            }
            bytes = changed
                ? BuildRecord(record.Version, record.Instance, record.Type,
                    Concat(rewrittenChildren))
                : record.CopyRecordBytes();
            return changed;
        }

        private static bool IsNotesBodyShape(LegacyPptRecord shapeContainer) {
            LegacyPptRecord? placeholder = shapeContainer.DescendantsAndSelf()
                .FirstOrDefault(record => record.Type == RecordPlaceholder
                    && record.PayloadLength >= 5);
            return placeholder?.ReadByte(4) is 0x06 or 0x0C;
        }

        private static bool TryBuildTextRecord(LegacyPptRecord textbox, LegacyPptRecord textRecord,
            string originalText, string replacementText, out byte[] bytes) {
            string raw = textRecord.Type == RecordTextChars
                ? textRecord.ReadUtf16Text()
                : textRecord.ReadLowByteUnicodeText();
            int contentLength = raw.Length;
            while (contentLength > 0 && raw[contentLength - 1] == '\0') contentLength--;
            while (contentLength > 0 && raw[contentLength - 1] == '\r') contentLength--;
            string decodedOriginal = NormalizeLogicalText(raw.Substring(0, contentLength));
            if (!string.Equals(decodedOriginal, NormalizeLogicalText(originalText), StringComparison.Ordinal)) {
                bytes = textRecord.CopyRecordBytes();
                return false;
            }

            string normalizedReplacement = NormalizeLogicalText(replacementText);
            if (normalizedReplacement.Length != contentLength
                && !IsStructurallyPlainTextBox(textbox)
                && !IsStructurallyPlainInteractiveTextBox(textbox)) {
                bytes = textRecord.CopyRecordBytes();
                return false;
            }
            string binaryReplacement = normalizedReplacement.Replace("\n", "\r") + raw.Substring(contentLength);
            byte[] payload;
            if (textRecord.Type == RecordTextChars) {
                payload = Encoding.Unicode.GetBytes(binaryReplacement);
            } else {
                if (binaryReplacement.Any(character => character > byte.MaxValue)) {
                    bytes = textRecord.CopyRecordBytes();
                    return false;
                }
                payload = binaryReplacement.Select(character => unchecked((byte)character)).ToArray();
            }
            bytes = BuildRecord(textRecord.Version, textRecord.Instance, textRecord.Type, payload);
            return true;
        }

        private static bool IsStructurallyPlainTextBox(LegacyPptRecord textbox) => textbox.Children.All(child =>
            child.Type == RecordTextHeader || child.Type == RecordTextChars || child.Type == RecordTextBytes);

        private static bool IsStructurallyPlainInteractiveTextBox(
            LegacyPptRecord textbox) {
            for (int index = 0; index < textbox.Children.Count; index++) {
                LegacyPptRecord child = textbox.Children[index];
                if (child.Type == RecordTextHeader || child.Type == RecordTextChars
                    || child.Type == RecordTextBytes) continue;
                if (child.Type != RecordInteractiveInfo
                    || !IsRewritableInteractiveInfo(child)
                    || index + 1 >= textbox.Children.Count
                    || textbox.Children[index + 1].Type != RecordTextInteractiveInfoAtom
                    || textbox.Children[index + 1].Version != 0
                    || textbox.Children[index + 1].Instance != child.Instance
                    || textbox.Children[index + 1].PayloadLength != 8) return false;
                index++;
            }
            return true;
        }

        private static bool TryReplaceDescendant(LegacyPptRecord record, int targetOffset, byte[] replacement,
            out byte[] bytes) {
            if (record.Offset == targetOffset) {
                bytes = replacement;
                return true;
            }
            if (record.Version != 0x0F || record.Children.Count == 0) {
                bytes = record.CopyRecordBytes();
                return false;
            }
            var children = new List<byte[]>(record.Children.Count);
            bool changed = false;
            foreach (LegacyPptRecord child in record.Children) {
                if (!changed && TryReplaceDescendant(child, targetOffset, replacement, out byte[] rewrittenChild)) {
                    children.Add(rewrittenChild);
                    changed = true;
                } else {
                    children.Add(child.CopyRecordBytes());
                }
            }
            bytes = changed
                ? BuildRecord(record.Version, record.Instance, record.Type, Concat(children))
                : record.CopyRecordBytes();
            return changed;
        }

        private static byte[] AppendIncrementalEdit(LegacyPptPackage package,
            IReadOnlyDictionary<uint, byte[]> modifiedPersistObjects, IReadOnlyList<uint> currentSlideIds,
            out uint editOffset) {
            using var output = new MemoryStream();
            output.Write(package.DocumentStream, 0, package.DocumentStream.Length);
            var offsets = new SortedDictionary<uint, uint>();
            foreach (KeyValuePair<uint, byte[]> persistObject in modifiedPersistObjects.OrderBy(pair => pair.Key)) {
                offsets.Add(persistObject.Key, checked((uint)output.Position));
                output.Write(persistObject.Value, 0, persistObject.Value.Length);
            }

            uint directoryOffset = checked((uint)output.Position);
            byte[] directory = BuildPersistDirectory(offsets);
            output.Write(directory, 0, directory.Length);

            editOffset = checked((uint)output.Position);
            LegacyPptRecord currentEdit = LegacyPptRecordReader.ReadSingle(package.DocumentStream,
                checked((int)package.CurrentEditOffset), new LegacyPptImportOptions());
            byte[] edit = currentEdit.CopyRecordBytes();
            if (currentEdit.PayloadLength < 20) {
                throw new InvalidDataException("The current UserEditAtom is too short for an incremental edit.");
            }
            uint lastViewedSlideId = ReadUInt32(edit, 8);
            if (lastViewedSlideId != 0 && !currentSlideIds.Contains(lastViewedSlideId)) {
                WriteUInt32(edit, 8, currentSlideIds.Count == 0 ? 0U : currentSlideIds[currentSlideIds.Count - 1]);
            }
            WriteUInt32(edit, 16, package.CurrentEditOffset);
            WriteUInt32(edit, 20, directoryOffset);
            WriteUInt32(edit, 24, package.DocumentPersistId);
            if (currentEdit.PayloadLength >= 24 && offsets.Count > 0) {
                WriteUInt32(edit, 28, Math.Max(currentEdit.ReadUInt32(20), offsets.Keys.Max()));
            }
            output.Write(edit, 0, edit.Length);
            return output.ToArray();
        }

        private static byte[] BuildPersistDirectory(IReadOnlyDictionary<uint, uint> offsets) {
            var payload = new List<byte>();
            KeyValuePair<uint, uint>[] entries = offsets.OrderBy(pair => pair.Key).ToArray();
            for (int index = 0; index < entries.Length;) {
                int count = 1;
                while (index + count < entries.Length && count < 0x0FFF
                       && entries[index + count].Key == entries[index].Key + unchecked((uint)count)) {
                    count++;
                }
                AppendUInt32(payload, (unchecked((uint)count) << 20) | entries[index].Key);
                for (int item = 0; item < count; item++) AppendUInt32(payload, entries[index + item].Value);
                index += count;
            }
            return BuildRecord(version: 0, instance: 0, RecordPersistDirectory, payload.ToArray());
        }

        private static byte[] PatchCurrentEditOffset(byte[] currentUserStream, uint editOffset) {
            byte[] patched = (byte[])currentUserStream.Clone();
            _ = LegacyPptCurrentUserAtom.Read(patched);
            WriteUInt32(patched, 16, editOffset);
            return patched;
        }

        private static byte[] PatchHiddenState(byte[] slideShowInfo, bool hidden) {
            if (slideShowInfo.Length < 19) {
                throw new InvalidDataException("The slide-show information atom is too short for its flags.");
            }
            slideShowInfo[18] = hidden
                ? unchecked((byte)(slideShowInfo[18] | 0x04))
                : unchecked((byte)(slideShowInfo[18] & ~0x04));
            return slideShowInfo;
        }

        private static byte[] BuildSlideShowInfo(bool hidden) {
            var payload = new byte[16];
            payload[10] = hidden ? (byte)0x05 : (byte)0x01;
            return BuildRecord(version: 0, instance: 0, RecordSlideShowSlideInfoAtom, payload);
        }

        private static LegacyPptBounds GetBounds(PowerPointShape shape) {
            int left = ToMasterUnits(shape.Left);
            int top = ToMasterUnits(shape.Top);
            int width = Math.Max(0, ToMasterUnits(shape.Width));
            int height = Math.Max(0, ToMasterUnits(shape.Height));
            return new LegacyPptBounds(left, top, width, height);
        }

        private static byte[] BuildAnchor(ushort type, ushort instance, LegacyPptBounds bounds) {
            int right = checked(bounds.Left + bounds.Width);
            int bottom = checked(bounds.Top + bounds.Height);
            if (FitsInt16(bounds.Left) && FitsInt16(bounds.Top) && FitsInt16(right) && FitsInt16(bottom)) {
                var payload = new byte[8];
                WriteInt16(payload, 0, unchecked((short)bounds.Top));
                WriteInt16(payload, 2, unchecked((short)bounds.Left));
                WriteInt16(payload, 4, unchecked((short)right));
                WriteInt16(payload, 6, unchecked((short)bottom));
                return BuildRecord(version: 0, instance, type, payload);
            }
            var largePayload = new byte[16];
            WriteInt32(largePayload, 0, bounds.Top);
            WriteInt32(largePayload, 4, bounds.Left);
            WriteInt32(largePayload, 8, right);
            WriteInt32(largePayload, 12, bottom);
            return BuildRecord(version: 0, instance, type, largePayload);
        }

        private static bool MatchesProjectedKind(PowerPointShape shape, LegacyPptShapeKind kind) {
            if (kind == LegacyPptShapeKind.TextBox) {
                return shape is PowerPointTextBox;
            }
            if (kind == LegacyPptShapeKind.Picture) return shape is PowerPointPicture;
            if (kind == LegacyPptShapeKind.Connector) return shape is PowerPointConnectionShape;
            if (kind == LegacyPptShapeKind.Group) return shape is PowerPointGroupShape;
            if (shape is not PowerPointAutoShape autoShape) return false;
            if (kind == LegacyPptShapeKind.AutoShape) return autoShape.ShapeType.HasValue;
            if (kind == LegacyPptShapeKind.Rectangle) return autoShape.ShapeType == A.ShapeTypeValues.Rectangle;
            if (kind == LegacyPptShapeKind.Ellipse) return autoShape.ShapeType == A.ShapeTypeValues.Ellipse;
            return kind == LegacyPptShapeKind.Line && autoShape.ShapeType == A.ShapeTypeValues.Line;
        }

        private static bool HasOnlyPlainProjectedText(PowerPointTextBox textBox) {
            P.Shape? shape = textBox.Element as P.Shape;
            if (shape?.TextBody == null) return true;
            if (shape.TextBody.Descendants<A.Field>().Any() || shape.TextBody.Descendants<A.Break>().Any()) {
                return false;
            }
            return !shape.TextBody.Descendants<A.RunProperties>().Any(properties =>
                       properties.HasAttributes || properties.ChildElements.Any(child =>
                           child is not A.HyperlinkOnClick
                               and not A.HyperlinkOnMouseOver))
                && !shape.TextBody.Descendants<A.ParagraphProperties>().Any(properties =>
                    properties.HasAttributes || properties.HasChildren)
                && !shape.TextBody.Descendants<A.EndParagraphRunProperties>().Any(properties =>
                    properties.HasAttributes || properties.HasChildren);
        }

        private static bool MatchesProjectedTextFormatting(PowerPointTextBox textBox,
            LegacyPptShapeProjection projection) {
            if (projection.TextFormattingFingerprint == null) return HasOnlyPlainProjectedText(textBox);
            P.Shape? shape = textBox.Element as P.Shape;
            return string.Equals(projection.TextFormattingFingerprint,
                LegacyPptTextProjection.CreateFormattingFingerprint(shape?.TextBody),
                StringComparison.Ordinal);
        }

        private static bool BoundsEqual(LegacyPptBounds left, LegacyPptBounds right) =>
            left.Left == right.Left && left.Top == right.Top && left.Width == right.Width && left.Height == right.Height;

        private static string NormalizeLogicalText(string value) => (value ?? string.Empty)
            .Replace("\r\n", "\n").Replace("\r", "\n");

        private static int ToMasterUnits(long emus) => checked((int)Math.Round(
            emus / 1587.5d, MidpointRounding.AwayFromZero));

        private static bool FitsInt16(int value) => value >= short.MinValue && value <= short.MaxValue;

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
            var result = new byte[values.Sum(record => record.Length)];
            int offset = 0;
            foreach (byte[] record in values) {
                Buffer.BlockCopy(record, 0, result, offset, record.Length);
                offset += record.Length;
            }
            return result;
        }

        private static void AppendUInt32(ICollection<byte> bytes, uint value) {
            bytes.Add(unchecked((byte)value));
            bytes.Add(unchecked((byte)(value >> 8)));
            bytes.Add(unchecked((byte)(value >> 16)));
            bytes.Add(unchecked((byte)(value >> 24)));
        }

        private static void WriteInt16(byte[] bytes, int offset, short value) =>
            WriteUInt16(bytes, offset, unchecked((ushort)value));

        private static uint ReadUInt32(byte[] bytes, int offset) => unchecked((uint)(bytes[offset]
            | bytes[offset + 1] << 8 | bytes[offset + 2] << 16 | bytes[offset + 3] << 24));

        private static void WriteUInt16(byte[] bytes, int offset, ushort value) {
            bytes[offset] = unchecked((byte)value);
            bytes[offset + 1] = unchecked((byte)(value >> 8));
        }

        private static void WriteInt32(byte[] bytes, int offset, int value) =>
            WriteUInt32(bytes, offset, unchecked((uint)value));

        private static void WriteUInt32(byte[] bytes, int offset, uint value) {
            bytes[offset] = unchecked((byte)value);
            bytes[offset + 1] = unchecked((byte)(value >> 8));
            bytes[offset + 2] = unchecked((byte)(value >> 16));
            bytes[offset + 3] = unchecked((byte)(value >> 24));
        }

        private readonly struct RecordRewrite {
            internal RecordRewrite(byte[] bytes, bool changed, int patchedShapeCount) {
                Bytes = bytes;
                Changed = changed;
                PatchedShapeCount = patchedShapeCount;
            }

            internal byte[] Bytes { get; }

            internal bool Changed { get; }

            internal int PatchedShapeCount { get; }
        }

        private sealed class ProjectedShapeEdit {
            internal ProjectedShapeEdit(LegacyPptBounds? bounds, string originalText,
                string? text, ProjectedInteractionEdit? interactions,
                bool rewriteAnimation,
                LegacyPptWriter.LegacyPptWriterAnimation? animation) {
                Bounds = bounds;
                OriginalText = originalText ?? string.Empty;
                Text = text;
                Interactions = interactions;
                RewriteAnimation = rewriteAnimation;
                Animation = animation;
            }

            internal LegacyPptBounds? Bounds { get; }

            internal string OriginalText { get; }

            internal string? Text { get; }

            internal ProjectedInteractionEdit? Interactions { get; }

            internal bool RewriteAnimation { get; }

            internal LegacyPptWriter.LegacyPptWriterAnimation? Animation { get; }
        }
    }
}
