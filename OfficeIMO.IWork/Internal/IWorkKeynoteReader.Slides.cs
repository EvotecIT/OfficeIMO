using OfficeIMO.IWork.Internal;

namespace OfficeIMO.IWork;

internal static partial class IWorkKeynoteReader {
    private static IWorkKeynoteSlide? ReadSlide(IWorkSourceDocument source, IWorkObjectIndex index, IWorkArchiveRecord slide,
        int position, bool skipped, IWorkProjectionBudget projectionBudget,
        ref int materializedCellCount,
        List<IWorkDiagnostic> diagnostics,
        ref bool supportsEditableReconstruction, List<IWorkObjectIdentity> omittedUnits,
        IWorkSourceReferenceIssueCollector references) {
        IWorkWireMessage message;
        try {
            message = index.Message(slide);
        } catch (InvalidDataException exception) when (!IWorkProtobuf.IsLimitException(exception)) {
            references.Declarations.Record(slide, "$", null);
            supportsEditableReconstruction = false;
            diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                "IWORK_KEYNOTE_SLIDE_MALFORMED",
                "A Keynote slide record is malformed; editable reconstruction is incomplete.",
                slide.EntryPath, slide.Identifier));
            return null;
        }
        foreach (int field in new[] { 7, 42, 5, 6 }) {
            projectionBudget.AddDrawableReferences(IWorkProtobuf.CountFields(
                slide.Payload, field, projectionBudget.MaximumProtobufFieldCount));
        }
        if (message.FieldCount(5) > 1) {
            supportsEditableReconstruction = false;
            diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                "IWORK_KEYNOTE_DRAWABLE_UNSUPPORTED",
                "A Keynote slide declares more than one title placeholder; editable reconstruction is incomplete.",
                slide.EntryPath, slide.Identifier));
        }
        IWorkArchiveRecord? titlePlaceholder = index.Dereference(message, 5);
        if (message.FieldCount(6) > 1) {
            supportsEditableReconstruction = false;
            diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                "IWORK_KEYNOTE_DRAWABLE_UNSUPPORTED",
                "A Keynote slide declares more than one body placeholder; editable reconstruction is incomplete.",
                slide.EntryPath, slide.Identifier));
        }
        IWorkArchiveRecord? bodyPlaceholder = index.Dereference(message, 6);
        if (titlePlaceholder != null && bodyPlaceholder != null
            && titlePlaceholder.Identifier == bodyPlaceholder.Identifier) {
            supportsEditableReconstruction = false;
            diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                "IWORK_KEYNOTE_DRAWABLE_UNSUPPORTED",
                "A Keynote slide assigns one drawable to both title and body placeholder roles; editable reconstruction is incomplete.",
                slide.EntryPath, slide.Identifier));
        }
        var candidates = new List<IWorkArchiveRecord>();
        var candidateIdentifiers = new HashSet<ulong>();
        bool hasUnresolvedDrawable = false;
        bool hasDuplicateDrawableOccurrence = false;
        foreach (int field in new[] { 7, 42, 5, 6 }) {
            var fieldIdentifiers = new HashSet<ulong>();
            IReadOnlyList<IWorkArchiveRecord> fieldCandidates = references.ReadAll(
                slide, message, field, out int unresolvedDrawableCount);
            hasUnresolvedDrawable |= unresolvedDrawableCount > 0;
            foreach (IWorkArchiveRecord candidate in fieldCandidates) {
                if (!fieldIdentifiers.Add(candidate.Identifier)) {
                    hasDuplicateDrawableOccurrence = true;
                    continue;
                }
                if (candidateIdentifiers.Add(candidate.Identifier)) candidates.Add(candidate);
            }
        }
        if (hasUnresolvedDrawable) {
            supportsEditableReconstruction = false;
            diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                "IWORK_KEYNOTE_DRAWABLE_UNSUPPORTED",
                "A Keynote slide contains an unresolved drawable reference; editable reconstruction is incomplete.",
                slide.EntryPath, slide.Identifier));
        }
        if (hasDuplicateDrawableOccurrence) {
            supportsEditableReconstruction = false;
            diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                "IWORK_KEYNOTE_DUPLICATE_DRAWABLE",
                "A Keynote slide repeats a drawable within the same ordered drawable field; editable reconstruction is incomplete.",
                slide.EntryPath, slide.Identifier));
        }

        IWorkTextBox? title = null;
        var textBoxes = new List<IWorkTextBox>();
        var images = new List<IWorkImageAsset>();
        var tables = new List<IWorkTable>();
        var drawables = new List<IWorkKeynoteDrawable>();
        var textCache = new Dictionary<ulong, IWorkTextContent>();
        foreach (IWorkArchiveRecord drawable in candidates) {
            try {
                index.Message(drawable);
            } catch (InvalidDataException exception) when (!IWorkProtobuf.IsLimitException(exception)) {
                references.Declarations.Record(drawable, "$", null);
                MarkDrawableIncomplete(drawable, diagnostics, ref supportsEditableReconstruction);
                omittedUnits.Add(new IWorkObjectIdentity(drawable));
                continue;
            }
            if (drawable.MessageType == 6000) {
                projectionBudget.AddTable();
                IWorkTable? table = IWorkTableReader.Read(source, drawable, projectionBudget, references, diagnostics,
                    ref materializedCellCount, ref supportsEditableReconstruction);
                if (table != null) {
                    tables.Add(table);
                    drawables.Add(new IWorkKeynoteDrawable(table));
                } else omittedUnits.Add(new IWorkObjectIdentity(drawable));
                continue;
            }
            if (drawable.MessageType == 3005) {
                projectionBudget.AddImage();
                IWorkImageAsset? image = IWorkDrawingReader.ReadImage(source, drawable,
                    projectionBudget, out bool imageComplete);
                if (!imageComplete || image == null) {
                    supportsEditableReconstruction = false;
                    diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                        "IWORK_KEYNOTE_IMAGE_UNSUPPORTED",
                        "A Keynote slide image could not be resolved completely; editable reconstruction is incomplete.",
                        drawable.EntryPath, drawable.Identifier));
                    omittedUnits.Add(new IWorkObjectIdentity(drawable));
                } else {
                    projectionBudget.AddProjectedImageBytes(image.Length);
                    images.Add(image);
                    drawables.Add(new IWorkKeynoteDrawable(image));
                }
                continue;
            }
            if (drawable.MessageType is not PlaceholderArchive and not TextShapeArchive) {
                omittedUnits.Add(new IWorkObjectIdentity(drawable));
                supportsEditableReconstruction = false;
                if (!diagnostics.Any(diagnostic => diagnostic.Code == "IWORK_KEYNOTE_DRAWABLE_UNSUPPORTED")) {
                    diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                        "IWORK_KEYNOTE_DRAWABLE_UNSUPPORTED",
                        "A Keynote slide contains an unsupported drawable type; editable reconstruction is incomplete.",
                        drawable.EntryPath, drawable.Identifier));
                }
                continue;
            }
            IWorkArchiveRecord? storage = DrawableStorage(index, drawable, references, out bool storageComplete);
            if (!storageComplete || storage == null) {
                supportsEditableReconstruction = false;
                if (!diagnostics.Any(diagnostic => diagnostic.Code == "IWORK_KEYNOTE_DRAWABLE_UNSUPPORTED")) {
                    diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                        "IWORK_KEYNOTE_DRAWABLE_UNSUPPORTED",
                        "A Keynote drawable contains malformed or unresolved text storage; editable reconstruction is incomplete.",
                        drawable.EntryPath, drawable.Identifier));
                }
            }
            if (storage == null || storage.MessageType != TextStorageArchive) continue;
            IWorkTextContent text;
            if (textCache.TryGetValue(storage.Identifier, out IWorkTextContent? cached)) {
                text = cached;
                projectionBudget.AddTextContentUse(text, includeCharacters: true);
            } else {
                try {
                    index.Message(storage);
                } catch (InvalidDataException exception) when (!IWorkProtobuf.IsLimitException(exception)) {
                    references.Declarations.Record(storage, "$", null);
                    MarkTextIncomplete(storage, diagnostics, ref supportsEditableReconstruction);
                    omittedUnits.Add(new IWorkObjectIdentity(storage));
                    continue;
                }
                text = IWorkTextReader.Read(index, storage, projectionBudget, references);
                textCache.Add(storage.Identifier, text);
            }
            if (!text.IsComplete) MarkTextIncomplete(storage, diagnostics, ref supportsEditableReconstruction, text);
            if (!text.IsTextComplete && text.Paragraphs.Count == 0)
                omittedUnits.Add(new IWorkObjectIdentity(storage));
            IWorkWireMessage? drawableMessage = IWorkDrawingReader.DrawableMessage(index, drawable,
                out bool drawableComplete);
            if (!drawableComplete) {
                supportsEditableReconstruction = false;
                if (!diagnostics.Any(diagnostic => diagnostic.Code == "IWORK_KEYNOTE_DRAWABLE_UNSUPPORTED")) {
                    diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                        "IWORK_KEYNOTE_DRAWABLE_UNSUPPORTED",
                        "A Keynote drawable contains malformed geometry; editable reconstruction is incomplete.",
                        drawable.EntryPath, drawable.Identifier));
                }
            }
            bool geometryComplete = true;
            bool hasDeclaredGeometry = drawableMessage?.HasField(1) == true;
            IWorkGeometry? geometry = drawableMessage == null
                ? null
                : IWorkDrawingReader.ReadGeometry(drawableMessage, out geometryComplete,
                    requirePositiveSize: hasDeclaredGeometry);
            if (titlePlaceholder != null && drawable.Identifier == titlePlaceholder.Identifier && title == null) {
                if (IWorkObjectIndex.TryGetMessage(message, 11, out bool malformedTitleGeometry)
                    is IWorkWireMessage titleGeometry) {
                    IWorkGeometry? placeholderGeometry = IWorkDrawingReader.ReadGeometryArchive(
                        titleGeometry, out bool titleGeometryComplete,
                        requirePositiveSize: true);
                    if (titleGeometryComplete) geometry = placeholderGeometry;
                    else MarkDrawableIncomplete(drawable, diagnostics, ref supportsEditableReconstruction);
                } else if (malformedTitleGeometry) {
                    MarkDrawableIncomplete(drawable, diagnostics, ref supportsEditableReconstruction);
                }
                bool metadataComplete = true;
                string? hyperlink = IWorkDrawingReader.ReadOptionalString(drawableMessage, 4,
                    projectionBudget, ref metadataComplete);
                string? accessibilityDescription = IWorkDrawingReader.ReadOptionalString(drawableMessage, 8,
                    projectionBudget, ref metadataComplete);
                if (!metadataComplete) {
                    MarkTextMetadataIncomplete(drawable, diagnostics, ref supportsEditableReconstruction);
                }
                if (text.PlainText.Length == 0 && hyperlink == null
                    && accessibilityDescription == null) continue;
                if (!geometryComplete) {
                    MarkDrawableIncomplete(drawable, diagnostics, ref supportsEditableReconstruction);
                }
                if (text.Paragraphs.Count == 0) projectionBudget.AddTextItem();
                title = new IWorkTextBox(text, geometry, hyperlink, accessibilityDescription, new IWorkObjectIdentity(drawable));
                drawables.Add(new IWorkKeynoteDrawable(title, isTitlePlaceholder: true));
            } else {
                bool isBodyPlaceholder = bodyPlaceholder?.Identifier == drawable.Identifier;
                if (isBodyPlaceholder) {
                    IWorkWireMessage? bodyGeometry = IWorkObjectIndex.TryGetMessage(
                        message, 14, out bool malformedBodyGeometry);
                    if (bodyGeometry != null) {
                        IWorkGeometry? placeholderGeometry = IWorkDrawingReader.ReadGeometryArchive(
                            bodyGeometry, out bool bodyGeometryComplete,
                            requirePositiveSize: true);
                        if (bodyGeometryComplete) geometry = placeholderGeometry;
                        else MarkDrawableIncomplete(drawable, diagnostics, ref supportsEditableReconstruction);
                    } else if (malformedBodyGeometry) {
                        MarkDrawableIncomplete(drawable, diagnostics, ref supportsEditableReconstruction);
                    }
                }
                bool metadataComplete = true;
                string? hyperlink = IWorkDrawingReader.ReadOptionalString(drawableMessage, 4,
                    projectionBudget, ref metadataComplete);
                string? accessibilityDescription = IWorkDrawingReader.ReadOptionalString(drawableMessage, 8,
                    projectionBudget, ref metadataComplete);
                if (!metadataComplete) {
                    MarkTextMetadataIncomplete(drawable, diagnostics, ref supportsEditableReconstruction);
                }
                if (text.PlainText.Length == 0 && hyperlink == null
                    && accessibilityDescription == null) continue;
                if (!geometryComplete || !isBodyPlaceholder
                    && !IWorkDrawingReader.HasPositiveSize(geometry)) {
                    MarkDrawableIncomplete(drawable, diagnostics, ref supportsEditableReconstruction);
                }
                if (text.Paragraphs.Count == 0) projectionBudget.AddTextItem();
                var textBox = new IWorkTextBox(text, geometry, hyperlink, accessibilityDescription, new IWorkObjectIdentity(drawable));
                textBoxes.Add(textBox);
                drawables.Add(new IWorkKeynoteDrawable(textBox, isTitlePlaceholder: false));
            }
        }

        IWorkTextContent notes = new(Array.Empty<IWorkTextParagraph>(),
            isComplete: true, isTextComplete: true);
        bool hasNoteReference = message.HasField(27);
        IReadOnlyList<IWorkArchiveRecord> noteRecords = references.ReadAll(
            slide, message, 27, out int unresolvedNoteCount);
        if (hasNoteReference && (unresolvedNoteCount > 0 || noteRecords.Count != 1)) {
            MarkNotesIncomplete(slide, diagnostics, ref supportsEditableReconstruction);
        } else if (noteRecords.Count == 1
                   && noteRecords[0].MessageType == PresenterNoteArchive) {
            IWorkArchiveRecord note = noteRecords[0];
            IWorkWireMessage? noteMessage = null;
            try {
                noteMessage = index.Message(note);
            } catch (InvalidDataException exception) when (!IWorkProtobuf.IsLimitException(exception)) {
                references.Declarations.Record(note, "$", null);
                noteMessage = null;
            }
            IReadOnlyList<IWorkArchiveRecord> noteStorages;
            int unresolvedStorageCount;
            if (noteMessage == null) {
                noteStorages = Array.Empty<IWorkArchiveRecord>();
                unresolvedStorageCount = 1;
            } else {
                noteStorages = references.ReadAll(note, noteMessage, 1, out unresolvedStorageCount);
            }
            if (unresolvedStorageCount == 0 && noteStorages.Count == 1
                && noteStorages[0].MessageType == TextStorageArchive) {
                IWorkArchiveRecord storage = noteStorages[0];
                bool storageMalformed = false;
                try {
                    index.Message(storage);
                } catch (InvalidDataException exception) when (!IWorkProtobuf.IsLimitException(exception)) {
                    references.Declarations.Record(storage, "$", null);
                    MarkTextIncomplete(storage, diagnostics, ref supportsEditableReconstruction);
                    storageMalformed = true;
                    omittedUnits.Add(new IWorkObjectIdentity(storage));
                }
                if (!storageMalformed) {
                    notes = IWorkTextReader.Read(index, storage, projectionBudget, references);
                    if (!notes.IsComplete) {
                        MarkTextIncomplete(storage, diagnostics, ref supportsEditableReconstruction, notes);
                    }
                }
            } else {
                foreach (IWorkArchiveRecord storage in noteStorages)
                    omittedUnits.Add(new IWorkObjectIdentity(storage));
                MarkNotesIncomplete(slide, diagnostics, ref supportsEditableReconstruction);
            }
        } else if (noteRecords.Count > 0) {
            MarkNotesIncomplete(slide, diagnostics, ref supportsEditableReconstruction);
        }
        string? slideName = message.GetString(10, out bool slideNameComplete);
        if (!slideNameComplete) {
            MarkTextMetadataIncomplete(slide, diagnostics, ref supportsEditableReconstruction);
        }
        if (slideName != null) projectionBudget.AddTextCharacters(slideName.Length);
        IEnumerable<IWorkTextContent> slideText =
            (title == null ? Array.Empty<IWorkTextContent>() : new[] { title.Content })
            .Concat(textBoxes.Select(textBox => textBox.Content))
            .Concat(tables.SelectMany(table => table.Cells)
                .Where(cell => cell.RichText != null).Select(cell => cell.RichText!))
            .Append(notes);
        if (slideText.SelectMany(content => content.Paragraphs).Any(paragraph =>
                paragraph.Style.PageBreakBefore == true
                || paragraph.Style.KeepWithNext == true
                || paragraph.Style.KeepLinesTogether == true)) {
            diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                "IWORK_KEYNOTE_PARAGRAPH_PAGINATION_UNSUPPORTED",
                "Keynote paragraph page/keep flags have no PPTX slide-text equivalent; editable text is preserved without those pagination flags.",
                slide.EntryPath, slide.Identifier));
        }
        return new IWorkKeynoteSlide(position, slideName ?? string.Empty,
            title, textBoxes, notes, images, tables, drawables, skipped, new IWorkObjectIdentity(slide));
    }

}
