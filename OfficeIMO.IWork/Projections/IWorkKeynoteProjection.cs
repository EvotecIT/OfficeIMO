using OfficeIMO.IWork.Internal;

namespace OfficeIMO.IWork;

/// <summary>One typed Keynote drawable retained in source stacking order.</summary>
public sealed class IWorkKeynoteDrawable {
    internal IWorkKeynoteDrawable(IWorkTextBox textBox, bool isTitlePlaceholder) {
        Kind = IWorkKeynoteDrawableKind.TextBox;
        TextBox = textBox;
        IsTitlePlaceholder = isTitlePlaceholder;
    }

    internal IWorkKeynoteDrawable(IWorkImageAsset image) {
        Kind = IWorkKeynoteDrawableKind.Image;
        Image = image;
    }

    internal IWorkKeynoteDrawable(IWorkTable table) {
        Kind = IWorkKeynoteDrawableKind.Table;
        Table = table;
    }

    /// <summary>Gets the drawable kind.</summary>
    public IWorkKeynoteDrawableKind Kind { get; }
    /// <summary>Gets the text-box payload when <see cref="Kind"/> is <see cref="IWorkKeynoteDrawableKind.TextBox"/>.</summary>
    public IWorkTextBox? TextBox { get; }
    /// <summary>Gets the image payload when <see cref="Kind"/> is <see cref="IWorkKeynoteDrawableKind.Image"/>.</summary>
    public IWorkImageAsset? Image { get; }
    /// <summary>Gets the table payload when <see cref="Kind"/> is <see cref="IWorkKeynoteDrawableKind.Table"/>.</summary>
    public IWorkTable? Table { get; }
    /// <summary>Gets whether this drawable is the slide title placeholder.</summary>
    public bool IsTitlePlaceholder { get; }
}

/// <summary>One Keynote slide recovered in presentation order.</summary>
public sealed class IWorkKeynoteSlide {
    internal IWorkKeynoteSlide(int index, string name, IWorkTextBox? titleBox,
        IReadOnlyList<IWorkTextBox> textBoxes, IWorkTextContent presenterNoteContent,
        IReadOnlyList<IWorkImageAsset> images, IReadOnlyList<IWorkTable> tables,
        IReadOnlyList<IWorkKeynoteDrawable> drawables, bool isSkipped) {
        Index = index;
        Name = name;
        TitleBox = titleBox;
        TextBoxes = Array.AsReadOnly(textBoxes.ToArray());
        PresenterNoteContent = presenterNoteContent;
        Images = Array.AsReadOnly(images.ToArray());
        Tables = Array.AsReadOnly(tables.ToArray());
        Drawables = Array.AsReadOnly(drawables.ToArray());
        Title = titleBox?.Content.PlainText ?? string.Empty;
        Body = Array.AsReadOnly(TextBoxes.Select(textBox => textBox.Content.PlainText)
            .Where(text => text.Length > 0).ToArray());
        PresenterNotes = presenterNoteContent.PlainText;
        IsSkipped = isSkipped;
    }

    /// <summary>Gets the one-based slide position.</summary>
    public int Index { get; }
    /// <summary>Gets the source slide name.</summary>
    public string Name { get; }
    /// <summary>Gets the positioned rich title placeholder.</summary>
    public IWorkTextBox? TitleBox { get; }
    /// <summary>Gets positioned rich body and freeform text boxes.</summary>
    public IReadOnlyList<IWorkTextBox> TextBoxes { get; }
    /// <summary>Gets rich presenter-note content.</summary>
    public IWorkTextContent PresenterNoteContent { get; }
    /// <summary>Gets embedded images in drawable order.</summary>
    public IReadOnlyList<IWorkImageAsset> Images { get; }
    /// <summary>Gets editable tables in drawable order.</summary>
    public IReadOnlyList<IWorkTable> Tables { get; }
    /// <summary>Gets text boxes, images, and tables in their shared source stacking order.</summary>
    public IReadOnlyList<IWorkKeynoteDrawable> Drawables { get; }
    /// <summary>Gets title-placeholder text.</summary>
    public string Title { get; }
    /// <summary>Gets remaining editable text blocks.</summary>
    public IReadOnlyList<string> Body { get; }
    /// <summary>Gets presenter-note text.</summary>
    public string PresenterNotes { get; }
    /// <summary>Gets whether the source slide is skipped in the show.</summary>
    public bool IsSkipped { get; }
}

/// <summary>Read-only Keynote structure recovered from a shared IWA object graph.</summary>
public sealed class IWorkKeynoteProjection {
    private readonly IWorkSourceDocument _source;
    private readonly bool _supportsEditableReconstruction;

    internal IWorkKeynoteProjection(IWorkSourceDocument source, IReadOnlyList<IWorkKeynoteSlide> slides,
        IWorkCanvasSize? slideSize,
        IReadOnlyList<IWorkDiagnostic> diagnostics, bool supportsEditableReconstruction) {
        _source = source;
        Slides = Array.AsReadOnly(slides.ToArray());
        Diagnostics = Array.AsReadOnly(diagnostics.ToArray());
        SlideSize = slideSize;
        _supportsEditableReconstruction = supportsEditableReconstruction;
    }

    /// <summary>Gets presented slides in source order.</summary>
    public IReadOnlyList<IWorkKeynoteSlide> Slides { get; }
    /// <summary>Gets the source presentation canvas size.</summary>
    public IWorkCanvasSize? SlideSize { get; }
    /// <summary>Gets projection diagnostics.</summary>
    public IReadOnlyList<IWorkDiagnostic> Diagnostics { get; }
    /// <summary>Gets whether at least one editable slide was recovered and all required slide references were resolved.</summary>
    public bool HasEditableContent => Slides.Count > 0 && _supportsEditableReconstruction;

    /// <summary>Creates an import report for an OfficeIMO semantic-owner projection.</summary>
    public IWorkImportReport CreateImportReport(IWorkProjectionKind kind, IWorkPreviewAsset? preview = null) =>
        CreateImportReport(kind, preview, Array.Empty<IWorkDiagnostic>());

    internal IWorkImportReport CreateImportReport(IWorkProjectionKind kind,
        IWorkPreviewAsset? preview, IReadOnlyList<IWorkDiagnostic> additionalDiagnostics) {
        ValidateReportRequest(kind, preview);
        return _source.CreateReport(kind, Diagnostics.Concat(additionalDiagnostics).ToArray(), preview,
            kind == IWorkProjectionKind.VisualFallback
                ? 0
                : Slides.Count + Slides.Sum(slide => slide.TextBoxes.Count + slide.Images.Count + slide.Tables.Count
                    + slide.Tables.Sum(table => table.Cells.Count)
                    + (slide.TitleBox != null ? 1 : 0) + (slide.PresenterNotes.Length > 0 ? 1 : 0)));
    }

    private void ValidateReportRequest(IWorkProjectionKind kind, IWorkPreviewAsset? preview) {
        if (kind == IWorkProjectionKind.EditableReconstruction && !HasEditableContent) {
            throw new InvalidOperationException("Editable Keynote content was not recovered.");
        }
        if (kind == IWorkProjectionKind.VisualFallback && preview == null) {
            throw new ArgumentNullException(nameof(preview), "A visual fallback report requires the preview used by the owner.");
        }
    }
}

public sealed partial class IWorkSourceDocument {
    /// <summary>Reads a Keynote package into a bounded semantic source projection, or returns a diagnostic-only projection in visual-only mode.</summary>
    public IWorkKeynoteProjection ReadKeynote() {
        if (Kind != IWorkDocumentKind.Keynote) throw new InvalidOperationException($"The source is {Kind}, not Keynote.");
        if (RequestedImportMode == IWorkImportMode.VisualOnly) {
            return new IWorkKeynoteProjection(this, Array.Empty<IWorkKeynoteSlide>(), null,
                new[] { IWorkProjectionDiagnostics.SemanticProjectionSkipped }, supportsEditableReconstruction: false);
        }
        return IWorkKeynoteReader.Read(this);
    }
}

internal static class IWorkKeynoteReader {
    private const uint DocumentArchive = 1;
    private const uint ShowArchive = 2;
    private const uint SlideNodeArchive = 4;
    private const uint SlideArchive = 5;
    private const uint PlaceholderArchive = 7;
    private const uint PresenterNoteArchive = 15;
    private const uint TextStorageArchive = 2001;
    private const uint TextShapeArchive = 2011;

    internal static IWorkKeynoteProjection Read(IWorkSourceDocument source) {
        var diagnostics = new List<IWorkDiagnostic>();
        var slides = new List<IWorkKeynoteSlide>();
        IWorkObjectIndex index = source.Index;
        IWorkArchiveRecord? document = index.UniqueOfType(DocumentArchive, out bool duplicateDocument);
        if (document == null) {
            diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                duplicateDocument ? "IWORK_KEYNOTE_DOCUMENT_DUPLICATE" : "IWORK_KEYNOTE_DOCUMENT_MISSING",
                duplicateDocument
                    ? "More than one Keynote document root was found; editable reconstruction is unavailable."
                    : "No supported Keynote document root was found; editable reconstruction is unavailable."));
            return new IWorkKeynoteProjection(source, slides, null, diagnostics, supportsEditableReconstruction: false);
        }
        IWorkWireMessage documentMessage = index.Message(document);
        bool showReferenceComplete = documentMessage.FieldCount(2) == 1
            && !documentMessage.HasUnexpectedWireKind(2, IWorkWireKind.Bytes);
        IWorkArchiveRecord? show = showReferenceComplete
            ? index.Dereference(documentMessage, 2)
            : null;
        if (!showReferenceComplete || show == null || show.MessageType != ShowArchive) {
            diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning, "IWORK_KEYNOTE_SHOW_MISSING",
                "The Keynote document root does not reference exactly one supported show object.", document.EntryPath, document.Identifier));
            return new IWorkKeynoteProjection(source, slides, null, diagnostics, supportsEditableReconstruction: false);
        }
        IWorkWireMessage showMessage = index.Message(show);
        byte[]? slideTreeBytes = showMessage.FieldCount(3) == 1
            ? showMessage.GetBytes(3)
            : null;
        int slideReferenceCount;
        int slideTreeFieldCount = 0;
        try {
            slideReferenceCount = slideTreeBytes == null
                || showMessage.HasUnexpectedWireKind(3, IWorkWireKind.Bytes)
                    ? -1
                    : IWorkProtobuf.CountFields(slideTreeBytes, 2,
                        source.Options.MaximumProtobufFieldCount,
                        out slideTreeFieldCount);
        } catch (InvalidDataException) {
            slideReferenceCount = -1;
        }
        if (slideReferenceCount < 0 || slideTreeFieldCount != slideReferenceCount) {
            diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning, "IWORK_KEYNOTE_SLIDE_TREE_MISSING",
                "The Keynote show does not contain a supported slide tree.", show.EntryPath, show.Identifier));
            return new IWorkKeynoteProjection(source, slides, null, diagnostics, supportsEditableReconstruction: false);
        }
        if (slideReferenceCount > source.Options.MaximumProjectedSlides) {
            throw new InvalidDataException($"Keynote slide count exceeds the configured projection limit of {source.Options.MaximumProjectedSlides}.");
        }
        IWorkWireMessage slideTree;
        try {
            slideTree = showMessage.ParseNestedMessage(slideTreeBytes!);
        } catch (InvalidDataException) {
            diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning, "IWORK_KEYNOTE_SLIDE_TREE_MISSING",
                "The Keynote show does not contain a supported slide tree.", show.EntryPath, show.Identifier));
            return new IWorkKeynoteProjection(source, slides, null, diagnostics, supportsEditableReconstruction: false);
        }

        bool supportsEditableReconstruction = true;
        IWorkCanvasSize? slideSize = ReadSlideSize(showMessage, out bool slideSizeComplete);
        if (!slideSizeComplete) {
            supportsEditableReconstruction = false;
            diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                "IWORK_KEYNOTE_SLIDE_SIZE_UNSUPPORTED",
                "The Keynote show declares an invalid slide size; editable reconstruction is incomplete.",
                show.EntryPath, show.Identifier));
        }
        int position = 0;
        int materializedCellCount = 0;
        var projectionBudget = new IWorkProjectionBudget(source.Options);
        IReadOnlyList<IWorkArchiveRecord> nodes = index.DereferenceAll(
            slideTree, 2, out int unresolvedNodeCount);
        if (nodes.Count > source.Options.MaximumProjectedSlides) {
            throw new InvalidDataException($"Keynote slide count exceeds the configured projection limit of {source.Options.MaximumProjectedSlides}.");
        }
        if (unresolvedNodeCount > 0) {
            supportsEditableReconstruction = false;
            diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                "IWORK_KEYNOTE_SLIDE_NODE_MISSING",
                "The Keynote slide tree references a missing node; editable reconstruction is incomplete.",
                show.EntryPath, show.Identifier));
        }
        var projectedNodeIdentifiers = new HashSet<ulong>();
        var projectedSlideIdentifiers = new HashSet<ulong>();
        foreach (IWorkArchiveRecord node in nodes) {
            position++;
            if (node.MessageType != SlideNodeArchive) {
                supportsEditableReconstruction = false;
                diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                    "IWORK_KEYNOTE_SLIDE_NODE_UNSUPPORTED",
                    "The Keynote slide tree references an unsupported node record; editable reconstruction is incomplete.",
                    node.EntryPath, node.Identifier));
                continue;
            }
            if (!projectedNodeIdentifiers.Add(node.Identifier)) {
                MarkDuplicateSlide(show, diagnostics, ref supportsEditableReconstruction);
                continue;
            }
            IWorkWireMessage nodeMessage = index.Message(node);
            ulong? skippedValue = nodeMessage.GetUnsigned(4);
            bool skipped = skippedValue == 1;
            if (nodeMessage.FieldCount(4) > 1
                || nodeMessage.HasUnexpectedWireKind(4, IWorkWireKind.Varint)
                || skippedValue > 1) {
                supportsEditableReconstruction = false;
                if (!diagnostics.Any(diagnostic =>
                        diagnostic.Code == "IWORK_KEYNOTE_SKIPPED_SLIDE_UNSUPPORTED")) {
                    diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                        "IWORK_KEYNOTE_SKIPPED_SLIDE_UNSUPPORTED",
                        "A Keynote slide-tree node declares an invalid skipped-slide flag; editable reconstruction is incomplete.",
                        node.EntryPath, node.Identifier));
                }
            }
            bool slideReferenceComplete = nodeMessage.FieldCount(2) == 1
                && !nodeMessage.HasUnexpectedWireKind(2, IWorkWireKind.Bytes);
            IWorkArchiveRecord? slide = slideReferenceComplete
                ? index.Dereference(nodeMessage, 2)
                : null;
            if (slide == null || slide.MessageType != SlideArchive) {
                supportsEditableReconstruction = false;
                diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                    "IWORK_KEYNOTE_SLIDE_MISSING",
                    "A Keynote slide-tree node references a missing or unsupported slide; editable reconstruction is incomplete.",
                    node.EntryPath, node.Identifier));
                continue;
            }
            if (!projectedSlideIdentifiers.Add(slide.Identifier)) {
                MarkDuplicateSlide(show, diagnostics, ref supportsEditableReconstruction);
                continue;
            }
            slides.Add(ReadSlide(source, index, slide, position, skipped, projectionBudget,
                ref materializedCellCount, diagnostics,
                ref supportsEditableReconstruction));
        }
        return new IWorkKeynoteProjection(source, slides, slideSize, diagnostics, supportsEditableReconstruction);
    }

    private static IWorkKeynoteSlide ReadSlide(IWorkSourceDocument source, IWorkObjectIndex index, IWorkArchiveRecord slide,
        int position, bool skipped, IWorkProjectionBudget projectionBudget,
        ref int materializedCellCount,
        List<IWorkDiagnostic> diagnostics,
        ref bool supportsEditableReconstruction) {
        foreach (int field in new[] { 7, 42, 5, 6 }) {
            projectionBudget.AddDrawableReferences(IWorkProtobuf.CountFields(
                slide.Payload, field, projectionBudget.MaximumProtobufFieldCount));
        }
        IWorkWireMessage message = index.Message(slide);
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
            IReadOnlyList<IWorkArchiveRecord> fieldCandidates = index.DereferenceAll(
                message, field, out int unresolvedDrawableCount);
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
            if (drawable.MessageType == 6000) {
                projectionBudget.AddTable();
                IWorkTable? table = IWorkTableReader.Read(source, drawable, projectionBudget, diagnostics,
                    ref materializedCellCount, ref supportsEditableReconstruction);
                if (table != null) {
                    tables.Add(table);
                    drawables.Add(new IWorkKeynoteDrawable(table));
                }
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
                } else {
                    projectionBudget.AddProjectedImageBytes(image.Length);
                    images.Add(image);
                    drawables.Add(new IWorkKeynoteDrawable(image));
                }
                continue;
            }
            if (drawable.MessageType is not PlaceholderArchive and not TextShapeArchive) {
                supportsEditableReconstruction = false;
                if (!diagnostics.Any(diagnostic => diagnostic.Code == "IWORK_KEYNOTE_DRAWABLE_UNSUPPORTED")) {
                    diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                        "IWORK_KEYNOTE_DRAWABLE_UNSUPPORTED",
                        "A Keynote slide contains an unsupported drawable type; editable reconstruction is incomplete.",
                        drawable.EntryPath, drawable.Identifier));
                }
                continue;
            }
            IWorkArchiveRecord? storage = DrawableStorage(index, drawable, out bool storageComplete);
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
                text = IWorkTextReader.Read(index, storage, projectionBudget);
                textCache.Add(storage.Identifier, text);
            }
            if (!text.IsComplete) MarkTextIncomplete(storage, diagnostics, ref supportsEditableReconstruction);
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
            IWorkGeometry? geometry = drawableMessage == null
                ? null
                : IWorkDrawingReader.ReadGeometry(drawableMessage, out geometryComplete,
                    requirePositiveSize: true);
            if (drawableMessage != null && !geometryComplete) {
                supportsEditableReconstruction = false;
                if (!diagnostics.Any(diagnostic => diagnostic.Code == "IWORK_KEYNOTE_DRAWABLE_UNSUPPORTED")) {
                    diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                        "IWORK_KEYNOTE_DRAWABLE_UNSUPPORTED",
                        "A Keynote drawable contains malformed geometry; editable reconstruction is incomplete.",
                        drawable.EntryPath, drawable.Identifier));
                }
            }
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
                if (text.Paragraphs.Count == 0) projectionBudget.AddTextItem();
                title = new IWorkTextBox(text, geometry, hyperlink, accessibilityDescription);
                drawables.Add(new IWorkKeynoteDrawable(title, isTitlePlaceholder: true));
            } else {
                if (bodyPlaceholder?.Identifier == drawable.Identifier) {
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
                if (text.Paragraphs.Count == 0) projectionBudget.AddTextItem();
                var textBox = new IWorkTextBox(text, geometry, hyperlink, accessibilityDescription);
                textBoxes.Add(textBox);
                drawables.Add(new IWorkKeynoteDrawable(textBox, isTitlePlaceholder: false));
            }
        }

        IWorkTextContent notes = new(Array.Empty<IWorkTextParagraph>(), isComplete: true);
        bool hasNoteReference = message.HasField(27);
        IReadOnlyList<IWorkArchiveRecord> noteRecords = index.DereferenceAll(
            message, 27, out int unresolvedNoteCount);
        if (hasNoteReference && (unresolvedNoteCount > 0 || noteRecords.Count != 1)) {
            MarkNotesIncomplete(slide, diagnostics, ref supportsEditableReconstruction);
        } else if (noteRecords.Count == 1
                   && noteRecords[0].MessageType == PresenterNoteArchive) {
            IWorkArchiveRecord note = noteRecords[0];
            IReadOnlyList<IWorkArchiveRecord> noteStorages = index.DereferenceAll(
                index.Message(note), 1, out int unresolvedStorageCount);
            if (unresolvedStorageCount == 0 && noteStorages.Count == 1
                && noteStorages[0].MessageType == TextStorageArchive) {
                IWorkArchiveRecord storage = noteStorages[0];
                notes = IWorkTextReader.Read(index, storage, projectionBudget);
                if (!notes.IsComplete) MarkTextIncomplete(storage, diagnostics, ref supportsEditableReconstruction);
            } else {
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
            title, textBoxes, notes, images, tables, drawables, skipped);
    }

    private static void MarkDrawableIncomplete(IWorkArchiveRecord drawable,
        List<IWorkDiagnostic> diagnostics, ref bool supportsEditableReconstruction) {
        supportsEditableReconstruction = false;
        if (diagnostics.Any(diagnostic => diagnostic.Code == "IWORK_KEYNOTE_DRAWABLE_UNSUPPORTED")) return;
        diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
            "IWORK_KEYNOTE_DRAWABLE_UNSUPPORTED",
            "A Keynote drawable contains malformed geometry; editable reconstruction is incomplete.",
            drawable.EntryPath, drawable.Identifier));
    }

    private static void MarkTextMetadataIncomplete(IWorkArchiveRecord record,
        List<IWorkDiagnostic> diagnostics, ref bool supportsEditableReconstruction) {
        supportsEditableReconstruction = false;
        if (diagnostics.Any(diagnostic => diagnostic.Code == "IWORK_KEYNOTE_TEXT_UNSUPPORTED")) return;
        diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
            "IWORK_KEYNOTE_TEXT_UNSUPPORTED",
            "Keynote text metadata contains invalid Unicode content; editable reconstruction is incomplete.",
            record.EntryPath, record.Identifier));
    }

    private static IWorkCanvasSize? ReadSlideSize(IWorkWireMessage show, out bool complete) {
        complete = true;
        if (!show.HasField(4)) return null;
        IWorkWireMessage? size = IWorkObjectIndex.TryGetMessage(show, 4, out bool malformedSize);
        if (show.HasUnexpectedWireKind(4, IWorkWireKind.Bytes) || malformedSize || size == null) {
            complete = false;
            return null;
        }
        IWorkWireMessage declaredSize = size;
        double width = declaredSize.GetFloat(1) ?? 0;
        double height = declaredSize.GetFloat(2) ?? 0;
        if (!declaredSize.HasField(1) || !declaredSize.HasField(2)
            || declaredSize.FieldCount(1) > 1 || declaredSize.FieldCount(2) > 1
            || declaredSize.HasUnexpectedWireKind(1, IWorkWireKind.Fixed32)
            || declaredSize.HasUnexpectedWireKind(2, IWorkWireKind.Fixed32)
            || !declaredSize.GetFloat(1).HasValue || !declaredSize.GetFloat(2).HasValue
            || width <= 0 || height <= 0 || double.IsNaN(width) || double.IsInfinity(width)
            || double.IsNaN(height) || double.IsInfinity(height)) {
            complete = false;
            return null;
        }
        return new IWorkCanvasSize(width, height);
    }

    private static void MarkNotesIncomplete(IWorkArchiveRecord slide,
        List<IWorkDiagnostic> diagnostics,
        ref bool supportsEditableReconstruction) {
        supportsEditableReconstruction = false;
        diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
            "IWORK_KEYNOTE_NOTES_UNSUPPORTED",
            "A Keynote slide contains an unresolved presenter-note reference; editable reconstruction is incomplete.",
            slide.EntryPath, slide.Identifier));
    }

    private static void MarkDuplicateSlide(IWorkArchiveRecord show,
        List<IWorkDiagnostic> diagnostics, ref bool supportsEditableReconstruction) {
        supportsEditableReconstruction = false;
        if (diagnostics.Any(diagnostic => diagnostic.Code == "IWORK_KEYNOTE_DUPLICATE_SLIDE")) return;
        diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
            "IWORK_KEYNOTE_DUPLICATE_SLIDE",
            "The Keynote slide tree repeats a node or slide; editable reconstruction is incomplete.",
            show.EntryPath, show.Identifier));
    }

    private static void MarkTextIncomplete(IWorkArchiveRecord storage,
        List<IWorkDiagnostic> diagnostics, ref bool supportsEditableReconstruction) {
        supportsEditableReconstruction = false;
        if (diagnostics.Any(diagnostic => diagnostic.Code == "IWORK_KEYNOTE_TEXT_STORAGE_UNSUPPORTED")) return;
        diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
            "IWORK_KEYNOTE_TEXT_STORAGE_UNSUPPORTED",
            "A Keynote text storage contains an invalid UTF-8 run; editable reconstruction is incomplete.",
            storage.EntryPath, storage.Identifier));
    }

    private static IWorkArchiveRecord? DrawableStorage(IWorkObjectIndex index, IWorkArchiveRecord drawable,
        out bool complete) {
        complete = true;
        IWorkWireMessage message = index.Message(drawable);
        if (drawable.MessageType == TextShapeArchive) {
            IWorkArchiveRecord? field4 = index.Dereference(message, 4);
            IWorkArchiveRecord? field2 = index.Dereference(message, 2);
            bool directAmbiguous = message.FieldCount(4) > 1
                || message.FieldCount(2) > 1
                || field4 != null && field2 != null && field4.Identifier != field2.Identifier;
            if (directAmbiguous
                || message.HasUnexpectedWireKind(4, IWorkWireKind.Bytes)
                || message.HasUnexpectedWireKind(2, IWorkWireKind.Bytes)
                || message.HasField(4) && (field4 == null || field4.MessageType != TextStorageArchive)
                || message.HasField(2) && (field2 == null || field2.MessageType != TextStorageArchive)) {
                complete = false;
            }
            IWorkArchiveRecord? direct = field4?.MessageType == TextStorageArchive ? field4
                : field2?.MessageType == TextStorageArchive ? field2
                : null;
            if (direct != null) return direct;
        }
        IWorkWireMessage? super = IWorkObjectIndex.TryGetMessage(message, 1, out bool malformedSuper);
        if (malformedSuper || message.HasUnexpectedWireKind(1, IWorkWireKind.Bytes)) complete = false;
        if (super == null) return null;
        IWorkArchiveRecord? nested = index.Dereference(super, 2);
        if (super.HasUnexpectedWireKind(2, IWorkWireKind.Bytes)
            || super.FieldCount(2) > 1
            || super.HasField(2) && (nested == null || nested.MessageType != TextStorageArchive)) {
            complete = false;
        }
        IWorkArchiveRecord? nestedStorage = nested?.MessageType == TextStorageArchive ? nested : null;
        return nestedStorage;
    }
}
