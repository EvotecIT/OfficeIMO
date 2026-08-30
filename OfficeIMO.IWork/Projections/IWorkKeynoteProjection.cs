using OfficeIMO.IWork.Internal;

namespace OfficeIMO.IWork;

/// <summary>One Keynote slide recovered in presentation order.</summary>
public sealed class IWorkKeynoteSlide {
    internal IWorkKeynoteSlide(int index, string name, IWorkTextBox? titleBox,
        IReadOnlyList<IWorkTextBox> textBoxes, IWorkTextContent presenterNoteContent,
        IReadOnlyList<IWorkImageAsset> images, IReadOnlyList<IWorkTable> tables, bool isSkipped) {
        Index = index;
        Name = name;
        TitleBox = titleBox;
        TextBoxes = Array.AsReadOnly(textBoxes.ToArray());
        PresenterNoteContent = presenterNoteContent;
        Images = Array.AsReadOnly(images.ToArray());
        Tables = Array.AsReadOnly(tables.ToArray());
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
    public IWorkImportReport CreateImportReport(IWorkProjectionKind kind, IWorkPreviewAsset? preview = null) {
        ValidateReportRequest(kind, preview);
        return _source.CreateReport(kind, Diagnostics, preview,
            kind == IWorkProjectionKind.VisualFallback
                ? 0
                : Slides.Count + Slides.Sum(slide => slide.Body.Count + slide.Images.Count + slide.Tables.Count
                    + slide.Tables.Sum(table => table.Cells.Count)
                    + (slide.Title.Length > 0 ? 1 : 0) + (slide.PresenterNotes.Length > 0 ? 1 : 0)));
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
    private const uint TextStorageArchive = 2001;

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
        IWorkArchiveRecord? show = index.Dereference(index.Message(document), 2);
        if (show == null || show.MessageType != ShowArchive) {
            diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning, "IWORK_KEYNOTE_SHOW_MISSING",
                "The Keynote document root does not reference a supported show object.", document.EntryPath, document.Identifier));
            return new IWorkKeynoteProjection(source, slides, null, diagnostics, supportsEditableReconstruction: false);
        }
        IWorkWireMessage? slideTree = IWorkObjectIndex.TryGetMessage(index.Message(show), 3);
        if (slideTree == null) {
            diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning, "IWORK_KEYNOTE_SLIDE_TREE_MISSING",
                "The Keynote show does not contain a supported slide tree.", show.EntryPath, show.Identifier));
            return new IWorkKeynoteProjection(source, slides, null, diagnostics, supportsEditableReconstruction: false);
        }

        bool supportsEditableReconstruction = true;
        IWorkCanvasSize? slideSize = ReadSlideSize(index.Message(show), out bool slideSizeComplete);
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
            if (nodeMessage.HasUnexpectedWireKind(4, IWorkWireKind.Varint)
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
            IWorkArchiveRecord? slide = index.Dereference(nodeMessage, 2);
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
        IWorkWireMessage message = index.Message(slide);
        IWorkArchiveRecord? titlePlaceholder = index.Dereference(message, 5);
        var candidates = new List<IWorkArchiveRecord>();
        var candidateIdentifiers = new HashSet<ulong>();
        bool hasUnresolvedDrawable = false;
        foreach (int field in new[] { 5, 6, 7, 42 }) {
            IReadOnlyList<IWorkArchiveRecord> fieldCandidates = index.DereferenceAll(
                message, field, out int unresolvedDrawableCount);
            hasUnresolvedDrawable |= unresolvedDrawableCount > 0;
            foreach (IWorkArchiveRecord candidate in fieldCandidates) {
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

        IWorkTextBox? title = null;
        var textBoxes = new List<IWorkTextBox>();
        var images = new List<IWorkImageAsset>();
        var tables = new List<IWorkTable>();
        var seenStorages = new HashSet<ulong>();
        foreach (IWorkArchiveRecord drawable in candidates) {
            if (drawable.MessageType == 6000) {
                projectionBudget.AddTable();
                IWorkTable? table = IWorkTableReader.Read(source, drawable, projectionBudget, diagnostics,
                    ref materializedCellCount, ref supportsEditableReconstruction);
                if (table != null) tables.Add(table);
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
                    images.Add(image);
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
            if (storage == null || storage.MessageType != TextStorageArchive || !seenStorages.Add(storage.Identifier)) continue;
            IWorkTextContent text = IWorkTextReader.Read(index, storage, projectionBudget);
            if (!text.IsComplete) MarkTextIncomplete(storage, diagnostics, ref supportsEditableReconstruction);
            if (text.PlainText.Length == 0) continue;
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
                : IWorkDrawingReader.ReadGeometry(drawableMessage, out geometryComplete);
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
                        titleGeometry, out bool titleGeometryComplete);
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
                title = new IWorkTextBox(text, geometry, hyperlink, accessibilityDescription);
            } else {
                if (index.Dereference(message, 6)?.Identifier == drawable.Identifier) {
                    IWorkWireMessage? bodyGeometry = IWorkObjectIndex.TryGetMessage(
                        message, 14, out bool malformedBodyGeometry);
                    if (bodyGeometry != null) {
                        IWorkGeometry? placeholderGeometry = IWorkDrawingReader.ReadGeometryArchive(
                            bodyGeometry, out bool bodyGeometryComplete);
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
                textBoxes.Add(new IWorkTextBox(text, geometry, hyperlink, accessibilityDescription));
            }
        }

        IWorkTextContent notes = new(Array.Empty<IWorkTextParagraph>(), isComplete: true);
        bool hasNoteReference = message.HasField(27);
        IWorkArchiveRecord? note = index.Dereference(message, 27);
        if (message.LacksWireKind(27, IWorkWireKind.Bytes)
            || hasNoteReference && note == null) {
            MarkNotesIncomplete(slide, diagnostics, ref supportsEditableReconstruction);
        } else if (note != null) {
            IWorkArchiveRecord? storage = index.Dereference(index.Message(note), 1);
            if (storage != null && storage.MessageType == TextStorageArchive) {
                notes = IWorkTextReader.Read(index, storage, projectionBudget);
                if (!notes.IsComplete) MarkTextIncomplete(storage, diagnostics, ref supportsEditableReconstruction);
            } else {
                MarkNotesIncomplete(slide, diagnostics, ref supportsEditableReconstruction);
            }
        }
        string? slideName = message.GetString(10, out bool slideNameComplete);
        if (!slideNameComplete) {
            MarkTextMetadataIncomplete(slide, diagnostics, ref supportsEditableReconstruction);
        }
        if (slideName != null) projectionBudget.AddTextCharacters(slideName.Length);
        return new IWorkKeynoteSlide(position, slideName ?? string.Empty,
            title, textBoxes, notes, images, tables, skipped);
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
        if (show.LacksWireKind(4, IWorkWireKind.Bytes) || malformedSize || size == null) {
            complete = false;
            return null;
        }
        IWorkWireMessage declaredSize = size;
        double width = declaredSize.GetFloat(1) ?? 0;
        double height = declaredSize.GetFloat(2) ?? 0;
        if (!declaredSize.HasField(1) || !declaredSize.HasField(2)
            || declaredSize.LacksWireKind(1, IWorkWireKind.Fixed32)
            || declaredSize.LacksWireKind(2, IWorkWireKind.Fixed32)
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
        IWorkArchiveRecord? field4 = index.Dereference(message, 4);
        IWorkArchiveRecord? field2 = index.Dereference(message, 2);
        IWorkArchiveRecord? direct = field4 ?? field2;
        if ((message.HasBytes(4) && field4 == null) || (message.HasBytes(2) && field2 == null)) complete = false;
        if (direct != null && direct.MessageType == TextStorageArchive) return direct;
        IWorkWireMessage? super = IWorkObjectIndex.TryGetMessage(message, 1, out bool malformedSuper);
        if (malformedSuper) complete = false;
        if (super == null) return null;
        IWorkArchiveRecord? nested = index.Dereference(super, 2);
        if (super.HasBytes(2) && nested == null) complete = false;
        return nested != null && nested.MessageType == TextStorageArchive ? nested : null;
    }
}
