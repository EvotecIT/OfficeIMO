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
        IReadOnlyList<IWorkKeynoteDrawable> drawables, bool isSkipped, IWorkObjectIdentity? sourceIdentity = null, IWorkCellFill? background = null) {
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
        SourceIdentity = sourceIdentity;
        HasBackgroundFill = background != null;
        BackgroundColor = background?.Color;
    }

    /// <summary>Gets the one-based slide position.</summary>
    public int Index { get; }
    /// <summary>Gets the native slide identity.</summary>
    public IWorkObjectIdentity? SourceIdentity { get; }
    /// <summary>Gets whether a supported explicit or inherited background fill was recovered.</summary>
    public bool HasBackgroundFill { get; }
    /// <summary>Gets the opaque background color, or null for no fill or an unrecovered background.
    /// Use <see cref="HasBackgroundFill"/> to distinguish a recovered no-fill declaration.</summary>
    public IWorkColor? BackgroundColor { get; }
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
public sealed partial class IWorkKeynoteProjection {
    private readonly IWorkSourceDocument _source;
    private readonly bool _supportsEditableReconstruction;

    internal IWorkKeynoteProjection(IWorkSourceDocument source, IReadOnlyList<IWorkKeynoteSlide> slides,
        IWorkCanvasSize? slideSize,
        IReadOnlyList<IWorkDiagnostic> diagnostics, bool supportsEditableReconstruction,
        IWorkObjectIdentity? sourceIdentity = null, IReadOnlyList<IWorkObjectIdentity>? omittedUnits = null,
        IReadOnlyList<IWorkSourceReferenceIssue>? referenceIssues = null,
        IReadOnlyList<IWorkSourceDeclarationIssue>? declarationIssues = null) {
        _source = source;
        SourceIdentity = sourceIdentity;
        OmittedSourceUnits = Array.AsReadOnly((omittedUnits ?? Array.Empty<IWorkObjectIdentity>()).ToArray());
        SourceReferenceIssues = Array.AsReadOnly((referenceIssues ?? Array.Empty<IWorkSourceReferenceIssue>()).ToArray());
        SourceDeclarationIssues = Array.AsReadOnly((declarationIssues ?? Array.Empty<IWorkSourceDeclarationIssue>()).ToArray());
        Slides = Array.AsReadOnly(slides.ToArray());
        Diagnostics = Array.AsReadOnly(diagnostics.ToArray());
        SlideSize = slideSize;
        _supportsEditableReconstruction = supportsEditableReconstruction;
    }

    /// <summary>Gets presented slides in source order.</summary>
    public IReadOnlyList<IWorkKeynoteSlide> Slides { get; }
    /// <summary>Gets the native document-root identity when exactly one root was identified.</summary>
    public IWorkObjectIdentity? SourceIdentity { get; }
    private IReadOnlyList<IWorkObjectIdentity> OmittedSourceUnits { get; }
    /// <summary>Gets unresolved declared show, slide-tree, slide, drawable, text-storage, presenter-note and assessed table/text-formatting reference occurrences.</summary>
    public IReadOnlyList<IWorkSourceReferenceIssue> SourceReferenceIssues { get; }
    /// <summary>Gets unreadable or rejected declarations in selected content paths, without inferring nested references or omitted objects.</summary>
    public IReadOnlyList<IWorkSourceDeclarationIssue> SourceDeclarationIssues { get; }
    /// <summary>Gets the source presentation canvas size.</summary>
    public IWorkCanvasSize? SlideSize { get; }
    /// <summary>Gets projection diagnostics.</summary>
    public IReadOnlyList<IWorkDiagnostic> Diagnostics { get; }
    /// <summary>Gets whether at least one editable slide was recovered and all required slide references were resolved.</summary>
    public bool HasEditableContent => Slides.Count > 0 && _supportsEditableReconstruction;

    /// <summary>Gets whether bounded source content is available for an explicitly partial editable conversion.</summary>
    public bool HasRecoverableContent => Slides.Count > 0;

    /// <summary>Creates a conversion report for an OfficeIMO semantic-owner projection.</summary>
    public IWorkConversionReport CreateConversionReport(IWorkProjectionKind kind, IWorkPreviewAsset? preview = null) =>
        CreateConversionReport(kind, preview, Array.Empty<IWorkDiagnostic>());

    internal IWorkConversionReport CreateConversionReport(IWorkProjectionKind kind,
        IWorkPreviewAsset? preview, IReadOnlyList<IWorkDiagnostic> additionalDiagnostics,
        bool allowPartialEditableReconstruction = false) {
        ValidateReportRequest(kind, preview, allowPartialEditableReconstruction);
        return _source.CreateReport(kind, Diagnostics.Concat(additionalDiagnostics).ToArray(), preview,
            kind == IWorkProjectionKind.VisualFallback
                ? 0
                : Slides.Count + Slides.Sum(slide => slide.TextBoxes.Count + slide.Images.Count
                    + slide.Tables.Count(table => table.RowCount > 0 && table.ColumnCount > 0)
                    + slide.Tables.Where(table => table.RowCount > 0 && table.ColumnCount > 0).Sum(table => table.Cells.Count)
                    + (slide.TitleBox != null ? 1 : 0) + (slide.PresenterNotes.Length > 0 ? 1 : 0)),
            ReconstructedUnits(), OmittedUnits(), Slides.SelectMany(slide => slide.Tables), SourceReferenceIssues, SourceDeclarationIssues);
    }

    private void ValidateReportRequest(IWorkProjectionKind kind, IWorkPreviewAsset? preview,
        bool allowPartialEditableReconstruction) {
        if (kind == IWorkProjectionKind.EditableReconstruction && !HasEditableContent
            && !(allowPartialEditableReconstruction && HasRecoverableContent)) {
            throw new InvalidOperationException("Editable Keynote content was not recovered.");
        }
        if (kind == IWorkProjectionKind.VisualFallback && preview == null) {
            throw new ArgumentNullException(nameof(preview), "A visual fallback report requires the preview used by the owner.");
        }
    }
}

public sealed partial class IWorkSourceDocument {
    /// <summary>Reads a Keynote package into a bounded semantic source projection.</summary>
    public IWorkKeynoteProjection ReadKeynote() {
        _cancellationToken.ThrowIfCancellationRequested();
        if (Kind != IWorkDocumentKind.Keynote) throw new InvalidOperationException($"The source is {Kind}, not Keynote.");
        IWorkKeynoteProjection projection = IWorkKeynoteReader.Read(this);
        _cancellationToken.ThrowIfCancellationRequested();
        return projection;
    }
}

internal static partial class IWorkKeynoteReader {
    private const uint DocumentArchive = 1;
    private const uint ShowArchive = 2;
    private const uint SlideNodeArchive = 4;
    private const uint SlideArchive = 5;
    private const uint PlaceholderArchive = 7;
    private const uint PresenterNoteArchive = 15;
    private const uint TextStorageArchive = 2001;
    private const uint TextShapeArchive = 2011;

    internal static IWorkKeynoteProjection Read(IWorkSourceDocument source) {
        var references = new IWorkSourceReferenceIssueCollector(source);
        var diagnostics = new List<IWorkDiagnostic>();
        var slides = new List<IWorkKeynoteSlide>();
        var omittedUnits = new List<IWorkObjectIdentity>();
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
        IWorkWireMessage documentMessage;
        try {
            documentMessage = index.Message(document);
        } catch (InvalidDataException exception) when (!IWorkProtobuf.IsLimitException(exception)) {
            references.Declarations.Record(document, "$", null);
            diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                "IWORK_KEYNOTE_DOCUMENT_MALFORMED",
                "The Keynote document root is malformed; editable reconstruction is unavailable.",
                document.EntryPath, document.Identifier));
            return new IWorkKeynoteProjection(source, slides, null, diagnostics,
                supportsEditableReconstruction: false, sourceIdentity: new IWorkObjectIdentity(document),
                declarationIssues: references.Declarations.Issues);
        }
        bool showReferenceComplete = documentMessage.FieldCount(2) == 1
            && !documentMessage.HasUnexpectedWireKind(2, IWorkWireKind.Bytes);
        IWorkArchiveRecord? show = references.ReadOne(document, documentMessage, 2, allowedType: type => type == ShowArchive);
        if (!showReferenceComplete || show == null || show.MessageType != ShowArchive) {
            diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning, "IWORK_KEYNOTE_SHOW_MISSING",
                "The Keynote document root does not reference exactly one supported show object.", document.EntryPath, document.Identifier));
            return new IWorkKeynoteProjection(source, slides, null, diagnostics, supportsEditableReconstruction: false, sourceIdentity: new IWorkObjectIdentity(document), referenceIssues: references.Issues, declarationIssues: references.Declarations.Issues);
        }
        IWorkWireMessage showMessage;
        try {
            showMessage = index.Message(show);
        } catch (InvalidDataException exception) when (!IWorkProtobuf.IsLimitException(exception)) {
            references.Declarations.Record(show, "$", null);
            diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                "IWORK_KEYNOTE_SHOW_MALFORMED",
                "The Keynote show object is malformed; editable reconstruction is unavailable.",
                show.EntryPath, show.Identifier));
            return new IWorkKeynoteProjection(source, slides, null, diagnostics,
                supportsEditableReconstruction: false, sourceIdentity: new IWorkObjectIdentity(document), referenceIssues: references.Issues, declarationIssues: references.Declarations.Issues);
        }
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
        } catch (InvalidDataException exception) when (!IWorkProtobuf.IsLimitException(exception)) {
            slideReferenceCount = -1;
        }
        if (slideReferenceCount < 0 || slideTreeFieldCount != slideReferenceCount) {
            if (showMessage.HasField(3)) references.Declarations.Record(show, "3", showMessage.FieldCount(3),
                IWorkSourceDeclarationIssueKind.RejectedMessageSet);
            diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning, "IWORK_KEYNOTE_SLIDE_TREE_MISSING",
                "The Keynote show does not contain a supported slide tree.", show.EntryPath, show.Identifier));
            return new IWorkKeynoteProjection(source, slides, null, diagnostics, supportsEditableReconstruction: false, sourceIdentity: new IWorkObjectIdentity(document), referenceIssues: references.Issues, declarationIssues: references.Declarations.Issues);
        }
        if (slideReferenceCount > source.Options.MaximumProjectedSlides) {
            throw new InvalidDataException($"Keynote slide count exceeds the configured projection limit of {source.Options.MaximumProjectedSlides}.");
        }
        IWorkWireMessage slideTree;
        try {
            slideTree = showMessage.ParseNestedMessage(slideTreeBytes!);
        } catch (InvalidDataException exception) when (!IWorkProtobuf.IsLimitException(exception)) {
            references.Declarations.Record(show, "3", showMessage.FieldCount(3));
            diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning, "IWORK_KEYNOTE_SLIDE_TREE_MISSING",
                "The Keynote show does not contain a supported slide tree.", show.EntryPath, show.Identifier));
            return new IWorkKeynoteProjection(source, slides, null, diagnostics, supportsEditableReconstruction: false, sourceIdentity: new IWorkObjectIdentity(document), referenceIssues: references.Issues, declarationIssues: references.Declarations.Issues);
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
        int materializedCellCount = 0;
        var projectionBudget = new IWorkProjectionBudget(source.Options);
        var nodePositions = new List<int>();
        IReadOnlyList<IWorkArchiveRecord> nodes = references.ReadAll(
            show, slideTree, 2, out int unresolvedNodeCount, "3/2", allowedType: type => type == SlideNodeArchive,
            resolvedPositions: nodePositions);
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
        for (int nodeIndex = 0; nodeIndex < nodes.Count; nodeIndex++) {
            IWorkArchiveRecord node = nodes[nodeIndex];
            int position = nodePositions[nodeIndex];
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
            IWorkWireMessage nodeMessage;
            try {
                nodeMessage = index.Message(node);
            } catch (InvalidDataException exception) when (!IWorkProtobuf.IsLimitException(exception)) {
                references.Declarations.Record(node, "$", null);
                supportsEditableReconstruction = false;
                diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                    "IWORK_KEYNOTE_SLIDE_NODE_UNSUPPORTED",
                    "The Keynote slide tree references a malformed node record; editable reconstruction is incomplete.",
                    node.EntryPath, node.Identifier));
                continue;
            }
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
            IWorkArchiveRecord? slide = references.ReadOne(node, nodeMessage, 2, allowedType: type => type == SlideArchive);
            if (!slideReferenceComplete || slide == null || slide.MessageType != SlideArchive) {
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
            IWorkKeynoteSlide? projectedSlide = ReadSlide(source, index, slide, position, skipped,
                projectionBudget, ref materializedCellCount, diagnostics,
                ref supportsEditableReconstruction, omittedUnits, references);
            if (projectedSlide != null) slides.Add(projectedSlide);
            else omittedUnits.Add(new IWorkObjectIdentity(slide));
        }
        return new IWorkKeynoteProjection(source, slides, slideSize, diagnostics, supportsEditableReconstruction,
            new IWorkObjectIdentity(document), omittedUnits, references.Issues, references.Declarations.Issues);
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
        if (!show.HasField(4)) {
            complete = false;
            return null;
        }
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
        List<IWorkDiagnostic> diagnostics, ref bool supportsEditableReconstruction,
        IWorkTextContent? content = null) {
        supportsEditableReconstruction = false;
        if (diagnostics.Any(diagnostic => diagnostic.Code == "IWORK_KEYNOTE_TEXT_STORAGE_UNSUPPORTED"
            && diagnostic.RecordIdentifier == storage.Identifier)) return;
        diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
            "IWORK_KEYNOTE_TEXT_STORAGE_UNSUPPORTED",
            IWorkTextDiagnostics.Describe(content) + " Complete editable reconstruction is unavailable.",
            storage.EntryPath, storage.Identifier));
    }

    private static IWorkArchiveRecord? DrawableStorage(IWorkObjectIndex index, IWorkArchiveRecord drawable,
        IWorkSourceReferenceIssueCollector references,
        out bool complete) {
        complete = true;
        IWorkWireMessage message = index.Message(drawable);
        if (drawable.MessageType == TextShapeArchive) {
            IWorkArchiveRecord? field4 = references.ReadOne(drawable, message, 4, allowedType: type => type == TextStorageArchive);
            IWorkArchiveRecord? field2 = references.ReadOne(drawable, message, 2, allowedType: type => type == TextStorageArchive);
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
        IWorkArchiveRecord? nested = references.ReadOne(drawable, super, 2, "1/2", allowedType: type => type == TextStorageArchive);
        if (super.HasUnexpectedWireKind(2, IWorkWireKind.Bytes)
            || super.FieldCount(2) > 1
            || super.HasField(2) && (nested == null || nested.MessageType != TextStorageArchive)) {
            complete = false;
        }
        IWorkArchiveRecord? nestedStorage = nested?.MessageType == TextStorageArchive ? nested : null;
        return nestedStorage;
    }
}
