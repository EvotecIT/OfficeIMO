using OfficeIMO.IWork.Internal;

namespace OfficeIMO.IWork;

/// <summary>One typed Pages drawable retained in source stacking order.</summary>
public sealed class IWorkPagesDrawable {
    internal IWorkPagesDrawable(IWorkTextBox textBox, int? pageIndex = null) {
        Kind = IWorkPagesDrawableKind.TextBox;
        TextBox = textBox;
        PageIndex = pageIndex;
    }

    internal IWorkPagesDrawable(IWorkImageAsset image, int? pageIndex = null) {
        Kind = IWorkPagesDrawableKind.Image;
        Image = image;
        PageIndex = pageIndex;
    }

    internal IWorkPagesDrawable(IWorkTable table, int? pageIndex = null) {
        Kind = IWorkPagesDrawableKind.Table;
        Table = table;
        PageIndex = pageIndex;
    }

    /// <summary>Gets the drawable kind.</summary>
    public IWorkPagesDrawableKind Kind { get; }
    /// <summary>Gets the text-box payload when <see cref="Kind"/> is <see cref="IWorkPagesDrawableKind.TextBox"/>.</summary>
    public IWorkTextBox? TextBox { get; }
    /// <summary>Gets the image payload when <see cref="Kind"/> is <see cref="IWorkPagesDrawableKind.Image"/>.</summary>
    public IWorkImageAsset? Image { get; }
    /// <summary>Gets the table payload when <see cref="Kind"/> is <see cref="IWorkPagesDrawableKind.Table"/>.</summary>
    public IWorkTable? Table { get; }
    /// <summary>Gets the one-based source page index when the floating-canvas graph identifies it.</summary>
    public int? PageIndex { get; }
}

/// <summary>Read-only Pages structure recovered from a shared IWA object graph.</summary>
public sealed partial class IWorkPagesProjection {
    private readonly IWorkSourceDocument _source;
    private readonly bool _supportsEditableReconstruction;

    internal IWorkPagesProjection(IWorkSourceDocument source, IWorkTextContent body,
        IReadOnlyList<IWorkPagesSection> sections,
        IReadOnlyList<IWorkTextBox> textBoxObjects, IReadOnlyList<IWorkImageAsset> images,
        IReadOnlyList<IWorkTable> tables, IReadOnlyList<IWorkPagesDrawable> drawables,
        IWorkPageLayout? pageLayout,
        IReadOnlyList<IWorkDiagnostic> diagnostics, bool supportsEditableReconstruction,
        IWorkObjectIdentity? sourceIdentity = null, IReadOnlyList<IWorkObjectIdentity>? omittedUnits = null,
        IReadOnlyList<IWorkSourceReferenceIssue>? referenceIssues = null,
        IReadOnlyList<IWorkSourceDeclarationIssue>? declarationIssues = null) {
        _source = source;
        SourceIdentity = sourceIdentity;
        OmittedSourceUnits = Array.AsReadOnly((omittedUnits ?? Array.Empty<IWorkObjectIdentity>()).ToArray());
        SourceReferenceIssues = Array.AsReadOnly((referenceIssues ?? Array.Empty<IWorkSourceReferenceIssue>()).ToArray());
        SourceDeclarationIssues = Array.AsReadOnly((declarationIssues ?? Array.Empty<IWorkSourceDeclarationIssue>()).ToArray());
        Body = body;
        Sections = Array.AsReadOnly(sections.ToArray());
        HeaderContents = Array.AsReadOnly(Sections.SelectMany(section => section.HeaderContents).ToArray());
        FooterContents = Array.AsReadOnly(Sections.SelectMany(section => section.FooterContents).ToArray());
        TextBoxObjects = Array.AsReadOnly(textBoxObjects.ToArray());
        TextBoxContents = Array.AsReadOnly(TextBoxObjects.Select(textBox => textBox.Content).ToArray());
        Images = Array.AsReadOnly(images.ToArray());
        Tables = Array.AsReadOnly(tables.ToArray());
        Drawables = Array.AsReadOnly(drawables.ToArray());
        PageLayout = pageLayout;
        Paragraphs = Array.AsReadOnly(body.Paragraphs.Select(paragraph => paragraph.Text).ToArray());
        Headers = Array.AsReadOnly(HeaderContents.Select(content => content.PlainText).ToArray());
        Footers = Array.AsReadOnly(FooterContents.Select(content => content.PlainText).ToArray());
        TextBoxes = Array.AsReadOnly(TextBoxContents.Select(content => content.PlainText).ToArray());
        Diagnostics = Array.AsReadOnly(diagnostics.ToArray());
        _supportsEditableReconstruction = supportsEditableReconstruction;
    }

    /// <summary>Gets the rich body text and paragraph structure.</summary>
    public IWorkTextContent Body { get; }
    /// <summary>Gets the native document-root identity when exactly one root was identified.</summary>
    public IWorkObjectIdentity? SourceIdentity { get; }
    private IReadOnlyList<IWorkObjectIdentity> OmittedSourceUnits { get; }
    /// <summary>Gets unresolved declared body, drawable, text-storage, section, header/footer and assessed table/text-formatting reference occurrences.</summary>
    public IReadOnlyList<IWorkSourceReferenceIssue> SourceReferenceIssues { get; }
    /// <summary>Gets unreadable or rejected declarations in selected content paths, without inferring nested references or omitted objects.</summary>
    public IReadOnlyList<IWorkSourceDeclarationIssue> SourceDeclarationIssues { get; }
    /// <summary>Gets source sections with their associated header and footer content.</summary>
    public IReadOnlyList<IWorkPagesSection> Sections { get; }
    /// <summary>Gets rich header storages flattened in section order.</summary>
    public IReadOnlyList<IWorkTextContent> HeaderContents { get; }
    /// <summary>Gets rich footer storages flattened in section order.</summary>
    public IReadOnlyList<IWorkTextContent> FooterContents { get; }
    /// <summary>Gets floating rich text-box content in object order.</summary>
    public IReadOnlyList<IWorkTextContent> TextBoxContents { get; }
    /// <summary>Gets positioned rich text boxes.</summary>
    public IReadOnlyList<IWorkTextBox> TextBoxObjects { get; }
    /// <summary>Gets embedded document images in source drawable order.</summary>
    public IReadOnlyList<IWorkImageAsset> Images { get; }
    /// <summary>Gets editable tables reachable from the document graph.</summary>
    public IReadOnlyList<IWorkTable> Tables { get; }
    /// <summary>Gets text boxes, images, and tables in their shared source stacking order.</summary>
    public IReadOnlyList<IWorkPagesDrawable> Drawables { get; }
    /// <summary>Gets source page dimensions and margins.</summary>
    public IWorkPageLayout? PageLayout { get; }
    /// <summary>Gets body paragraphs in source order.</summary>
    public IReadOnlyList<string> Paragraphs { get; }
    /// <summary>Gets distinct section header text recovered from the source.</summary>
    public IReadOnlyList<string> Headers { get; }
    /// <summary>Gets distinct section footer text recovered from the source.</summary>
    public IReadOnlyList<string> Footers { get; }
    /// <summary>Gets floating text-box content in object order.</summary>
    public IReadOnlyList<string> TextBoxes { get; }
    /// <summary>Gets projection diagnostics.</summary>
    public IReadOnlyList<IWorkDiagnostic> Diagnostics { get; }
    /// <summary>Gets whether the supported editable document structure was recovered completely.</summary>
    public bool HasEditableContent => _supportsEditableReconstruction;

    /// <summary>Gets whether bounded source content is available for an explicitly partial editable conversion.</summary>
    public bool HasRecoverableContent => Body.Paragraphs.Count > 0 || Drawables.Count > 0;

    /// <summary>Creates a conversion report for an OfficeIMO semantic-owner projection.</summary>
    public IWorkConversionReport CreateConversionReport(IWorkProjectionKind kind, IWorkPreviewAsset? preview = null) =>
        CreateConversionReport(kind, preview, Array.Empty<IWorkDiagnostic>());

    internal IWorkConversionReport CreateConversionReport(IWorkProjectionKind kind,
        IWorkPreviewAsset? preview, IReadOnlyList<IWorkDiagnostic> additionalDiagnostics,
        bool allowPartialEditableReconstruction = false, int? reconstructedSectionCount = null) {
        ValidateReportRequest(kind, preview, allowPartialEditableReconstruction);
        return _source.CreateReport(kind, Diagnostics.Concat(additionalDiagnostics).ToArray(), preview,
            kind == IWorkProjectionKind.VisualFallback
                ? 0
                : Body.Paragraphs.Count
                    + ReconstructedSections(reconstructedSectionCount).Sum(section => section.HeaderContents.Sum(content => content.Paragraphs.Count)
                        + section.FooterContents.Sum(content => content.Paragraphs.Count))
                    + TextBoxObjects.Count + Images.Count
                    + Tables.Count(table => table.RowCount > 0 && table.ColumnCount > 0)
                    + Tables.Where(table => table.RowCount > 0 && table.ColumnCount > 0).Sum(table => table.Cells.Count),
            ReconstructedUnits(reconstructedSectionCount), OmittedUnits(reconstructedSectionCount), Tables, SourceReferenceIssues, SourceDeclarationIssues);
    }

    private void ValidateReportRequest(IWorkProjectionKind kind, IWorkPreviewAsset? preview,
        bool allowPartialEditableReconstruction) {
        if (kind == IWorkProjectionKind.EditableReconstruction && !HasEditableContent
            && !(allowPartialEditableReconstruction && HasRecoverableContent)) {
            throw new InvalidOperationException("Editable Pages content was not recovered.");
        }
        if (kind == IWorkProjectionKind.VisualFallback && preview == null) {
            throw new ArgumentNullException(nameof(preview), "A visual fallback report requires the preview used by the owner.");
        }
    }
}

public sealed partial class IWorkSourceDocument {
    /// <summary>Reads a Pages package into a bounded semantic source projection.</summary>
    public IWorkPagesProjection ReadPages() {
        _cancellationToken.ThrowIfCancellationRequested();
        if (Kind != IWorkDocumentKind.Pages) throw new InvalidOperationException($"The source is {Kind}, not Pages.");
        IWorkPagesProjection projection = IWorkPagesReader.Read(this);
        _cancellationToken.ThrowIfCancellationRequested();
        return projection;
    }
}

internal static partial class IWorkPagesReader {
    private const uint DocumentArchive = 10000;
    private const uint SectionArchive = 10011;
    private const uint HeadersFootersArchive = 10143;
    private const int FirstPageTemplateField = 23;
    private const int EvenPageTemplateField = 24;
    private const int DefaultPageTemplateField = 25;
    private const uint TextStorageArchive = 2001;
    private const uint ShapeInfoArchive = 2011;

    internal static IWorkPagesProjection Read(IWorkSourceDocument source) {
        var diagnostics = new List<IWorkDiagnostic>();
        IWorkTextContent bodyContent = new(Array.Empty<IWorkTextParagraph>(),
            isComplete: false, isTextComplete: false);
        var sections = new List<IWorkPagesSection>();
        var textBoxes = new List<IWorkTextBox>();
        var images = new List<IWorkImageAsset>();
        var tables = new List<IWorkTable>();
        var drawables = new List<IWorkPagesDrawable>();
        var projectedTextBoxes = new Dictionary<ulong, IWorkTextBox>();
        var projectedImages = new Dictionary<ulong, IWorkImageAsset>();
        var projectedTables = new Dictionary<ulong, IWorkTable>();
        var omittedUnits = new List<IWorkObjectIdentity>();
        var references = new IWorkSourceReferenceIssueCollector(source);
        var projectionBudget = new IWorkProjectionBudget(source.Options);
        IWorkObjectIndex index = source.Index;
        IWorkArchiveRecord? document = index.UniqueOfType(DocumentArchive, out bool duplicateDocument);
        if (document == null) {
            diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                duplicateDocument ? "IWORK_PAGES_DOCUMENT_DUPLICATE" : "IWORK_PAGES_DOCUMENT_MISSING",
                duplicateDocument
                    ? "More than one Pages document root was found; editable reconstruction is unavailable."
                    : "No supported Pages document root was found; editable reconstruction is unavailable."));
            return new IWorkPagesProjection(source, bodyContent, sections, textBoxes, images, tables,
                drawables, null, diagnostics,
                supportsEditableReconstruction: false);
        }

        bool supportsEditableReconstruction = true;
        IWorkWireMessage documentMessage;
        try {
            documentMessage = index.Message(document);
        } catch (InvalidDataException exception) when (!IWorkProtobuf.IsLimitException(exception)) {
            references.Declarations.Record(document, "$", null);
            diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                "IWORK_PAGES_DOCUMENT_MALFORMED",
                "The Pages document root is malformed; editable reconstruction is unavailable.",
                document.EntryPath, document.Identifier));
            return new IWorkPagesProjection(source, bodyContent, sections, textBoxes, images,
                tables, drawables, null, diagnostics,
                supportsEditableReconstruction: false, sourceIdentity: new IWorkObjectIdentity(document),
                declarationIssues: references.Declarations.Issues);
        }
        IWorkPageLayout? pageLayout = ReadPageLayout(documentMessage, out bool pageLayoutComplete);
        if (!pageLayoutComplete) {
            supportsEditableReconstruction = false;
            diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                "IWORK_PAGES_LAYOUT_UNSUPPORTED",
                "The Pages document declares invalid page-layout metadata; editable reconstruction is incomplete.",
                document.EntryPath, document.Identifier));
        }
        bool bodyReferenceComplete = documentMessage.FieldCount(4) == 1
            && !documentMessage.HasUnexpectedWireKind(4, IWorkWireKind.Bytes);
        IWorkArchiveRecord? body = references.ReadOne(document, documentMessage, 4);
        if (!bodyReferenceComplete || body == null || body.MessageType != TextStorageArchive) {
            supportsEditableReconstruction = false;
            diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning, "IWORK_PAGES_BODY_MISSING",
                "The Pages document root does not reference exactly one supported body text storage.", document.EntryPath, document.Identifier));
        } else {
            if (!TryReadMessage(index, body, references, out _)) {
                supportsEditableReconstruction = false;
                diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                    "IWORK_PAGES_BODY_MALFORMED",
                    "The Pages body text storage is malformed; editable reconstruction is incomplete.",
                    body.EntryPath, body.Identifier));
                omittedUnits.Add(new IWorkObjectIdentity(body));
            } else {
                bodyContent = IWorkTextReader.Read(index, body, projectionBudget, references, resolveInlineObjects: true);
                if (!bodyContent.IsComplete) MarkTextIncomplete(body, diagnostics, ref supportsEditableReconstruction, bodyContent);
                int maximumSectionCount = bodyContent.Paragraphs.Count(paragraph =>
                    paragraph.BreakKind == IWorkParagraphBreakKind.Section) + 1;
                ReadHeadersAndFooters(index, body, sections, projectionBudget, references, diagnostics,
                    maximumSectionCount, omittedUnits, ref supportsEditableReconstruction);
            }
        }

        IReadOnlyList<IWorkArchiveRecord> documentDrawables = CollectDocumentDrawables(index, document,
            documentMessage, projectionBudget, references, out IReadOnlyDictionary<ulong, int> drawablePageIndexes,
            out bool drawableGraphComplete, bodyContent);
        if (!drawableGraphComplete) {
            supportsEditableReconstruction = false;
            diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                "IWORK_PAGES_DRAWABLE_UNSUPPORTED",
                "The Pages drawable graph is malformed or contains unresolved references; editable reconstruction is incomplete.",
                document.EntryPath, document.Identifier));
        }
        var textCache = new Dictionary<ulong, IWorkTextContent>();
        foreach (IWorkArchiveRecord shape in documentDrawables
                     .Where(record => record.MessageType == ShapeInfoArchive)) {
            if (!TryReadMessage(index, shape, references, out IWorkWireMessage shapeMessage)) {
                supportsEditableReconstruction = false;
                diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                    "IWORK_PAGES_DRAWABLE_UNSUPPORTED",
                    "A Pages drawable record is malformed; editable reconstruction is incomplete.",
                    shape.EntryPath, shape.Identifier));
                continue;
            }
            IWorkArchiveRecord? field4Storage = references.ReadOne(shape, shapeMessage, 4);
            IWorkArchiveRecord? field2Storage = references.ReadOne(shape, shapeMessage, 2);
            IWorkArchiveRecord? storage = field4Storage ?? field2Storage;
            bool hasAmbiguousStorage = shapeMessage.FieldCount(4) > 1
                || shapeMessage.FieldCount(2) > 1
                || field4Storage != null && field2Storage != null
                    && field4Storage.Identifier != field2Storage.Identifier;
            if (hasAmbiguousStorage
                || shapeMessage.HasUnexpectedWireKind(4, IWorkWireKind.Bytes)
                || shapeMessage.HasUnexpectedWireKind(2, IWorkWireKind.Bytes)
                || shapeMessage.HasField(4) && (field4Storage == null || field4Storage.MessageType != TextStorageArchive)
                || shapeMessage.HasField(2) && (field2Storage == null || field2Storage.MessageType != TextStorageArchive)) {
                supportsEditableReconstruction = false;
                if (!diagnostics.Any(diagnostic => diagnostic.Code == "IWORK_PAGES_DRAWABLE_UNSUPPORTED")) {
                    diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                        "IWORK_PAGES_DRAWABLE_UNSUPPORTED",
                        "A Pages drawable contains an unresolved or ambiguous text-storage reference; editable reconstruction is incomplete.",
                        shape.EntryPath, shape.Identifier));
                }
                continue;
            }
            if (storage == null) {
                supportsEditableReconstruction = false;
                if (!diagnostics.Any(diagnostic => diagnostic.Code == "IWORK_PAGES_DRAWABLE_UNSUPPORTED")) {
                    diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                        "IWORK_PAGES_DRAWABLE_UNSUPPORTED",
                        "A Pages shape has no supported text storage; editable reconstruction is incomplete.",
                        shape.EntryPath, shape.Identifier));
                }
                continue;
            }
            if (storage.MessageType != TextStorageArchive) continue;
            IWorkTextContent text;
            if (textCache.TryGetValue(storage.Identifier, out IWorkTextContent? cached)) {
                text = cached;
                projectionBudget.AddTextContentUse(text, includeCharacters: true);
            } else {
                if (!TryReadMessage(index, storage, references, out _)) {
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
            IWorkWireMessage? drawable = IWorkDrawingReader.DrawableMessage(index, shape,
                out bool drawableComplete);
            if (!drawableComplete) {
                supportsEditableReconstruction = false;
                if (!diagnostics.Any(diagnostic => diagnostic.Code == "IWORK_PAGES_DRAWABLE_UNSUPPORTED")) {
                    diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                        "IWORK_PAGES_DRAWABLE_UNSUPPORTED",
                        "A Pages drawable contains malformed geometry; editable reconstruction is incomplete.",
                        shape.EntryPath, shape.Identifier));
                }
            }
            bool geometryComplete = true;
            IWorkGeometry? geometry = drawable == null
                ? null
                : IWorkDrawingReader.ReadGeometry(drawable, out geometryComplete,
                    requirePositiveSize: true);
            bool metadataComplete = true;
            string? hyperlink = IWorkDrawingReader.ReadOptionalString(drawable, 4,
                projectionBudget, ref metadataComplete);
            string? accessibilityDescription = IWorkDrawingReader.ReadOptionalString(drawable, 8,
                projectionBudget, ref metadataComplete);
            if (!metadataComplete) {
                supportsEditableReconstruction = false;
                if (!diagnostics.Any(diagnostic => diagnostic.Code == "IWORK_PAGES_DRAWABLE_UNSUPPORTED")) {
                    diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                        "IWORK_PAGES_DRAWABLE_UNSUPPORTED",
                        "A Pages drawable contains invalid text metadata; editable reconstruction is incomplete.",
                        shape.EntryPath, shape.Identifier));
                }
            }
            if (text.PlainText.Length == 0 && hyperlink == null
                && accessibilityDescription == null) continue;
            if (!geometryComplete || !IWorkDrawingReader.HasPositiveSize(geometry)) {
                supportsEditableReconstruction = false;
                if (!diagnostics.Any(diagnostic => diagnostic.Code == "IWORK_PAGES_DRAWABLE_UNSUPPORTED")) {
                    diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                        "IWORK_PAGES_DRAWABLE_UNSUPPORTED",
                        "A Pages drawable contains malformed geometry; editable reconstruction is incomplete.",
                        shape.EntryPath, shape.Identifier));
                }
            }
            if (text.Paragraphs.Count == 0) projectionBudget.AddTextItem();
            var textBox = new IWorkTextBox(text, geometry, hyperlink, accessibilityDescription, new IWorkObjectIdentity(shape));
            textBoxes.Add(textBox);
            projectedTextBoxes.Add(shape.Identifier, textBox);
        }
        foreach (IWorkArchiveRecord unsupportedDrawable in documentDrawables.Where(record =>
                     record.MessageType is not TextStorageArchive and not ShapeInfoArchive
                         and not 3005 and not 6000 and not 6007)) {
            omittedUnits.Add(new IWorkObjectIdentity(unsupportedDrawable));
            supportsEditableReconstruction = false;
            if (!diagnostics.Any(diagnostic => diagnostic.Code == "IWORK_PAGES_DRAWABLE_UNSUPPORTED"))
                diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                "IWORK_PAGES_DRAWABLE_UNSUPPORTED",
                $"Pages drawable type {unsupportedDrawable.MessageType} is preserved but cannot be reconstructed; editable reconstruction is incomplete.",
                unsupportedDrawable.EntryPath, unsupportedDrawable.Identifier));
        }
        var seenImages = new HashSet<ulong>();
        foreach (IWorkArchiveRecord drawable in documentDrawables) {
            if (drawable.MessageType == 3005 && seenImages.Add(drawable.Identifier)) {
                projectionBudget.AddImage();
                IWorkImageAsset? image = IWorkDrawingReader.ReadImage(source, drawable,
                    projectionBudget, out bool imageComplete);
                if (!imageComplete || image == null) {
                    supportsEditableReconstruction = false;
                    diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                        "IWORK_PAGES_IMAGE_UNSUPPORTED",
                        "A Pages document image could not be resolved completely; editable reconstruction is incomplete.",
                        drawable.EntryPath, drawable.Identifier));
                    omittedUnits.Add(new IWorkObjectIdentity(drawable));
                    continue;
                }
                projectionBudget.AddProjectedImageBytes(image.Length);
                images.Add(image);
                projectedImages.Add(drawable.Identifier, image);
            }
        }
        int materializedCellCount = 0;
        foreach (IWorkArchiveRecord tableRecord in documentDrawables
                     .Where(record => record.MessageType is 6000 or 6007)) {
            projectionBudget.AddTable();
            IWorkTable? table = IWorkTableReader.Read(source, tableRecord, projectionBudget, references, diagnostics,
                ref materializedCellCount, ref supportsEditableReconstruction);
            if (table != null) {
                tables.Add(table);
                projectedTables.Add(tableRecord.Identifier, table);
            } else omittedUnits.Add(new IWorkObjectIdentity(tableRecord));
        }
        foreach (IWorkArchiveRecord drawable in documentDrawables) {
            int? pageIndex = drawablePageIndexes.TryGetValue(drawable.Identifier, out int sourcePageIndex)
                ? sourcePageIndex
                : null;
            if (projectedTextBoxes.TryGetValue(drawable.Identifier, out IWorkTextBox? textBox)) {
                drawables.Add(new IWorkPagesDrawable(textBox, pageIndex));
            } else if (projectedImages.TryGetValue(drawable.Identifier, out IWorkImageAsset? image)) {
                drawables.Add(new IWorkPagesDrawable(image, pageIndex));
            } else if (projectedTables.TryGetValue(drawable.Identifier, out IWorkTable? table)) {
                drawables.Add(new IWorkPagesDrawable(table, pageIndex));
            }
        }
        int bodySectionCount = bodyContent.Paragraphs.Count(paragraph =>
            paragraph.BreakKind == IWorkParagraphBreakKind.Section) + 1;
        if (sections.Count > 0 && bodySectionCount != sections.Count) {
            supportsEditableReconstruction = false;
            diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                "IWORK_PAGES_SECTION_UNSUPPORTED",
                "Pages body section boundaries do not match the section table; editable reconstruction is incomplete.",
                body?.EntryPath, body?.Identifier));
        }
        return new IWorkPagesProjection(source, bodyContent, sections, textBoxes, images, tables,
            drawables, pageLayout, diagnostics,
            supportsEditableReconstruction, new IWorkObjectIdentity(document), omittedUnits, references.Issues, references.Declarations.Issues);
    }

    private static IWorkPageLayout? ReadPageLayout(IWorkWireMessage document, out bool complete) {
        int[] fields = { 30, 31, 32, 33, 34, 35, 36, 37 };
        bool declared = fields.Any(document.HasField);
        complete = true;
        if (!declared) return null;
        double width = document.GetFloat(30) ?? 0;
        double height = document.GetFloat(31) ?? 0;
        double left = document.GetFloat(32) ?? 0;
        double right = document.GetFloat(33) ?? 0;
        double top = document.GetFloat(34) ?? 0;
        double bottom = document.GetFloat(35) ?? 0;
        double header = document.GetFloat(36) ?? 0;
        double footer = document.GetFloat(37) ?? 0;
        ulong? landscape = document.GetUnsigned(42);
        if (fields.Any(field => document.FieldCount(field) > 1
                || document.HasUnexpectedWireKind(field, IWorkWireKind.Fixed32)
                || document.HasField(field) && !document.GetFloat(field).HasValue)
            || document.FieldCount(42) > 1
            || document.HasUnexpectedWireKind(42, IWorkWireKind.Varint)
            || landscape > 1
            || width <= 0 || height <= 0 || new[] { width, height, left, right, top, bottom, header, footer }
                .Any(value => double.IsNaN(value) || double.IsInfinity(value) || value < 0)) {
            complete = false;
            return null;
        }
        return new IWorkPageLayout(width, height, left, right, top, bottom, header, footer,
            landscape == 1);
    }

    private static void ReadHeadersAndFooters(IWorkObjectIndex index, IWorkArchiveRecord body,
        List<IWorkPagesSection> sections, IWorkProjectionBudget projectionBudget,
        IWorkSourceReferenceIssueCollector references,
        List<IWorkDiagnostic> diagnostics, int maximumSectionCount, List<IWorkObjectIdentity> omittedUnits,
        ref bool supportsEditableReconstruction) {
        IWorkWireMessage bodyMessage = index.Message(body);
        bool hasSectionTable = bodyMessage.HasField(17);
        if (!hasSectionTable) return;
        byte[]? sectionTableBytes = bodyMessage.GetBytes(17);
        int declaredSectionCount;
        int totalSectionTableFields = -1;
        try {
            declaredSectionCount = sectionTableBytes == null
                ? -1
                : bodyMessage.CountNestedFields(sectionTableBytes, 1,
                    out totalSectionTableFields);
        } catch (InvalidDataException exception) when (!IWorkProtobuf.IsLimitException(exception)) {
            declaredSectionCount = -1;
            totalSectionTableFields = -1;
        }
        if (bodyMessage.HasUnexpectedWireKind(17, IWorkWireKind.Bytes)
            || bodyMessage.FieldCount(17) != 1 || declaredSectionCount < 0
            || totalSectionTableFields != declaredSectionCount
            || declaredSectionCount > maximumSectionCount) {
            supportsEditableReconstruction = false;
            diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                "IWORK_PAGES_SECTION_UNSUPPORTED",
                "The Pages section table is malformed; editable reconstruction is incomplete.",
                body.EntryPath, body.Identifier));
            references.Declarations.Record(body, "17", bodyMessage.FieldCount(17),
                IWorkSourceDeclarationIssueKind.RejectedMessageSet);
            return;
        }
        IWorkWireMessage sectionTable;
        try {
            sectionTable = bodyMessage.ParseNestedMessage(sectionTableBytes!);
        } catch (InvalidDataException exception) when (!IWorkProtobuf.IsLimitException(exception)) {
            supportsEditableReconstruction = false;
            diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                "IWORK_PAGES_SECTION_UNSUPPORTED",
                "The Pages section table is malformed; editable reconstruction is incomplete.",
                body.EntryPath, body.Identifier));
            references.Declarations.Record(body, "17", bodyMessage.FieldCount(17),
                IWorkSourceDeclarationIssueKind.RejectedMessageSet);
            return;
        }
        IReadOnlyList<IWorkWireMessage> entries = IWorkObjectIndex.TryGetMessages(
            sectionTable, 1, out bool malformedEntries);
        if (malformedEntries) {
            references.Declarations.Record(body, "17/1", sectionTable.FieldCount(1),
                IWorkSourceDeclarationIssueKind.RejectedMessageSet);
            supportsEditableReconstruction = false;
            diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                "IWORK_PAGES_SECTION_UNSUPPORTED",
                "The Pages section table is malformed; editable reconstruction is incomplete.",
                body.EntryPath, body.Identifier));
        }
        var textCache = new Dictionary<ulong, IWorkTextContent>();
        int sectionIndex = 0;
        foreach (IWorkWireMessage entry in entries) {
            List<IWorkTextContent>? firstPageHeaders = null;
            List<IWorkTextContent>? firstPageFooters = null;
            List<IWorkTextContent>? evenPageHeaders = null;
            List<IWorkTextContent>? evenPageFooters = null;
            List<IWorkTextContent>? defaultPageHeaders = null;
            List<IWorkTextContent>? defaultPageFooters = null;
            IReadOnlyList<IWorkArchiveRecord> referencedSections = references.ReadAll(
                body, entry, 2, out int unresolvedSectionCount,
                "17/1[" + (sectionIndex + 1).ToString(System.Globalization.CultureInfo.InvariantCulture) + "]/2");
            if (unresolvedSectionCount > 0 || referencedSections.Count != 1
                || referencedSections[0].MessageType != SectionArchive) {
                supportsEditableReconstruction = false;
                if (!diagnostics.Any(diagnostic => diagnostic.Code == "IWORK_PAGES_SECTION_UNSUPPORTED")) {
                    diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                        "IWORK_PAGES_SECTION_UNSUPPORTED",
                        "The Pages section table contains an unresolved or unsupported section; editable reconstruction is incomplete.",
                        body.EntryPath, body.Identifier));
                }
                sections.Add(new IWorkPagesSection(sectionIndex++, null, null, null, null, null, null));
                continue;
            }
            IWorkArchiveRecord section = referencedSections[0];
            if (!TryReadMessage(index, section, references, out IWorkWireMessage sectionMessage)) {
                supportsEditableReconstruction = false;
                diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                    "IWORK_PAGES_SECTION_UNSUPPORTED",
                    "A Pages section record is malformed; editable reconstruction is incomplete.",
                    section.EntryPath, section.Identifier));
                sections.Add(new IWorkPagesSection(sectionIndex++, null, null, null, null, null, null));
                continue;
            }
            foreach (int field in new[] {
                         FirstPageTemplateField, EvenPageTemplateField, DefaultPageTemplateField
                     }) {
                if (!sectionMessage.HasField(field)) continue;
                var headers = new List<IWorkTextContent>();
                var footers = new List<IWorkTextContent>();
                switch (field) {
                    case FirstPageTemplateField:
                        firstPageHeaders = headers;
                        firstPageFooters = footers;
                        break;
                    case EvenPageTemplateField:
                        evenPageHeaders = headers;
                        evenPageFooters = footers;
                        break;
                    default:
                        defaultPageHeaders = headers;
                        defaultPageFooters = footers;
                        break;
                }
                bool templateReferenceComplete = sectionMessage.FieldCount(field) == 1
                    && !sectionMessage.HasUnexpectedWireKind(field, IWorkWireKind.Bytes);
                IWorkArchiveRecord? archive = references.ReadOne(section, sectionMessage, field);
                if (!templateReferenceComplete
                    || archive == null || archive.MessageType != HeadersFootersArchive) {
                    supportsEditableReconstruction = false;
                    if (!diagnostics.Any(diagnostic => diagnostic.Code == "IWORK_PAGES_HEADER_FOOTER_UNSUPPORTED")) {
                        diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                            "IWORK_PAGES_HEADER_FOOTER_UNSUPPORTED",
                            "A Pages section contains an unresolved or unsupported header/footer archive; editable reconstruction is incomplete.",
                            section.EntryPath, section.Identifier));
                    }
                    continue;
                }
                if (!TryReadMessage(index, archive, references, out IWorkWireMessage archiveMessage)) {
                    supportsEditableReconstruction = false;
                    diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                        "IWORK_PAGES_HEADER_FOOTER_UNSUPPORTED",
                        "A Pages header/footer archive is malformed; editable reconstruction is incomplete.",
                        archive.EntryPath, archive.Identifier));
                    continue;
                }
                AddSectionStorageText(index, archiveMessage, 1, archive, headers, new HashSet<ulong>(),
                    textCache, projectionBudget, references, diagnostics, omittedUnits, ref supportsEditableReconstruction);
                AddSectionStorageText(index, archiveMessage, 2, archive, footers, new HashSet<ulong>(),
                    textCache, projectionBudget, references, diagnostics, omittedUnits, ref supportsEditableReconstruction);
            }
            sections.Add(new IWorkPagesSection(sectionIndex++,
                firstPageHeaders, firstPageFooters, evenPageHeaders, evenPageFooters,
                defaultPageHeaders, defaultPageFooters));
        }
    }

    private static void AddSectionStorageText(IWorkObjectIndex index, IWorkWireMessage message, int field,
        IWorkArchiveRecord archive, List<IWorkTextContent> destination, HashSet<ulong> seen,
        Dictionary<ulong, IWorkTextContent> textCache, IWorkProjectionBudget projectionBudget,
        IWorkSourceReferenceIssueCollector references,
        List<IWorkDiagnostic> diagnostics, List<IWorkObjectIdentity> omittedUnits,
        ref bool supportsEditableReconstruction) {
        IReadOnlyList<IWorkArchiveRecord> storages = references.ReadAll(
            archive, message, field, out int unresolvedStorageCount);
        if (unresolvedStorageCount > 0) {
            supportsEditableReconstruction = false;
            diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                "IWORK_PAGES_HEADER_FOOTER_UNSUPPORTED",
                "A Pages header or footer contains an unresolved text reference; editable reconstruction is incomplete.",
                archive.EntryPath, archive.Identifier));
        }
        foreach (IWorkArchiveRecord storage in storages) {
            if (storage.MessageType != TextStorageArchive) {
                supportsEditableReconstruction = false;
                if (!diagnostics.Any(diagnostic => diagnostic.Code == "IWORK_PAGES_HEADER_FOOTER_UNSUPPORTED")) {
                    diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                        "IWORK_PAGES_HEADER_FOOTER_UNSUPPORTED",
                        "A Pages header or footer references an unsupported text object; editable reconstruction is incomplete.",
                        archive.EntryPath, archive.Identifier));
                }
                continue;
            }
            if (!seen.Add(storage.Identifier)) {
                supportsEditableReconstruction = false;
                if (!diagnostics.Any(diagnostic =>
                        diagnostic.Code == "IWORK_PAGES_HEADER_FOOTER_DUPLICATE")) {
                    diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                        "IWORK_PAGES_HEADER_FOOTER_DUPLICATE",
                        "A Pages header or footer repeats the same text storage; editable reconstruction is incomplete.",
                        archive.EntryPath, archive.Identifier));
                }
                continue;
            }
            bool reused = textCache.TryGetValue(storage.Identifier, out IWorkTextContent? text);
            if (!reused) {
                if (!TryReadMessage(index, storage, references, out _)) {
                    MarkTextIncomplete(storage, diagnostics, ref supportsEditableReconstruction);
                    omittedUnits.Add(new IWorkObjectIdentity(storage));
                    continue;
                }
                text = IWorkTextReader.Read(index, storage, projectionBudget, references);
                textCache.Add(storage.Identifier, text);
            }
            if (text == null) throw new InvalidDataException("The cached Pages text content is unavailable.");
            if (!text.IsComplete) MarkTextIncomplete(storage, diagnostics, ref supportsEditableReconstruction, text);
            if (!text.IsTextComplete && text.Paragraphs.Count == 0)
                omittedUnits.Add(new IWorkObjectIdentity(storage));
            if (text.PlainText.Length == 0) {
                if (!text.IsTextComplete) omittedUnits.Add(new IWorkObjectIdentity(storage));
                continue;
            }
            if (reused) projectionBudget.AddTextContentUse(text, includeCharacters: true);
            destination.Add(text);
        }
    }

    internal static string StorageText(IWorkWireMessage storage, IWorkProjectionBudget projectionBudget,
        out bool fullyDecoded) {
        var text = new System.Text.StringBuilder();
        fullyDecoded = true;
        if (storage.HasUnexpectedWireKind(3, IWorkWireKind.Bytes)) fullyDecoded = false;
        foreach (byte[] bytes in storage.EnumerateRepeatedBytes(3)) {
            if (IWorkTextReader.TryDecodeUtf8(bytes, projectionBudget, out string part)) text.Append(part);
            else fullyDecoded = false;
        }
        string value = text.ToString();
        if (value.IndexOf('\ufffc') >= 0 || value.IndexOf('\ufffb') >= 0) fullyDecoded = false;
        return CleanText(value);
    }

    private static void MarkTextIncomplete(IWorkArchiveRecord storage,
        List<IWorkDiagnostic> diagnostics, ref bool supportsEditableReconstruction,
        IWorkTextContent? content = null) {
        supportsEditableReconstruction = false;
        if (diagnostics.Any(diagnostic => diagnostic.Code == "IWORK_PAGES_TEXT_UNSUPPORTED"
            && diagnostic.RecordIdentifier == storage.Identifier)) return;
        diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
            "IWORK_PAGES_TEXT_UNSUPPORTED",
            IWorkTextDiagnostics.Describe(content) + " Complete editable reconstruction is unavailable.",
            storage.EntryPath, storage.Identifier));
    }

    private static bool TryReadMessage(IWorkObjectIndex index, IWorkArchiveRecord record,
        IWorkSourceReferenceIssueCollector references, out IWorkWireMessage message) {
        try {
            message = index.Message(record);
            return true;
        } catch (InvalidDataException exception) when (!IWorkProtobuf.IsLimitException(exception)) {
            references.Declarations.Record(record, "$", null);
            message = null!;
            return false;
        }
    }

    internal static string CleanText(string value) => value
        .Replace("\uFFFC", string.Empty)
        .Replace("\uFFFB", string.Empty)
        .Replace("\u0004", "\n")
        .Replace("\u0005", "\n")
        .Replace("\u000C", "\n")
        .Replace("\u2028", "\n")
        .Replace("\u2029", "\n")
        .Replace("\r\n", "\n")
        .Replace("\r", "\n");

}
