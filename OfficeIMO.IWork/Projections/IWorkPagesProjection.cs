using OfficeIMO.IWork.Internal;

namespace OfficeIMO.IWork;

/// <summary>One typed Pages drawable retained in source stacking order.</summary>
public sealed class IWorkPagesDrawable {
    internal IWorkPagesDrawable(IWorkTextBox textBox) {
        Kind = IWorkPagesDrawableKind.TextBox;
        TextBox = textBox;
    }

    internal IWorkPagesDrawable(IWorkImageAsset image) {
        Kind = IWorkPagesDrawableKind.Image;
        Image = image;
    }

    internal IWorkPagesDrawable(IWorkTable table) {
        Kind = IWorkPagesDrawableKind.Table;
        Table = table;
    }

    /// <summary>Gets the drawable kind.</summary>
    public IWorkPagesDrawableKind Kind { get; }
    /// <summary>Gets the text-box payload when <see cref="Kind"/> is <see cref="IWorkPagesDrawableKind.TextBox"/>.</summary>
    public IWorkTextBox? TextBox { get; }
    /// <summary>Gets the image payload when <see cref="Kind"/> is <see cref="IWorkPagesDrawableKind.Image"/>.</summary>
    public IWorkImageAsset? Image { get; }
    /// <summary>Gets the table payload when <see cref="Kind"/> is <see cref="IWorkPagesDrawableKind.Table"/>.</summary>
    public IWorkTable? Table { get; }
}

/// <summary>Read-only Pages structure recovered from a shared IWA object graph.</summary>
public sealed class IWorkPagesProjection {
    private readonly IWorkSourceDocument _source;
    private readonly bool _supportsEditableReconstruction;

    internal IWorkPagesProjection(IWorkSourceDocument source, IWorkTextContent body,
        IReadOnlyList<IWorkPagesSection> sections,
        IReadOnlyList<IWorkTextBox> textBoxObjects, IReadOnlyList<IWorkImageAsset> images,
        IReadOnlyList<IWorkTable> tables, IReadOnlyList<IWorkPagesDrawable> drawables,
        IWorkPageLayout? pageLayout,
        IReadOnlyList<IWorkDiagnostic> diagnostics, bool supportsEditableReconstruction) {
        _source = source;
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

    /// <summary>Creates an import report for an OfficeIMO semantic-owner projection.</summary>
    public IWorkImportReport CreateImportReport(IWorkProjectionKind kind, IWorkPreviewAsset? preview = null) {
        ValidateReportRequest(kind, preview);
        return _source.CreateReport(kind, Diagnostics, preview,
            kind == IWorkProjectionKind.VisualFallback
                ? 0
                : Body.Paragraphs.Count
                    + HeaderContents.Sum(content => content.Paragraphs.Count)
                    + FooterContents.Sum(content => content.Paragraphs.Count)
                    + TextBoxObjects.Count + Images.Count
                    + Tables.Count + Tables.Sum(table => table.Cells.Count));
    }

    private void ValidateReportRequest(IWorkProjectionKind kind, IWorkPreviewAsset? preview) {
        if (kind == IWorkProjectionKind.EditableReconstruction && !HasEditableContent) {
            throw new InvalidOperationException("Editable Pages content was not recovered.");
        }
        if (kind == IWorkProjectionKind.VisualFallback && preview == null) {
            throw new ArgumentNullException(nameof(preview), "A visual fallback report requires the preview used by the owner.");
        }
    }
}

public sealed partial class IWorkSourceDocument {
    /// <summary>Reads a Pages package into a bounded semantic source projection, or returns a diagnostic-only projection in visual-only mode.</summary>
    public IWorkPagesProjection ReadPages() {
        if (Kind != IWorkDocumentKind.Pages) throw new InvalidOperationException($"The source is {Kind}, not Pages.");
        if (RequestedImportMode == IWorkImportMode.VisualOnly) {
            return new IWorkPagesProjection(this, EmptyText(), Array.Empty<IWorkPagesSection>(),
                Array.Empty<IWorkTextBox>(), Array.Empty<IWorkImageAsset>(),
                Array.Empty<IWorkTable>(), Array.Empty<IWorkPagesDrawable>(), null,
                new[] { IWorkProjectionDiagnostics.SemanticProjectionSkipped }, supportsEditableReconstruction: false);
        }
        return IWorkPagesReader.Read(this);
    }

    private static IWorkTextContent EmptyText() => new(Array.Empty<IWorkTextParagraph>(), isComplete: false);
}

internal static class IWorkPagesReader {
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
        IWorkTextContent bodyContent = new(Array.Empty<IWorkTextParagraph>(), isComplete: false);
        var sections = new List<IWorkPagesSection>();
        var textBoxes = new List<IWorkTextBox>();
        var images = new List<IWorkImageAsset>();
        var tables = new List<IWorkTable>();
        var drawables = new List<IWorkPagesDrawable>();
        var projectedTextBoxes = new Dictionary<ulong, IWorkTextBox>();
        var projectedImages = new Dictionary<ulong, IWorkImageAsset>();
        var projectedTables = new Dictionary<ulong, IWorkTable>();
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
        IWorkWireMessage documentMessage = index.Message(document);
        IWorkPageLayout? pageLayout = ReadPageLayout(documentMessage, out bool pageLayoutComplete);
        if (!pageLayoutComplete) {
            supportsEditableReconstruction = false;
            diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                "IWORK_PAGES_LAYOUT_UNSUPPORTED",
                "The Pages document declares invalid page measurements; editable reconstruction is incomplete.",
                document.EntryPath, document.Identifier));
        }
        IWorkArchiveRecord? body = index.Dereference(documentMessage, 4);
        if (body == null || body.MessageType != TextStorageArchive) {
            supportsEditableReconstruction = false;
            diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning, "IWORK_PAGES_BODY_MISSING",
                "The Pages document root does not reference a supported body text storage.", document.EntryPath, document.Identifier));
        } else {
            bodyContent = IWorkTextReader.Read(index, body, projectionBudget);
            if (!bodyContent.IsComplete) MarkTextIncomplete(body, diagnostics, ref supportsEditableReconstruction);
            ReadHeadersAndFooters(index, body, sections, projectionBudget, diagnostics,
                ref supportsEditableReconstruction);
        }

        IReadOnlyList<IWorkArchiveRecord> documentDrawables = CollectDocumentDrawables(index, document,
            documentMessage, out bool drawableGraphComplete);
        if (!drawableGraphComplete) {
            supportsEditableReconstruction = false;
            diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                "IWORK_PAGES_DRAWABLE_UNSUPPORTED",
                "The Pages drawable graph is malformed or contains unresolved references; editable reconstruction is incomplete.",
                document.EntryPath, document.Identifier));
        }
        var skippedStorages = new HashSet<ulong>();
        if (body != null) skippedStorages.Add(body.Identifier);
        var textCache = new Dictionary<ulong, IWorkTextContent>();
        foreach (IWorkArchiveRecord shape in documentDrawables
                     .Where(record => record.MessageType == ShapeInfoArchive)) {
            IWorkWireMessage shapeMessage = index.Message(shape);
            IWorkArchiveRecord? field4Storage = index.Dereference(shapeMessage, 4);
            IWorkArchiveRecord? field2Storage = index.Dereference(shapeMessage, 2);
            IWorkArchiveRecord? storage = field4Storage ?? field2Storage;
            if (shapeMessage.LacksWireKind(4, IWorkWireKind.Bytes)
                || shapeMessage.LacksWireKind(2, IWorkWireKind.Bytes)
                || shapeMessage.HasField(4) && (field4Storage == null || field4Storage.MessageType != TextStorageArchive)
                || shapeMessage.HasField(2) && (field2Storage == null || field2Storage.MessageType != TextStorageArchive)) {
                supportsEditableReconstruction = false;
                if (!diagnostics.Any(diagnostic => diagnostic.Code == "IWORK_PAGES_DRAWABLE_UNSUPPORTED")) {
                    diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                        "IWORK_PAGES_DRAWABLE_UNSUPPORTED",
                        "A Pages drawable contains an unresolved text-storage reference; editable reconstruction is incomplete.",
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
            if (storage.MessageType != TextStorageArchive
                || skippedStorages.Contains(storage.Identifier)) continue;
            IWorkTextContent text;
            if (textCache.TryGetValue(storage.Identifier, out IWorkTextContent? cached)) {
                text = cached;
                projectionBudget.AddTextContentUse(text, includeCharacters: true);
            } else {
                text = IWorkTextReader.Read(index, storage, projectionBudget);
                textCache.Add(storage.Identifier, text);
            }
            if (!text.IsComplete) MarkTextIncomplete(storage, diagnostics, ref supportsEditableReconstruction);
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
                : IWorkDrawingReader.ReadGeometry(drawable, out geometryComplete);
            if (drawable != null && !geometryComplete) {
                supportsEditableReconstruction = false;
                if (!diagnostics.Any(diagnostic => diagnostic.Code == "IWORK_PAGES_DRAWABLE_UNSUPPORTED")) {
                    diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                        "IWORK_PAGES_DRAWABLE_UNSUPPORTED",
                        "A Pages drawable contains malformed geometry; editable reconstruction is incomplete.",
                        shape.EntryPath, shape.Identifier));
                }
            }
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
            if (text.Paragraphs.Count == 0) projectionBudget.AddTextItem();
            var textBox = new IWorkTextBox(text, geometry, hyperlink, accessibilityDescription);
            textBoxes.Add(textBox);
            projectedTextBoxes.Add(shape.Identifier, textBox);
        }
        IWorkArchiveRecord? unsupportedDrawable = documentDrawables.FirstOrDefault(record =>
            record.MessageType is not TextStorageArchive and not ShapeInfoArchive
                and not 3005 and not 6000 and not 6007);
        if (unsupportedDrawable != null) {
            supportsEditableReconstruction = false;
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
                    continue;
                }
                images.Add(image);
                projectedImages.Add(drawable.Identifier, image);
            }
        }
        int materializedCellCount = 0;
        foreach (IWorkArchiveRecord tableRecord in documentDrawables
                     .Where(record => record.MessageType is 6000 or 6007)) {
            projectionBudget.AddTable();
            IWorkTable? table = IWorkTableReader.Read(source, tableRecord, projectionBudget, diagnostics,
                ref materializedCellCount, ref supportsEditableReconstruction);
            if (table != null) {
                tables.Add(table);
                projectedTables.Add(tableRecord.Identifier, table);
            }
        }
        foreach (IWorkArchiveRecord drawable in documentDrawables) {
            if (projectedTextBoxes.TryGetValue(drawable.Identifier, out IWorkTextBox? textBox)) {
                drawables.Add(new IWorkPagesDrawable(textBox));
            } else if (projectedImages.TryGetValue(drawable.Identifier, out IWorkImageAsset? image)) {
                drawables.Add(new IWorkPagesDrawable(image));
            } else if (projectedTables.TryGetValue(drawable.Identifier, out IWorkTable? table)) {
                drawables.Add(new IWorkPagesDrawable(table));
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
            supportsEditableReconstruction);
    }

    private static IReadOnlyList<IWorkArchiveRecord> CollectDocumentDrawables(IWorkObjectIndex index,
        IWorkArchiveRecord document, IWorkWireMessage documentMessage, out bool complete) {
        complete = true;
        var identifiers = new HashSet<ulong>();
        var ordered = new List<IWorkArchiveRecord>();
        void Add(IWorkArchiveRecord record) {
            if (identifiers.Add(record.Identifier)) ordered.Add(record);
        }

        IWorkArchiveRecord? zOrder = index.Dereference(documentMessage, 20);
        if (documentMessage.LacksWireKind(20, IWorkWireKind.Bytes)
            || documentMessage.HasField(20) && zOrder == null) complete = false;
        if (zOrder != null) {
            int unresolvedZOrderCount;
            foreach (IWorkArchiveRecord record in index.DereferenceAll(
                         index.Message(zOrder), 1, out unresolvedZOrderCount)) Add(record);
            if (unresolvedZOrderCount > 0) complete = false;
        }
        IWorkArchiveRecord? floating = index.Dereference(documentMessage, 3);
        if (documentMessage.LacksWireKind(3, IWorkWireKind.Bytes)
            || documentMessage.HasField(3) && floating == null) complete = false;
        if (floating != null) {
            IReadOnlyList<IWorkWireMessage> pageGroups = IWorkObjectIndex.TryGetMessages(
                index.Message(floating), 1, out bool malformedPageGroups);
            if (malformedPageGroups) complete = false;
            foreach (IWorkWireMessage pageGroup in pageGroups) {
                foreach (int field in new[] { 2, 3, 4 }) {
                    IReadOnlyList<IWorkWireMessage> entries = IWorkObjectIndex.TryGetMessages(
                        pageGroup, field, out bool malformedEntries);
                    if (malformedEntries) complete = false;
                    foreach (IWorkWireMessage entry in entries) {
                        IWorkArchiveRecord? record = index.Dereference(entry, 1);
                        if (entry.LacksWireKind(1, IWorkWireKind.Bytes)
                            || entry.HasField(1) && record == null) complete = false;
                        else if (record != null) Add(record);
                    }
                }
            }
        }
        var reachable = new HashSet<ulong>(index.ReachableFrom(document).Select(record => record.Identifier));
        foreach (IWorkArchiveRecord record in index.PrimaryRecords.Where(record =>
                     reachable.Contains(record.Identifier)
                     && record.MessageType is ShapeInfoArchive or 3005 or 6000 or 6007)) Add(record);
        return Array.AsReadOnly(ordered.ToArray());
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
        if (fields.Any(field => document.LacksWireKind(field, IWorkWireKind.Fixed32)
                || document.HasField(field) && !document.GetFloat(field).HasValue)
            || document.LacksWireKind(42, IWorkWireKind.Varint)
            || width <= 0 || height <= 0 || new[] { width, height, left, right, top, bottom, header, footer }
                .Any(value => double.IsNaN(value) || double.IsInfinity(value) || value < 0)) {
            complete = false;
            return null;
        }
        return new IWorkPageLayout(width, height, left, right, top, bottom, header, footer,
            document.GetUnsigned(42).GetValueOrDefault() != 0);
    }

    private static void ReadHeadersAndFooters(IWorkObjectIndex index, IWorkArchiveRecord body,
        List<IWorkPagesSection> sections, IWorkProjectionBudget projectionBudget,
        List<IWorkDiagnostic> diagnostics, ref bool supportsEditableReconstruction) {
        IWorkWireMessage bodyMessage = index.Message(body);
        bool hasSectionTable = bodyMessage.HasField(17);
        IWorkWireMessage? sectionTable = IWorkObjectIndex.TryGetMessage(bodyMessage, 17, out bool malformedSectionTable);
        if (!hasSectionTable) return;
        if (bodyMessage.LacksWireKind(17, IWorkWireKind.Bytes)
            || malformedSectionTable || sectionTable == null) {
            supportsEditableReconstruction = false;
            diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                "IWORK_PAGES_SECTION_UNSUPPORTED",
                "The Pages section table is malformed; editable reconstruction is incomplete.",
                body.EntryPath, body.Identifier));
            return;
        }
        IReadOnlyList<IWorkWireMessage> entries = IWorkObjectIndex.TryGetMessages(
            sectionTable, 1, out bool malformedEntries);
        if (malformedEntries) {
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
            IReadOnlyList<IWorkArchiveRecord> referencedSections = index.DereferenceAll(
                entry, 2, out int unresolvedSectionCount);
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
            IWorkWireMessage sectionMessage = index.Message(section);
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
                IWorkArchiveRecord? archive = index.Dereference(sectionMessage, field);
                if (sectionMessage.LacksWireKind(field, IWorkWireKind.Bytes)
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
                IWorkWireMessage archiveMessage = index.Message(archive);
                AddSectionStorageText(index, archiveMessage, 1, archive, headers, new HashSet<ulong>(),
                    textCache, projectionBudget, diagnostics, ref supportsEditableReconstruction);
                AddSectionStorageText(index, archiveMessage, 2, archive, footers, new HashSet<ulong>(),
                    textCache, projectionBudget, diagnostics, ref supportsEditableReconstruction);
            }
            sections.Add(new IWorkPagesSection(sectionIndex++,
                firstPageHeaders, firstPageFooters, evenPageHeaders, evenPageFooters,
                defaultPageHeaders, defaultPageFooters));
        }
    }

    private static void AddSectionStorageText(IWorkObjectIndex index, IWorkWireMessage message, int field,
        IWorkArchiveRecord archive, List<IWorkTextContent> destination, HashSet<ulong> seen,
        Dictionary<ulong, IWorkTextContent> textCache, IWorkProjectionBudget projectionBudget,
        List<IWorkDiagnostic> diagnostics,
        ref bool supportsEditableReconstruction) {
        IReadOnlyList<IWorkArchiveRecord> storages = index.DereferenceAll(
            message, field, out int unresolvedStorageCount);
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
            if (!seen.Add(storage.Identifier)) continue;
            bool reused = textCache.TryGetValue(storage.Identifier, out IWorkTextContent? text);
            if (!reused) {
                text = IWorkTextReader.Read(index, storage, projectionBudget);
                textCache.Add(storage.Identifier, text);
            }
            if (text == null) throw new InvalidDataException("The cached Pages text content is unavailable.");
            if (!text.IsComplete) MarkTextIncomplete(storage, diagnostics, ref supportsEditableReconstruction);
            if (text.PlainText.Length == 0) continue;
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
        List<IWorkDiagnostic> diagnostics, ref bool supportsEditableReconstruction) {
        supportsEditableReconstruction = false;
        if (diagnostics.Any(diagnostic => diagnostic.Code == "IWORK_PAGES_TEXT_UNSUPPORTED")) return;
        diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
            "IWORK_PAGES_TEXT_UNSUPPORTED",
            "A Pages text storage contains an invalid UTF-8 run; editable reconstruction is incomplete.",
            storage.EntryPath, storage.Identifier));
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
