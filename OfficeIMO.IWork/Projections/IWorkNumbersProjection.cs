using OfficeIMO.IWork.Internal;

namespace OfficeIMO.IWork;

/// <summary>Read-only Numbers structure recovered from a shared IWA object graph.</summary>
public sealed partial class IWorkNumbersProjection {
    private readonly IWorkSourceDocument _source;
    private readonly bool _supportsEditableReconstruction;

    internal IWorkNumbersProjection(IWorkSourceDocument source, IReadOnlyList<IWorkNumbersSheet> sheets,
        IReadOnlyList<IWorkDiagnostic> diagnostics, bool supportsEditableReconstruction,
        IWorkObjectIdentity? sourceIdentity = null, IReadOnlyList<IWorkObjectIdentity>? omittedUnits = null,
        IReadOnlyList<IWorkSourceReferenceIssue>? referenceIssues = null) {
        _source = source;
        SourceIdentity = sourceIdentity;
        OmittedSourceUnits = Array.AsReadOnly((omittedUnits ?? Array.Empty<IWorkObjectIdentity>()).ToArray());
        SourceReferenceIssues = Array.AsReadOnly((referenceIssues ?? Array.Empty<IWorkSourceReferenceIssue>()).ToArray());
        Sheets = Array.AsReadOnly(sheets.ToArray());
        Diagnostics = Array.AsReadOnly(diagnostics.ToArray());
        _supportsEditableReconstruction = supportsEditableReconstruction;
    }

    /// <summary>Gets sheets in source order.</summary>
    public IReadOnlyList<IWorkNumbersSheet> Sheets { get; }
    /// <summary>Gets the native document-root identity when exactly one root was identified.</summary>
    public IWorkObjectIdentity? SourceIdentity { get; }
    private IReadOnlyList<IWorkObjectIdentity> OmittedSourceUnits { get; }
    /// <summary>Gets unresolved declared sheet, drawable and text-storage reference occurrences in assessed content paths.</summary>
    public IReadOnlyList<IWorkSourceReferenceIssue> SourceReferenceIssues { get; }
    /// <summary>Gets projection diagnostics.</summary>
    public IReadOnlyList<IWorkDiagnostic> Diagnostics { get; }
    /// <summary>Gets whether at least one editable sheet was recovered and its required semantic references were resolved.</summary>
    public bool HasEditableContent => Sheets.Count > 0 && _supportsEditableReconstruction;

    /// <summary>Gets whether bounded source content is available for an explicitly partial editable conversion.</summary>
    public bool HasRecoverableContent => Sheets.Count > 0;

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
                : Sheets.Count + Sheets.Sum(sheet => sheet.TextBoxes.Count + sheet.Tables.Count
                    + sheet.Tables.Sum(table => table.Cells.Count)), ReconstructedUnits(), OmittedUnits(), Sheets.SelectMany(sheet => sheet.Tables), SourceReferenceIssues);
    }

    private void ValidateReportRequest(IWorkProjectionKind kind, IWorkPreviewAsset? preview,
        bool allowPartialEditableReconstruction) {
        if (kind == IWorkProjectionKind.EditableReconstruction && !HasEditableContent
            && !(allowPartialEditableReconstruction && HasRecoverableContent)) {
            throw new InvalidOperationException("Editable Numbers content was not recovered.");
        }
        if (kind == IWorkProjectionKind.VisualFallback && preview == null) {
            throw new ArgumentNullException(nameof(preview), "A visual fallback report requires the preview used by the owner.");
        }
    }
}

public sealed partial class IWorkSourceDocument {
    /// <summary>Reads a Numbers package into a bounded semantic source projection.</summary>
    public IWorkNumbersProjection ReadNumbers() {
        _cancellationToken.ThrowIfCancellationRequested();
        if (Kind != IWorkDocumentKind.Numbers) throw new InvalidOperationException($"The source is {Kind}, not Numbers.");
        IWorkNumbersProjection projection = IWorkNumbersReader.Read(this);
        _cancellationToken.ThrowIfCancellationRequested();
        return projection;
    }
}

internal static class IWorkNumbersReader {
    private const uint DocumentArchive = 1;
    private const uint SheetArchive = 2;
    private const uint TableInfoArchive = 6000;
    private const uint WordProcessingTableInfoArchive = 6007;
    private const uint TextStorageArchive = 2001;
    private const uint TextShapeArchive = 2011;

    internal static IWorkNumbersProjection Read(IWorkSourceDocument source) {
        var diagnostics = new List<IWorkDiagnostic>();
        var sheets = new List<IWorkNumbersSheet>();
        var omittedUnits = new List<IWorkObjectIdentity>();
        var references = new IWorkSourceReferenceIssueCollector(source);
        IWorkObjectIndex index = source.Index;
        IWorkArchiveRecord? document = index.UniqueOfType(DocumentArchive, out bool duplicateDocument);
        if (document == null) {
            diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                duplicateDocument ? "IWORK_NUMBERS_DOCUMENT_DUPLICATE" : "IWORK_NUMBERS_DOCUMENT_MISSING",
                duplicateDocument
                    ? "More than one Numbers document root was found; editable reconstruction is unavailable."
                    : "No supported Numbers document root was found; editable reconstruction is unavailable."));
            return new IWorkNumbersProjection(source, sheets, diagnostics, supportsEditableReconstruction: false);
        }
        int materializedCellCount = 0;
        var projectionBudget = new IWorkProjectionBudget(source.Options);
        var projectedDrawableIdentifiers = new HashSet<ulong>();
        bool supportsEditableReconstruction = true;
        IWorkWireMessage documentMessage;
        int declaredSheetCount;
        try {
            declaredSheetCount = IWorkProtobuf.CountFields(document.Payload, 1,
                source.Options.MaximumProtobufFieldCount);
            documentMessage = index.Message(document);
        } catch (InvalidDataException) {
            diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                "IWORK_NUMBERS_DOCUMENT_MALFORMED",
                "The Numbers document root is malformed; editable reconstruction is unavailable.",
                document.EntryPath, document.Identifier));
            return new IWorkNumbersProjection(source, sheets, diagnostics,
                supportsEditableReconstruction: false, sourceIdentity: new IWorkObjectIdentity(document));
        }
        if (declaredSheetCount > source.Options.MaximumProjectedSheets) {
            throw new InvalidDataException($"Numbers sheet count exceeds the configured projection limit of {source.Options.MaximumProjectedSheets}.");
        }
        IReadOnlyList<IWorkArchiveRecord> sheetRecords = references.ReadAll(
            document, documentMessage, 1, out int unresolvedSheetCount);
        if (sheetRecords.Count > source.Options.MaximumProjectedSheets) {
            throw new InvalidDataException($"Numbers sheet count exceeds the configured projection limit of {source.Options.MaximumProjectedSheets}.");
        }
        if (unresolvedSheetCount > 0) {
            supportsEditableReconstruction = false;
            diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                "IWORK_NUMBERS_SHEET_UNSUPPORTED",
                "The Numbers document references a missing sheet; editable reconstruction is incomplete.",
                document.EntryPath, document.Identifier));
        }

        var projectedSheetIdentifiers = new HashSet<ulong>();
        foreach (IWorkArchiveRecord sheetRecord in sheetRecords) {
            if (!projectedSheetIdentifiers.Add(sheetRecord.Identifier)) {
                supportsEditableReconstruction = false;
                if (!diagnostics.Any(diagnostic => diagnostic.Code == "IWORK_NUMBERS_DUPLICATE_SHEET")) {
                    diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                        "IWORK_NUMBERS_DUPLICATE_SHEET",
                        "The Numbers document references the same sheet more than once; editable reconstruction is incomplete.",
                        document.EntryPath, document.Identifier));
                }
                continue;
            }
            if (sheetRecord.MessageType != SheetArchive) {
                omittedUnits.Add(new IWorkObjectIdentity(sheetRecord));
                supportsEditableReconstruction = false;
                diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                    "IWORK_NUMBERS_SHEET_TYPE_UNSUPPORTED",
                    "The Numbers document references an object that is not a supported sheet; editable reconstruction is incomplete.",
                    sheetRecord.EntryPath, sheetRecord.Identifier));
                continue;
            }
            int drawableReferenceCount;
            try {
                drawableReferenceCount = IWorkProtobuf.CountFields(
                    sheetRecord.Payload, 2, projectionBudget.MaximumProtobufFieldCount);
            } catch (InvalidDataException exception)
                when (!IWorkProtobuf.IsFieldLimitException(exception)) {
                supportsEditableReconstruction = false;
                diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                    "IWORK_NUMBERS_SHEET_MALFORMED",
                    "A Numbers sheet is malformed; editable reconstruction is incomplete.",
                    sheetRecord.EntryPath, sheetRecord.Identifier));
                omittedUnits.Add(new IWorkObjectIdentity(sheetRecord));
                continue;
            }
            IWorkWireMessage sheetMessage = index.Message(sheetRecord);
            projectionBudget.AddDrawableReferences(drawableReferenceCount);
            var tables = new List<IWorkTable>();
            var textBoxes = new List<string>();
            var orderedDrawables = new List<IWorkNumbersDrawable>();
            IReadOnlyList<IWorkArchiveRecord> drawables = references.ReadAll(
                sheetRecord, sheetMessage, 2, out int unresolvedDrawableCount);
            if (unresolvedDrawableCount > 0) {
                supportsEditableReconstruction = false;
                diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                    "IWORK_NUMBERS_DRAWABLE_UNSUPPORTED",
                    "A Numbers sheet references a missing drawable; editable reconstruction is incomplete.",
                    sheetRecord.EntryPath, sheetRecord.Identifier));
            }
            foreach (IWorkArchiveRecord drawable in drawables) {
                if (!projectedDrawableIdentifiers.Add(drawable.Identifier)) {
                    supportsEditableReconstruction = false;
                    if (!diagnostics.Any(diagnostic => diagnostic.Code == "IWORK_NUMBERS_DUPLICATE_DRAWABLE")) {
                        diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                            "IWORK_NUMBERS_DUPLICATE_DRAWABLE",
                            "The Numbers document references the same drawable more than once; editable reconstruction is incomplete.",
                            sheetRecord.EntryPath, sheetRecord.Identifier));
                    }
                    continue;
                }
                if (drawable.MessageType == TableInfoArchive) {
                    projectionBudget.AddTable();
                    IWorkTable? table = IWorkTableReader.Read(source, drawable, projectionBudget, diagnostics,
                        ref materializedCellCount, ref supportsEditableReconstruction);
                    if (table != null) {
                        tables.Add(table);
                        orderedDrawables.Add(new IWorkNumbersDrawable(table));
                    } else omittedUnits.Add(new IWorkObjectIdentity(drawable));
                } else if (drawable.MessageType == TextShapeArchive) {
                    IWorkWireMessage? drawableMessage = IWorkDrawingReader.DrawableMessage(index, drawable,
                        out bool drawableComplete);
                    IWorkWireMessage? storageOwner = null;
                    try {
                        storageOwner = index.Message(drawable);
                    } catch (InvalidDataException) {
                        drawableComplete = false;
                    }
                    bool geometryComplete = true;
                    if (drawableMessage != null) {
                        IWorkDrawingReader.ReadGeometry(drawableMessage, out geometryComplete,
                            requirePositiveSize: true);
                    }
                    bool metadataComplete = true;
                    string? hyperlink = IWorkDrawingReader.ReadOptionalString(drawableMessage, 4,
                        projectionBudget, ref metadataComplete);
                    string? accessibilityDescription = IWorkDrawingReader.ReadOptionalString(
                        drawableMessage, 8, projectionBudget, ref metadataComplete);
                    bool storageReferenceComplete = storageOwner != null
                        && storageOwner.FieldCount(2) == 1
                        && !storageOwner.HasUnexpectedWireKind(2, IWorkWireKind.Bytes);
                    IWorkArchiveRecord? storage = storageOwner != null
                        ? references.ReadOne(drawable, storageOwner, 2)
                        : null;
                    if (storageReferenceComplete && storage != null && storage.MessageType == TextStorageArchive) {
                        string text;
                        bool textComplete;
                        try {
                            text = IWorkPagesReader.StorageText(index.Message(storage), projectionBudget,
                                out textComplete);
                        } catch (InvalidDataException) {
                            text = string.Empty;
                            textComplete = false;
                        }
                        if (!textComplete) {
                            supportsEditableReconstruction = false;
                            if (!diagnostics.Any(diagnostic => diagnostic.Code == "IWORK_NUMBERS_TEXT_STORAGE_UNSUPPORTED")) {
                                diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                                    "IWORK_NUMBERS_TEXT_STORAGE_UNSUPPORTED",
                                    "A Numbers text storage contains an invalid UTF-8 run; editable reconstruction is incomplete.",
                                    storage.EntryPath, storage.Identifier));
                            }
                        }
                        if (text.Length > 0 || hyperlink != null || accessibilityDescription != null) {
                            if (!drawableComplete || !geometryComplete || !metadataComplete || hyperlink != null
                                || accessibilityDescription != null) {
                                supportsEditableReconstruction = false;
                                if (!diagnostics.Any(diagnostic =>
                                        diagnostic.Code == "IWORK_NUMBERS_TEXT_METADATA_UNSUPPORTED")) {
                                    diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                                        "IWORK_NUMBERS_TEXT_METADATA_UNSUPPORTED",
                                        "A Numbers text shape contains malformed or unsupported drawable metadata; editable reconstruction is incomplete.",
                                        drawable.EntryPath, drawable.Identifier));
                                }
                            }
                        }
                        if (text.Length > 0) {
                            projectionBudget.AddTextItem();
                            textBoxes.Add(text);
                            orderedDrawables.Add(new IWorkNumbersDrawable(text, new IWorkObjectIdentity(storage)));
                        } else if (!textComplete) omittedUnits.Add(new IWorkObjectIdentity(storage));
                    } else {
                        supportsEditableReconstruction = false;
                        diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                            "IWORK_NUMBERS_TEXT_STORAGE_UNSUPPORTED",
                            "A Numbers text shape does not reference supported text storage; editable reconstruction is incomplete.",
                            drawable.EntryPath, drawable.Identifier));
                    }
                } else {
                    omittedUnits.Add(new IWorkObjectIdentity(drawable));
                    supportsEditableReconstruction = false;
                    if (!diagnostics.Any(diagnostic => diagnostic.Code == "IWORK_NUMBERS_DRAWABLE_UNSUPPORTED")) {
                        diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                            "IWORK_NUMBERS_DRAWABLE_UNSUPPORTED",
                            "A Numbers sheet contains an unsupported drawable; editable reconstruction is incomplete.",
                            drawable.EntryPath, drawable.Identifier));
                    }
                }
            }
            string? sheetName = sheetMessage.GetString(1, out bool sheetNameComplete);
            if (!sheetNameComplete) {
                MarkTextMetadataUnsupported(sheetRecord, diagnostics, ref supportsEditableReconstruction);
            }
            if (sheetName != null) projectionBudget.AddTextCharacters(sheetName.Length);
            sheets.Add(new IWorkNumbersSheet(sheetName ?? string.Empty, tables, textBoxes,
                orderedDrawables, new IWorkObjectIdentity(sheetRecord)));
        }
        if (sheets.Count == 0) {
            supportsEditableReconstruction = false;
            if (!diagnostics.Any(diagnostic => diagnostic.Code == "IWORK_NUMBERS_SHEET_MISSING")) {
                diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                    "IWORK_NUMBERS_SHEET_MISSING",
                    "No supported Numbers sheet was resolved; editable reconstruction is unavailable.",
                    document.EntryPath, document.Identifier));
            }
        }
        return new IWorkNumbersProjection(source, sheets, diagnostics, supportsEditableReconstruction,
            new IWorkObjectIdentity(document), omittedUnits, references.Issues);
    }

    private static void MarkTextMetadataUnsupported(IWorkArchiveRecord record,
        List<IWorkDiagnostic> diagnostics, ref bool supportsEditableReconstruction) {
        supportsEditableReconstruction = false;
        if (diagnostics.Any(diagnostic => diagnostic.Code == "IWORK_NUMBERS_TEXT_UNSUPPORTED")) return;
        diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
            "IWORK_NUMBERS_TEXT_UNSUPPORTED",
            "Numbers text metadata contains invalid Unicode content; editable reconstruction is incomplete.",
            record.EntryPath, record.Identifier));
    }

}
