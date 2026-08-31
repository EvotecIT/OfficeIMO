using OfficeIMO.IWork.Internal;
using System.Globalization;
using System.Numerics;

namespace OfficeIMO.IWork;

/// <summary>One Numbers sheet and its semantic drawables.</summary>
public sealed class IWorkNumbersSheet {
    internal IWorkNumbersSheet(string name, IReadOnlyList<IWorkTable> tables,
        IReadOnlyList<string> textBoxes) {
        Name = name;
        Tables = Array.AsReadOnly(tables.ToArray());
        TextBoxes = Array.AsReadOnly(textBoxes.ToArray());
    }

    /// <summary>Gets the source sheet name.</summary>
    public string Name { get; }
    /// <summary>Gets tables in drawable order.</summary>
    public IReadOnlyList<IWorkTable> Tables { get; }
    /// <summary>Gets text-box content in drawable order.</summary>
    public IReadOnlyList<string> TextBoxes { get; }
}

/// <summary>Read-only Numbers structure recovered from a shared IWA object graph.</summary>
public sealed class IWorkNumbersProjection {
    private readonly IWorkSourceDocument _source;
    private readonly bool _supportsEditableReconstruction;

    internal IWorkNumbersProjection(IWorkSourceDocument source, IReadOnlyList<IWorkNumbersSheet> sheets,
        IReadOnlyList<IWorkDiagnostic> diagnostics, bool supportsEditableReconstruction) {
        _source = source;
        Sheets = Array.AsReadOnly(sheets.ToArray());
        Diagnostics = Array.AsReadOnly(diagnostics.ToArray());
        _supportsEditableReconstruction = supportsEditableReconstruction;
    }

    /// <summary>Gets sheets in source order.</summary>
    public IReadOnlyList<IWorkNumbersSheet> Sheets { get; }
    /// <summary>Gets projection diagnostics.</summary>
    public IReadOnlyList<IWorkDiagnostic> Diagnostics { get; }
    /// <summary>Gets whether at least one editable sheet was recovered and its required semantic references were resolved.</summary>
    public bool HasEditableContent => Sheets.Count > 0 && _supportsEditableReconstruction;

    /// <summary>Creates an import report for an OfficeIMO semantic-owner projection.</summary>
    public IWorkImportReport CreateImportReport(IWorkProjectionKind kind, IWorkPreviewAsset? preview = null) {
        ValidateReportRequest(kind, preview);
        return _source.CreateReport(kind, Diagnostics, preview,
            kind == IWorkProjectionKind.VisualFallback
                ? 0
                : Sheets.Count + Sheets.Sum(sheet => sheet.TextBoxes.Count + sheet.Tables.Count
                    + sheet.Tables.Sum(table => table.Cells.Count)));
    }

    private void ValidateReportRequest(IWorkProjectionKind kind, IWorkPreviewAsset? preview) {
        if (kind == IWorkProjectionKind.EditableReconstruction && !HasEditableContent) {
            throw new InvalidOperationException("Editable Numbers content was not recovered.");
        }
        if (kind == IWorkProjectionKind.VisualFallback && preview == null) {
            throw new ArgumentNullException(nameof(preview), "A visual fallback report requires the preview used by the owner.");
        }
    }
}

public sealed partial class IWorkSourceDocument {
    /// <summary>Reads a Numbers package into a bounded semantic source projection, or returns a diagnostic-only projection in visual-only mode.</summary>
    public IWorkNumbersProjection ReadNumbers() {
        if (Kind != IWorkDocumentKind.Numbers) throw new InvalidOperationException($"The source is {Kind}, not Numbers.");
        if (RequestedImportMode == IWorkImportMode.VisualOnly) {
            return new IWorkNumbersProjection(this, Array.Empty<IWorkNumbersSheet>(),
                new[] { IWorkProjectionDiagnostics.SemanticProjectionSkipped }, supportsEditableReconstruction: false);
        }
        return IWorkNumbersReader.Read(this);
    }
}

internal static class IWorkNumbersReader {
    private const uint DocumentArchive = 1;
    private const uint SheetArchive = 2;
    private const uint TableInfoArchive = 6000;
    private const uint WordProcessingTableInfoArchive = 6007;
    private const uint TableModelArchive = 6001;
    private const uint TableTileArchive = 6002;
    private const uint TextStorageArchive = 2001;
    private const uint TextShapeArchive = 2011;
    private const int TileRowStride = 256;
    private const int MaximumTileMetadataFields = 6;
    private const uint RecognizedCellValueMask = (1u << 21) - 1;

    internal static IWorkNumbersProjection Read(IWorkSourceDocument source) {
        var diagnostics = new List<IWorkDiagnostic>();
        var sheets = new List<IWorkNumbersSheet>();
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
        int declaredSheetCount = IWorkProtobuf.CountFields(document.Payload, 1,
            source.Options.MaximumProtobufFieldCount);
        if (declaredSheetCount > source.Options.MaximumProjectedSheets) {
            throw new InvalidDataException($"Numbers sheet count exceeds the configured projection limit of {source.Options.MaximumProjectedSheets}.");
        }
        IReadOnlyList<IWorkArchiveRecord> sheetRecords = index.DereferenceAll(
            index.Message(document), 1, out int unresolvedSheetCount);
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
                supportsEditableReconstruction = false;
                diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                    "IWORK_NUMBERS_SHEET_TYPE_UNSUPPORTED",
                    "The Numbers document references an object that is not a supported sheet; editable reconstruction is incomplete.",
                    sheetRecord.EntryPath, sheetRecord.Identifier));
                continue;
            }
            projectionBudget.AddDrawableReferences(IWorkProtobuf.CountFields(
                sheetRecord.Payload, 2, projectionBudget.MaximumProtobufFieldCount));
            IWorkWireMessage sheetMessage = index.Message(sheetRecord);
            var tables = new List<IWorkTable>();
            var textBoxes = new List<string>();
            IReadOnlyList<IWorkArchiveRecord> drawables = index.DereferenceAll(
                sheetMessage, 2, out int unresolvedDrawableCount);
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
                    if (table != null) tables.Add(table);
                } else if (drawable.MessageType == TextShapeArchive) {
                    IWorkWireMessage? drawableMessage = IWorkDrawingReader.DrawableMessage(index, drawable,
                        out bool drawableComplete);
                    bool geometryComplete = true;
                    if (drawableMessage != null) {
                        IWorkDrawingReader.ReadGeometry(drawableMessage, out geometryComplete);
                    }
                    bool metadataComplete = true;
                    string? hyperlink = IWorkDrawingReader.ReadOptionalString(drawableMessage, 4,
                        projectionBudget, ref metadataComplete);
                    string? accessibilityDescription = IWorkDrawingReader.ReadOptionalString(
                        drawableMessage, 8, projectionBudget, ref metadataComplete);
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
                    IWorkWireMessage storageOwner = index.Message(drawable);
                    bool storageReferenceComplete = storageOwner.FieldCount(2) == 1
                        && !storageOwner.HasUnexpectedWireKind(2, IWorkWireKind.Bytes);
                    IWorkArchiveRecord? storage = storageReferenceComplete
                        ? index.Dereference(storageOwner, 2)
                        : null;
                    if (storage != null && storage.MessageType == TextStorageArchive) {
                        string text = IWorkPagesReader.StorageText(index.Message(storage), projectionBudget,
                            out bool textComplete);
                        if (!textComplete) {
                            supportsEditableReconstruction = false;
                            if (!diagnostics.Any(diagnostic => diagnostic.Code == "IWORK_NUMBERS_TEXT_STORAGE_UNSUPPORTED")) {
                                diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                                    "IWORK_NUMBERS_TEXT_STORAGE_UNSUPPORTED",
                                    "A Numbers text storage contains an invalid UTF-8 run; editable reconstruction is incomplete.",
                                    storage.EntryPath, storage.Identifier));
                            }
                        }
                        if (text.Length > 0) {
                            projectionBudget.AddTextItem();
                            textBoxes.Add(text);
                        }
                    } else {
                        supportsEditableReconstruction = false;
                        diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                            "IWORK_NUMBERS_TEXT_STORAGE_UNSUPPORTED",
                            "A Numbers text shape does not reference supported text storage; editable reconstruction is incomplete.",
                            drawable.EntryPath, drawable.Identifier));
                    }
                } else {
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
            sheets.Add(new IWorkNumbersSheet(sheetName ?? string.Empty, tables, textBoxes));
        }
        return new IWorkNumbersProjection(source, sheets, diagnostics, supportsEditableReconstruction);
    }

    internal static IWorkTable? ReadTableInfo(IWorkSourceDocument source, IWorkArchiveRecord tableRecord,
        IWorkProjectionBudget projectionBudget, List<IWorkDiagnostic> diagnostics,
        ref int materializedCellCount,
        ref bool supportsEditableReconstruction) {
        IWorkWireMessage recordMessage = source.Index.Message(tableRecord);
        IWorkWireMessage? tableInfo = tableRecord.MessageType switch {
            TableInfoArchive => recordMessage,
            WordProcessingTableInfoArchive => IWorkObjectIndex.TryGetMessage(recordMessage, 1),
            _ => null
        };
        if (tableInfo == null) {
            supportsEditableReconstruction = false;
            diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                "IWORK_TABLE_INFO_UNSUPPORTED",
                "An iWork table has no supported table-info payload; editable reconstruction is incomplete.",
                tableRecord.EntryPath, tableRecord.Identifier));
            return null;
        }
        bool modelReferenceComplete = tableInfo.FieldCount(2) == 1
            && !tableInfo.HasUnexpectedWireKind(2, IWorkWireKind.Bytes);
        IWorkArchiveRecord? model = modelReferenceComplete
            ? source.Index.Dereference(tableInfo, 2)
            : null;
        if (model == null || model.MessageType != TableModelArchive) {
            supportsEditableReconstruction = false;
            diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                "IWORK_TABLE_MODEL_UNSUPPORTED",
                "An iWork table does not reference a supported table model; editable reconstruction is incomplete.",
                tableRecord.EntryPath, tableRecord.Identifier));
            return null;
        }
        IWorkWireMessage? drawable = IWorkObjectIndex.TryGetMessage(tableInfo, 1, out bool malformedDrawable);
        if (malformedDrawable || tableInfo.HasUnexpectedWireKind(1, IWorkWireKind.Bytes)
            || tableInfo.HasField(1) && drawable == null) {
            supportsEditableReconstruction = false;
            diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                "IWORK_TABLE_DRAWABLE_UNSUPPORTED",
                "An iWork table contains malformed drawable geometry; editable reconstruction is incomplete.",
                tableRecord.EntryPath, tableRecord.Identifier));
        }
        bool geometryComplete = true;
        IWorkGeometry? geometry = drawable == null
            ? null
            : IWorkDrawingReader.ReadGeometry(drawable, out geometryComplete);
        if (drawable != null && !geometryComplete) {
            supportsEditableReconstruction = false;
            if (!diagnostics.Any(diagnostic => diagnostic.Code == "IWORK_TABLE_DRAWABLE_UNSUPPORTED"
                    && diagnostic.RecordIdentifier == tableRecord.Identifier)) {
                diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                    "IWORK_TABLE_DRAWABLE_UNSUPPORTED",
                    "An iWork table contains malformed drawable geometry; editable reconstruction is incomplete.",
                    tableRecord.EntryPath, tableRecord.Identifier));
            }
        }
        bool metadataComplete = true;
        string? hyperlink = IWorkDrawingReader.ReadOptionalString(drawable, 4,
            projectionBudget, ref metadataComplete);
        string? accessibilityDescription = IWorkDrawingReader.ReadOptionalString(drawable, 8,
            projectionBudget, ref metadataComplete);
        if (!metadataComplete) {
            supportsEditableReconstruction = false;
            if (!diagnostics.Any(diagnostic => diagnostic.Code == "IWORK_TABLE_DRAWABLE_UNSUPPORTED"
                    && diagnostic.RecordIdentifier == tableRecord.Identifier)) {
                diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                    "IWORK_TABLE_DRAWABLE_UNSUPPORTED",
                    "An iWork table contains malformed drawable metadata; editable reconstruction is incomplete.",
                    tableRecord.EntryPath, tableRecord.Identifier));
            }
        }
        if (hyperlink != null) {
            supportsEditableReconstruction = false;
            diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                "IWORK_TABLE_HYPERLINK_UNSUPPORTED",
                "An iWork table contains a drawable hyperlink that is preserved but cannot be represented by the editable table owners.",
                tableRecord.EntryPath, tableRecord.Identifier));
        }
        if (accessibilityDescription != null && source.Kind == IWorkDocumentKind.Numbers) {
            supportsEditableReconstruction = false;
            diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                "IWORK_NUMBERS_TABLE_ACCESSIBILITY_UNSUPPORTED",
                "A Numbers table accessibility description is preserved but cannot be represented by the editable worksheet projection.",
                tableRecord.EntryPath, tableRecord.Identifier));
        }
        return ReadTable(source, source.Index, model, geometry, projectionBudget, diagnostics,
            accessibilityDescription, ref materializedCellCount, ref supportsEditableReconstruction);
    }

    private static IWorkTable ReadTable(IWorkSourceDocument source, IWorkObjectIndex index,
        IWorkArchiveRecord model, IWorkGeometry? geometry, IWorkProjectionBudget projectionBudget,
        List<IWorkDiagnostic> diagnostics,
        string? accessibilityDescription,
        ref int materializedCellCount, ref bool supportsEditableReconstruction) {
        IWorkWireMessage message = index.Message(model);
        if (HasUnsupportedTableScalarEncoding(message)) {
            supportsEditableReconstruction = false;
            diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                "IWORK_TABLE_DIMENSIONS_UNSUPPORTED",
                "An iWork table declares dimensions or default sizing with an unsupported wire encoding; editable reconstruction is incomplete.",
                model.EntryPath, model.Identifier));
        }
        int rows = CheckedDimension(message.GetUnsigned(6), source.Options.MaximumTableRows, "row", model);
        int columns = CheckedDimension(message.GetUnsigned(7), source.Options.MaximumTableColumns, "column", model);
        string? tableName = message.GetString(8, out bool tableNameComplete);
        if (!tableNameComplete) {
            MarkTextMetadataUnsupported(model, diagnostics, ref supportsEditableReconstruction);
        }
        if (tableName != null) projectionBudget.AddTextCharacters(tableName.Length);
        string name = tableName ?? string.Empty;
        int headerRows = CheckedSubDimension(message.GetUnsigned(9), rows, "header row", model);
        int headerColumns = CheckedSubDimension(message.GetUnsigned(10), columns, "header column", model);
        int footerRows = CheckedSubDimension(message.GetUnsigned(11), rows, "footer row", model);
        if ((long)headerRows + footerRows > rows) {
            supportsEditableReconstruction = false;
            diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                "IWORK_TABLE_REGIONS_UNSUPPORTED",
                "An iWork table declares overlapping header and footer row regions; editable reconstruction is incomplete.",
                model.EntryPath, model.Identifier));
        }
        double? declaredRowHeight = message.GetDouble(16);
        double? declaredColumnWidth = message.GetDouble(17);
        bool invalidDeclaredSizing = HasInvalidDeclaredDimension(message, 16, declaredRowHeight)
            || HasInvalidDeclaredDimension(message, 17, declaredColumnWidth);
        if (invalidDeclaredSizing) {
            supportsEditableReconstruction = false;
            diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                "IWORK_TABLE_DIMENSIONS_UNSUPPORTED",
                "An iWork table declares a non-positive or non-finite default size; editable reconstruction is incomplete.",
                model.EntryPath, model.Identifier));
        }
        double? defaultRowHeight = ValidDimension(declaredRowHeight);
        double? defaultColumnWidth = ValidDimension(declaredColumnWidth);
        IReadOnlyList<IWorkTableMergeRange> mergedRanges = ReadMergedRanges(message, rows, columns,
            source.Options.MaximumTableMergedRanges, source.Options.MaximumFormulaNodes,
            model, diagnostics, ref supportsEditableReconstruction);
        var cells = new List<IWorkTableCell>();
        var coordinates = new HashSet<long>();
        IWorkWireMessage? store = IWorkObjectIndex.TryGetMessage(message, 4);
        if (store == null) {
            MarkTableStorageUnsupported(model, diagnostics, ref supportsEditableReconstruction);
            return CreateTable();
        }

        int remainingMaterializedCells = source.Options.MaximumMaterializedCells - materializedCellCount;
        IReadOnlyDictionary<uint, string> strings = ReadStrings(index, store,
            projectionBudget, source.Options, remainingMaterializedCells,
            out bool stringStorageComplete);
        IReadOnlyDictionary<uint, IWorkWireMessage> formulas = ReadFormulas(index, store,
            source.Options, remainingMaterializedCells, out bool formulaStorageComplete);
        if (!stringStorageComplete) {
            supportsEditableReconstruction = false;
            diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                "IWORK_TABLE_STRING_STORAGE_UNSUPPORTED",
                "An iWork string table contains malformed or duplicate entries; editable reconstruction is incomplete.",
                model.EntryPath, model.Identifier));
        }
        if (!formulaStorageComplete) {
            diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                "IWORK_TABLE_FORMULA_STORAGE_UNSUPPORTED",
                "An iWork formula table contains malformed or duplicate entries; affected formulas retain cached values only.",
                model.EntryPath, model.Identifier));
        }
        byte[]? tileStorageBytes = store.GetBytes(3);
        int declaredTileCount;
        try {
            declaredTileCount = tileStorageBytes == null
                || store.FieldCount(3) != 1
                || store.HasUnexpectedWireKind(3, IWorkWireKind.Bytes)
                    ? -1
                    : IWorkProtobuf.CountFields(tileStorageBytes, 1,
                        source.Options.MaximumProtobufFieldCount);
        } catch (InvalidDataException) {
            declaredTileCount = -1;
        }
        int maximumTileCount = checked((rows + TileRowStride - 1) / TileRowStride);
        if (declaredTileCount > maximumTileCount) {
            supportsEditableReconstruction = false;
            diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                "IWORK_TABLE_TILE_COUNT_UNSUPPORTED",
                "An iWork table declares more tiles than can fit in its logical row range; editable reconstruction is incomplete.",
                model.EntryPath, model.Identifier));
            return CreateTable();
        }
        IWorkWireMessage? tileStorage;
        try {
            tileStorage = declaredTileCount < 0
                ? null
                : store.ParseNestedMessage(tileStorageBytes!);
        } catch (InvalidDataException) {
            tileStorage = null;
        }
        if (tileStorage == null) {
            MarkTableStorageUnsupported(model, diagnostics, ref supportsEditableReconstruction);
            return CreateTable();
        }
        IReadOnlyList<IWorkWireMessage> tileEntries = IWorkObjectIndex.TryGetMessages(
            tileStorage, 1, out bool malformedTileEntries);
        if (malformedTileEntries) {
            MarkTableStorageUnsupported(model, diagnostics, ref supportsEditableReconstruction);
            return CreateTable();
        }
        var tileIndexes = new HashSet<ulong>();
        var tileIdentifiers = new HashSet<ulong>();
        foreach (IWorkWireMessage tileEntry in tileEntries) {
            ulong? declaredTileId = tileEntry.GetUnsigned(1);
            if (!declaredTileId.HasValue || declaredTileId.Value > int.MaxValue) {
                supportsEditableReconstruction = false;
                diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                    "IWORK_TABLE_TILE_INDEX_UNSUPPORTED",
                    "An iWork table tile index is missing or exceeds the supported range; editable reconstruction is incomplete.",
                    model.EntryPath, model.Identifier));
                continue;
            }
            ulong rawTileId = declaredTileId.Value;
            if (!tileIndexes.Add(rawTileId)) {
                MarkDuplicateTile(model, diagnostics, ref supportsEditableReconstruction);
                continue;
            }
            IWorkArchiveRecord? tile = index.Dereference(tileEntry, 2);
            if (tile == null || tile.MessageType != TableTileArchive) {
                supportsEditableReconstruction = false;
                diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                    "IWORK_TABLE_TILE_UNSUPPORTED",
                    "An iWork table references a missing or unsupported tile object; editable reconstruction is incomplete.",
                    model.EntryPath, model.Identifier));
                continue;
            }
            if (!tileIdentifiers.Add(tile.Identifier)) {
                MarkDuplicateTile(model, diagnostics, ref supportsEditableReconstruction);
                continue;
            }
            long tileStartRow = checked((long)rawTileId * TileRowStride);
            long remainingRows = rows - tileStartRow;
            int maximumRowsInTile = remainingRows <= 0
                ? 0
                : (int)Math.Min(TileRowStride, remainingRows);
            int declaredRowsInTile = IWorkProtobuf.CountFields(tile.Payload, 5,
                source.Options.MaximumProtobufFieldCount, out int totalTileFieldCount);
            if (totalTileFieldCount - declaredRowsInTile > MaximumTileMetadataFields) {
                supportsEditableReconstruction = false;
                diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                    "IWORK_TABLE_TILE_FIELDS_UNSUPPORTED",
                    "An iWork table tile contains more metadata fields than the supported tile envelope; editable reconstruction is incomplete.",
                    tile.EntryPath, tile.Identifier));
                continue;
            }
            if (declaredRowsInTile > maximumRowsInTile) {
                supportsEditableReconstruction = false;
                diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                    "IWORK_TABLE_TILE_ROW_COUNT_UNSUPPORTED",
                    "An iWork table tile declares more row messages than can fit in its logical table range; editable reconstruction is incomplete.",
                    tile.EntryPath, tile.Identifier));
                continue;
            }
            IReadOnlyList<IWorkWireMessage> rowsInTile = IWorkObjectIndex.TryGetMessages(
                index.Message(tile), 5, out bool malformedRows);
            if (malformedRows) {
                supportsEditableReconstruction = false;
                diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                    "IWORK_TABLE_TILE_ROWS_UNSUPPORTED",
                    "An iWork table tile contains malformed row metadata; editable reconstruction is incomplete.",
                    tile.EntryPath, tile.Identifier));
            }
            var rowIndexes = new HashSet<ulong>();
            foreach (IWorkWireMessage rowInfo in rowsInTile) {
                byte[]? currentBuffer = rowInfo.GetBytes(6);
                byte[]? currentOffsets = rowInfo.GetBytes(7);
                if (rowInfo.HasUnexpectedWireKind(1, IWorkWireKind.Varint)
                    || rowInfo.HasUnexpectedWireKind(3, IWorkWireKind.Bytes)
                    || rowInfo.HasUnexpectedWireKind(4, IWorkWireKind.Bytes)
                    || rowInfo.HasUnexpectedWireKind(6, IWorkWireKind.Bytes)
                    || rowInfo.HasUnexpectedWireKind(7, IWorkWireKind.Bytes)
                    || rowInfo.HasUnexpectedWireKind(8, IWorkWireKind.Varint)
                    || rowInfo.FieldCount(3) > 1
                    || rowInfo.FieldCount(4) > 1
                    || rowInfo.FieldCount(6) > 1
                    || rowInfo.FieldCount(7) > 1
                    || (currentBuffer == null) != (currentOffsets == null)
                    || currentOffsets != null && currentOffsets.Length % 2 != 0) {
                    MarkCellStorageUnsupported(tile, diagnostics, ref supportsEditableReconstruction);
                    continue;
                }
                bool hasPreBncStorage = (rowInfo.GetBytes(3)?.Length ?? 0) > 0
                    || (rowInfo.GetBytes(4)?.Length ?? 0) > 0;
                if ((currentBuffer == null || currentOffsets == null) && hasPreBncStorage) {
                    supportsEditableReconstruction = false;
                    if (!diagnostics.Any(diagnostic => diagnostic.Code == "IWORK_TABLE_LEGACY_CELL_STORAGE")) {
                        diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                            "IWORK_TABLE_LEGACY_CELL_STORAGE",
                            "The source uses pre-BNC iWork table cell storage. Records are preserved, but editable reconstruction is unavailable.",
                            tile.EntryPath, tile.Identifier));
                    }
                    continue;
                }
                ulong? declaredRow = rowInfo.GetUnsigned(1);
                if (!declaredRow.HasValue || declaredRow.Value >= TileRowStride
                    || !rowIndexes.Add(declaredRow.Value)) {
                    supportsEditableReconstruction = false;
                    if (!diagnostics.Any(diagnostic => diagnostic.Code == "IWORK_TABLE_TILE_ROW_UNSUPPORTED")) {
                        diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                            "IWORK_TABLE_TILE_ROW_UNSUPPORTED",
                            "An iWork table tile contains a missing, repeated, or out-of-range row index; editable reconstruction is incomplete.",
                            tile.EntryPath, tile.Identifier));
                    }
                    continue;
                }
                ulong rawRow = declaredRow.Value;
                long zeroBasedRow = checked((long)rawTileId * TileRowStride + (long)rawRow);
                if (zeroBasedRow < 0 || zeroBasedRow >= rows) {
                    supportsEditableReconstruction = false;
                    if (!diagnostics.Any(diagnostic => diagnostic.Code == "IWORK_TABLE_TILE_ROW_UNSUPPORTED")) {
                        diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                            "IWORK_TABLE_TILE_ROW_UNSUPPORTED",
                            "An iWork table tile contains a row outside the declared table bounds; editable reconstruction is incomplete.",
                            tile.EntryPath, tile.Identifier));
                    }
                    continue;
                }
                byte[] buffer = currentBuffer ?? Array.Empty<byte>();
                byte[] offsets = currentOffsets ?? Array.Empty<byte>();
                bool hasWideOffsets = (rowInfo.GetUnsigned(8) ?? 0) != 0;
                int offsetColumnCount = offsets.Length / 2;
                int availableColumns = Math.Min(columns, offsetColumnCount);
                bool hasPopulatedTrailingOffset = false;
                bool hasExcessiveTrailingOffsets = offsetColumnCount > TileRowStride;
                if (!hasExcessiveTrailingOffsets) {
                    for (int column = columns; column < offsetColumnCount; column++) {
                        int encodedOffset = offsets[column * 2] | offsets[column * 2 + 1] << 8;
                        if (encodedOffset != ushort.MaxValue) {
                            hasPopulatedTrailingOffset = true;
                            break;
                        }
                    }
                }
                if (hasExcessiveTrailingOffsets || hasPopulatedTrailingOffset) {
                    MarkCellStorageUnsupported(tile, diagnostics, ref supportsEditableReconstruction);
                }
                int[] populatedOffsets = Enumerable.Range(0, availableColumns)
                    .Select(column => offsets[column * 2] | offsets[column * 2 + 1] << 8)
                    .Where(encodedOffset => encodedOffset != ushort.MaxValue)
                    .Select(encodedOffset => hasWideOffsets ? checked(encodedOffset * 4) : encodedOffset)
                    .Distinct()
                    .OrderBy(cellOffset => cellOffset)
                    .ToArray();
                var cellLimits = new Dictionary<int, int>(populatedOffsets.Length);
                for (int offsetIndex = 0; offsetIndex < populatedOffsets.Length; offsetIndex++) {
                    cellLimits.Add(populatedOffsets[offsetIndex],
                        offsetIndex + 1 < populatedOffsets.Length
                            ? populatedOffsets[offsetIndex + 1]
                            : buffer.Length);
                }
                for (int column = 0; column < availableColumns; column++) {
                    int encodedOffset = offsets[column * 2] | offsets[column * 2 + 1] << 8;
                    if (encodedOffset == ushort.MaxValue) continue;
                    int offset = hasWideOffsets ? checked(encodedOffset * 4) : encodedOffset;
                    IWorkTableCell cell = DecodeCell(buffer, offset, cellLimits[offset],
                        checked((int)zeroBasedRow + 1), column + 1,
                        strings, formulas, source.Options, projectionBudget);
                    if (cell.Kind == IWorkCellKind.Empty) continue;
                    if (materializedCellCount >= source.Options.MaximumMaterializedCells) {
                        throw new InvalidDataException($"iWork cell count exceeds the configured source-wide limit of {source.Options.MaximumMaterializedCells}.");
                    }
                    long coordinate = ((long)cell.Row << 32) | (uint)cell.Column;
                    if (!coordinates.Add(coordinate)) {
                        supportsEditableReconstruction = false;
                        if (!diagnostics.Any(diagnostic => diagnostic.Code == "IWORK_TABLE_DUPLICATE_CELL")) {
                            diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                                "IWORK_TABLE_DUPLICATE_CELL",
                                "An iWork table defines more than one value for the same cell; editable reconstruction is incomplete.",
                                tile.EntryPath, tile.Identifier));
                        }
                        continue;
                    }
                    cells.Add(cell);
                    materializedCellCount++;
                }
            }
        }

        int errorCount = cells.Count(cell => cell.Kind == IWorkCellKind.Error && cell.Error != "#ERROR");
        if (errorCount > 0) {
            supportsEditableReconstruction = false;
            diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning, "IWORK_TABLE_CELL_DECODE",
                $"{errorCount} cells in table '{name}' could not be decoded completely.", model.EntryPath, model.Identifier));
        }
        int incompleteCachedFormulaCount = cells.Count(cell => cell.Kind == IWorkCellKind.Formula
            && !cell.FormulaIsComplete && cell.Value != null);
        if (incompleteCachedFormulaCount > 0) {
            diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning, "IWORK_TABLE_FORMULA_PARTIAL",
                $"{incompleteCachedFormulaCount} formulas in table '{name}' retain typed cached values because their expressions were not reconstructed completely.",
                model.EntryPath, model.Identifier));
        }
        int incompleteUncachedFormulaCount = cells.Count(cell => cell.Kind == IWorkCellKind.Formula
            && !cell.FormulaIsComplete && cell.Value == null);
        if (incompleteUncachedFormulaCount > 0) {
            supportsEditableReconstruction = false;
            diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                "IWORK_TABLE_FORMULA_UNSUPPORTED",
                $"{incompleteUncachedFormulaCount} formulas in table '{name}' have neither a complete expression nor a cached value; editable reconstruction is incomplete.",
                model.EntryPath, model.Identifier));
        }
        return CreateTable();

        IWorkTable CreateTable() => new(name, rows, columns, cells,
            headerRows, headerColumns, footerRows, defaultRowHeight, defaultColumnWidth,
            mergedRanges, geometry, accessibilityDescription);
    }

    private static void MarkDuplicateTile(IWorkArchiveRecord model,
        List<IWorkDiagnostic> diagnostics, ref bool supportsEditableReconstruction) {
        supportsEditableReconstruction = false;
        if (diagnostics.Any(diagnostic => diagnostic.Code == "IWORK_TABLE_DUPLICATE_TILE")) return;
        diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
            "IWORK_TABLE_DUPLICATE_TILE",
            "An iWork table repeats a logical or physical tile; editable reconstruction is incomplete.",
            model.EntryPath, model.Identifier));
    }

    private static void MarkTableStorageUnsupported(IWorkArchiveRecord model,
        List<IWorkDiagnostic> diagnostics, ref bool supportsEditableReconstruction) {
        supportsEditableReconstruction = false;
        diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
            "IWORK_TABLE_STORAGE_UNSUPPORTED",
            "An iWork table has no supported tile storage; editable reconstruction is incomplete.",
            model.EntryPath, model.Identifier));
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

    private static IReadOnlyDictionary<uint, string> ReadStrings(IWorkObjectIndex index,
        IWorkWireMessage store, IWorkProjectionBudget projectionBudget,
        IWorkReadOptions options, int maximumEntries,
        out bool fullyReconstructed) {
        var strings = new Dictionary<uint, string>();
        fullyReconstructed = true;
        IWorkArchiveRecord? list = index.Dereference(store, 4);
        if (list == null) return strings;
        EnsureCatalogEntryCount(list, maximumEntries, options, "string");
        IReadOnlyList<IWorkWireMessage> entries = IWorkObjectIndex.TryGetMessages(
            index.Message(list), 3, out bool malformedEntries);
        if (malformedEntries) fullyReconstructed = false;
        foreach (IWorkWireMessage entry in entries) {
            ulong? key = entry.GetUnsigned(1);
            string? value = entry.GetString(3);
            if (!key.HasValue || key.Value > uint.MaxValue || value == null) {
                fullyReconstructed = false;
                continue;
            }
            projectionBudget.AddTextCharacters(value.Length);
            uint normalizedKey = (uint)key.Value;
            if (strings.ContainsKey(normalizedKey)) fullyReconstructed = false;
            else strings.Add(normalizedKey, value);
        }
        return strings;
    }

    private static IReadOnlyDictionary<uint, IWorkWireMessage> ReadFormulas(IWorkObjectIndex index,
        IWorkWireMessage store, IWorkReadOptions options, int maximumEntries,
        out bool fullyReconstructed) {
        var formulas = new Dictionary<uint, IWorkWireMessage>();
        var ambiguousIdentifiers = new HashSet<uint>();
        fullyReconstructed = true;
        IWorkArchiveRecord? list = index.Dereference(store, 6);
        if (list == null) return formulas;
        EnsureCatalogEntryCount(list, maximumEntries, options, "formula");
        IReadOnlyList<IWorkWireMessage> entries = IWorkObjectIndex.TryGetMessages(
            index.Message(list), 3, out bool malformedEntries);
        if (malformedEntries) fullyReconstructed = false;
        foreach (IWorkWireMessage entry in entries) {
            ulong? key = entry.GetUnsigned(1);
            IWorkWireMessage? formula = IWorkObjectIndex.TryGetMessage(entry, 5, out bool malformedFormula);
            if (!key.HasValue || key.Value > uint.MaxValue || malformedFormula || formula == null) {
                fullyReconstructed = false;
                continue;
            }
            uint normalizedKey = (uint)key.Value;
            if (ambiguousIdentifiers.Contains(normalizedKey)) {
                fullyReconstructed = false;
            } else if (formulas.ContainsKey(normalizedKey)) {
                formulas.Remove(normalizedKey);
                ambiguousIdentifiers.Add(normalizedKey);
                fullyReconstructed = false;
            } else {
                formulas.Add(normalizedKey, formula);
            }
        }
        return formulas;
    }

    private static void EnsureCatalogEntryCount(IWorkArchiveRecord list, int maximumEntries,
        IWorkReadOptions options, string catalogName) {
        int declaredEntryCount = IWorkProtobuf.CountFields(
            list.Payload, 3, options.MaximumProtobufFieldCount,
            out int totalFieldCount);
        int identifierFieldCount = IWorkProtobuf.CountFields(
            list.Payload, 1, options.MaximumProtobufFieldCount);
        int metadataFieldCount = IWorkProtobuf.CountFields(
            list.Payload, 2, options.MaximumProtobufFieldCount);
        if (identifierFieldCount > 1 || metadataFieldCount > 1
            || totalFieldCount - declaredEntryCount
                != identifierFieldCount + metadataFieldCount) {
            throw new InvalidDataException(
                $"An iWork {catalogName} catalog contains fields outside the supported entry envelope.");
        }
        if (declaredEntryCount > maximumEntries) {
            throw new InvalidDataException(
                $"An iWork {catalogName} catalog exceeds the remaining materialized-cell limit of {maximumEntries}.");
        }
    }

    private static bool HasUnsupportedTableScalarEncoding(IWorkWireMessage message) =>
        message.HasUnexpectedWireKind(6, IWorkWireKind.Varint)
        || message.HasUnexpectedWireKind(7, IWorkWireKind.Varint)
        || message.HasUnexpectedWireKind(9, IWorkWireKind.Varint)
        || message.HasUnexpectedWireKind(10, IWorkWireKind.Varint)
        || message.HasUnexpectedWireKind(11, IWorkWireKind.Varint)
        || message.HasUnexpectedWireKind(16, IWorkWireKind.Fixed64)
        || message.HasUnexpectedWireKind(17, IWorkWireKind.Fixed64);

    private static IReadOnlyList<IWorkTableMergeRange> ReadMergedRanges(IWorkWireMessage table,
        int rowCount, int columnCount, int maximumRanges, int maximumFormulaNodes,
        IWorkArchiveRecord model,
        List<IWorkDiagnostic> diagnostics,
        ref bool supportsEditableReconstruction) {
        IWorkWireMessage? mergeOwner = IWorkObjectIndex.TryGetMessage(table, 47, out bool malformedOwner);
        if (malformedOwner) {
            MarkMergeStorageUnsupported(model, diagnostics, ref supportsEditableReconstruction);
            return Array.Empty<IWorkTableMergeRange>();
        }
        if (mergeOwner == null) return Array.Empty<IWorkTableMergeRange>();
        byte[]? formulaStoreBytes = mergeOwner.GetBytes(2);
        int pairCount;
        try {
            pairCount = formulaStoreBytes == null
                || mergeOwner.FieldCount(2) != 1
                || mergeOwner.HasUnexpectedWireKind(2, IWorkWireKind.Bytes)
                    ? -1
                    : mergeOwner.CountNestedFields(formulaStoreBytes, 3);
        } catch (InvalidDataException) {
            pairCount = -1;
        }
        if (pairCount > maximumRanges) {
            throw new InvalidDataException($"iWork table merged-range count exceeds the configured limit of {maximumRanges} in object {model.Identifier}.");
        }
        if (pairCount < 0) {
            MarkMergeStorageUnsupported(model, diagnostics, ref supportsEditableReconstruction);
            return Array.Empty<IWorkTableMergeRange>();
        }
        IWorkWireMessage formulaStore;
        try {
            formulaStore = mergeOwner.ParseNestedMessage(formulaStoreBytes!);
        } catch (InvalidDataException) {
            MarkMergeStorageUnsupported(model, diagnostics, ref supportsEditableReconstruction);
            return Array.Empty<IWorkTableMergeRange>();
        }
        IReadOnlyList<IWorkWireMessage> pairs = IWorkObjectIndex.TryGetMessages(formulaStore, 3, out bool malformedPairs);
        if (malformedPairs) {
            MarkMergeStorageUnsupported(model, diagnostics, ref supportsEditableReconstruction);
            return Array.Empty<IWorkTableMergeRange>();
        }
        if (pairs.Count != pairCount) {
            MarkMergeStorageUnsupported(model, diagnostics, ref supportsEditableReconstruction);
            return Array.Empty<IWorkTableMergeRange>();
        }
        var result = new List<IWorkTableMergeRange>();
        var seen = new HashSet<string>(StringComparer.Ordinal);
        foreach (IWorkWireMessage pair in pairs) {
            IWorkWireMessage? formula = IWorkObjectIndex.TryGetMessage(pair, 2, out bool malformedFormula);
            if (malformedFormula || formula == null
                || !IWorkFormulaReader.TryReadAbsoluteRange(formula, maximumFormulaNodes,
                    out int firstRow, out int firstColumn, out int lastRow, out int lastColumn)
                || firstRow >= rowCount || lastRow >= rowCount
                || firstColumn >= columnCount || lastColumn >= columnCount) {
                MarkMergeStorageUnsupported(model, diagnostics, ref supportsEditableReconstruction);
                continue;
            }
            if (firstRow == lastRow && firstColumn == lastColumn) continue;
            string key = firstRow.ToString(CultureInfo.InvariantCulture) + ":"
                + firstColumn.ToString(CultureInfo.InvariantCulture) + ":"
                + lastRow.ToString(CultureInfo.InvariantCulture) + ":"
                + lastColumn.ToString(CultureInfo.InvariantCulture);
            if (!seen.Add(key)) continue;
            result.Add(new IWorkTableMergeRange(firstRow + 1, firstColumn + 1, lastRow + 1, lastColumn + 1));
        }
        if (IWorkMergeRangeValidator.HasOverlaps(result, columnCount)) {
            MarkMergeStorageUnsupported(model, diagnostics, ref supportsEditableReconstruction);
        }
        return Array.AsReadOnly(result.ToArray());
    }

    private static void MarkMergeStorageUnsupported(IWorkArchiveRecord model,
        List<IWorkDiagnostic> diagnostics, ref bool supportsEditableReconstruction) {
        supportsEditableReconstruction = false;
        if (diagnostics.Any(diagnostic => diagnostic.Code == "IWORK_TABLE_MERGE_UNSUPPORTED"
                && diagnostic.RecordIdentifier == model.Identifier)) return;
        diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
            "IWORK_TABLE_MERGE_UNSUPPORTED",
            "An iWork table contains a malformed or unsupported merged range; editable reconstruction is incomplete.",
            model.EntryPath, model.Identifier));
    }

    private static int CheckedSubDimension(ulong? value, int maximum, string label, IWorkArchiveRecord record) {
        ulong resolved = value ?? 0;
        if (resolved > (ulong)maximum || resolved > int.MaxValue) {
            throw new InvalidDataException($"iWork table {label} count {resolved} in object {record.Identifier} exceeds the table dimensions.");
        }
        return (int)resolved;
    }

    private static double? ValidDimension(double? value) => value.HasValue && IsFinite(value.Value) && value.Value > 0
        ? value
        : null;

    private static bool HasInvalidDeclaredDimension(IWorkWireMessage message, int field,
        double? value) => message.FieldCount(field) > 1
        || message.HasField(field)
            && (message.HasUnexpectedWireKind(field, IWorkWireKind.Fixed64)
                || !value.HasValue || !IsFinite(value.Value) || value.Value <= 0);

    private static int CheckedDimension(ulong? value, int maximum, string label, IWorkArchiveRecord record) {
        ulong resolved = value ?? 0;
        if (resolved > (ulong)maximum || resolved > int.MaxValue) {
            throw new InvalidDataException($"iWork table {label} count {resolved} in object {record.Identifier} exceeds the configured limit of {maximum}.");
        }
        return (int)resolved;
    }

    private static IWorkTableCell DecodeCell(byte[] buffer, int offset, int endOffset,
        int row, int column,
        IReadOnlyDictionary<uint, string> strings, IReadOnlyDictionary<uint, IWorkWireMessage> formulas,
        IWorkReadOptions options, IWorkProjectionBudget projectionBudget) {
        if (offset < 0 || endOffset < offset || endOffset > buffer.Length
            || offset > endOffset - 12) return Error(row, column, "Truncated cell record.");
        int version = buffer[offset];
        int type = buffer[offset + 1];
        if (version != 5) return Error(row, column, $"Unsupported cell storage version {version}.");
        uint flags = IWorkProtobuf.ReadUInt32(buffer, offset + 8);
        if ((flags & ~RecognizedCellValueMask) != 0) {
            return Error(row, column, "Cell storage contains unsupported value fields.");
        }
        int position = offset + 12;
        double? decimalValue = null;
        double doubleValue = 0;
        double dateValue = 0;
        uint stringIdentifier = 0;
        uint formulaIdentifier = 0;
        bool hasDecimal = false;
        bool hasDouble = false;
        bool hasDate = false;
        bool hasString = false;
        bool hasFormula = false;
        for (int bit = 0; bit < 21; bit++) {
            if ((flags & (1u << bit)) == 0) continue;
            int size = bit == 0 ? 16 : bit is 1 or 2 ? 8 : 4;
            if (position < 0 || position > endOffset - size) return Error(row, column, "Truncated cell value field.");
            switch (bit) {
                case 0:
                    decimalValue = ReadDecimal128(buffer, position);
                    hasDecimal = true;
                    break;
                case 1:
                    doubleValue = ReadDouble(buffer, position);
                    hasDouble = true;
                    break;
                case 2:
                    dateValue = ReadDouble(buffer, position);
                    hasDate = true;
                    break;
                case 3:
                    stringIdentifier = IWorkProtobuf.ReadUInt32(buffer, position);
                    hasString = true;
                    break;
                case 9:
                    formulaIdentifier = IWorkProtobuf.ReadUInt32(buffer, position);
                    hasFormula = true;
                    break;
            }
            position += size;
        }

        switch (type) {
            case 0:
                return new IWorkTableCell(row, column, IWorkCellKind.Empty, null);
            case 2:
            case 10:
                if (hasDecimal) {
                    return decimalValue.HasValue
                        ? FiniteNumber(row, column, decimalValue.Value, hasFormula,
                            formulaIdentifier, formulas, options, projectionBudget)
                        : Error(row, column,
                            "Decimal128 value exceeds XLSX numeric precision.");
                }
                if (hasDouble) return FiniteNumber(row, column, doubleValue, hasFormula, formulaIdentifier, formulas, options, projectionBudget);
                return hasFormula ? Formula(row, column, formulaIdentifier, formulas, options, projectionBudget) : Error(row, column, "Number cell has no value field.");
            case 3:
                if (hasString && strings.TryGetValue(stringIdentifier, out string? text)) {
                    return hasFormula
                        ? Formula(row, column, formulaIdentifier, formulas, options, projectionBudget, text, IWorkCellKind.Text)
                        : new IWorkTableCell(row, column, IWorkCellKind.Text, text);
                }
                return hasFormula ? Formula(row, column, formulaIdentifier, formulas, options, projectionBudget) : Error(row, column, $"Unresolved shared string {stringIdentifier}.");
            case 5:
                if (!hasDate) return Error(row, column, "Date cell has no date value field.");
                if (!IsFinite(dateValue)) return Error(row, column, "Date cell has a non-finite value.");
                if (!TryReadDateTime(dateValue, out DateTime value)) {
                    return Error(row, column, "Date cell is outside the supported DateTime range.");
                }
                return hasFormula
                    ? Formula(row, column, formulaIdentifier, formulas, options, projectionBudget, value, IWorkCellKind.DateTime)
                    : new IWorkTableCell(row, column, IWorkCellKind.DateTime, value);
            case 6:
                if (!hasDouble) return Error(row, column, "Boolean cell has no value field.");
                if (!IsFinite(doubleValue)) return Error(row, column, "Boolean cell has a non-finite value.");
                if (doubleValue is not 0d and not 1d) {
                    return Error(row, column, "Boolean cell value is not 0 or 1.");
                }
                bool booleanValue = doubleValue == 1d;
                return hasFormula
                    ? Formula(row, column, formulaIdentifier, formulas, options, projectionBudget, booleanValue, IWorkCellKind.Boolean)
                    : new IWorkTableCell(row, column, IWorkCellKind.Boolean, booleanValue);
            case 7:
                if (!hasDouble) return Error(row, column, "Duration cell has no value field.");
                if (!IsFinite(doubleValue)) return Error(row, column, "Duration cell has a non-finite value.");
                return hasFormula
                    ? Formula(row, column, formulaIdentifier, formulas, options, projectionBudget, doubleValue, IWorkCellKind.Duration)
                    : new IWorkTableCell(row, column, IWorkCellKind.Duration, doubleValue);
            case 8:
                return hasFormula
                    ? Formula(row, column, formulaIdentifier, formulas, options, projectionBudget, "#ERROR", IWorkCellKind.Error)
                    : Error(row, column, "#ERROR");
            case 9:
                return hasFormula ? Formula(row, column, formulaIdentifier, formulas, options, projectionBudget) : new IWorkTableCell(row, column, IWorkCellKind.Text, string.Empty);
            default:
                return Error(row, column, $"Unknown cell type {type}.");
        }
    }

    private static IWorkTableCell Formula(int row, int column, uint formulaIdentifier,
        IReadOnlyDictionary<uint, IWorkWireMessage> formulas, IWorkReadOptions options,
        IWorkProjectionBudget projectionBudget,
        object? cachedValue = null,
        IWorkCellKind? cachedValueKind = null) {
        IWorkFormulaResult result = formulas.TryGetValue(formulaIdentifier, out IWorkWireMessage? formula)
            ? IWorkFormulaReader.Render(formula, row - 1, column - 1,
                options.MaximumFormulaNodes, options.MaximumFormulaCharacters)
            : new IWorkFormulaResult("=?", false);
        string formulaText = result.Text.Length == 0 ? "=?" : result.Text;
        projectionBudget.AddTextCharacters(formulaText.Length);
        projectionBudget.AddTextItem();
        return new IWorkTableCell(row, column, IWorkCellKind.Formula, cachedValue,
            formula: formulaText, valueKind: cachedValueKind,
            formulaIsComplete: result.IsComplete);
    }

    private static IWorkTableCell FiniteNumber(int row, int column, double value, bool hasFormula,
        uint formulaIdentifier, IReadOnlyDictionary<uint, IWorkWireMessage> formulas,
        IWorkReadOptions options, IWorkProjectionBudget projectionBudget) =>
        IsFinite(value)
            ? hasFormula
                ? Formula(row, column, formulaIdentifier, formulas, options, projectionBudget, value, IWorkCellKind.Number)
                : new IWorkTableCell(row, column, IWorkCellKind.Number, value)
            : Error(row, column, "Number cell has a non-finite value.");

    private static IWorkTableCell Error(int row, int column, string message) =>
        new(row, column, IWorkCellKind.Error, null, error: message);

    private static void MarkCellStorageUnsupported(IWorkArchiveRecord tile,
        ICollection<IWorkDiagnostic> diagnostics, ref bool supportsEditableReconstruction) {
        supportsEditableReconstruction = false;
        if (!diagnostics.Any(diagnostic => diagnostic.Code == "IWORK_TABLE_CELL_STORAGE_UNSUPPORTED")) {
            diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                "IWORK_TABLE_CELL_STORAGE_UNSUPPORTED",
                "An iWork table row declares malformed or incomplete modern cell storage; editable reconstruction is incomplete.",
                tile.EntryPath, tile.Identifier));
        }
    }

    private static double ReadDouble(byte[] buffer, int offset) =>
        BitConverter.Int64BitsToDouble(unchecked((long)IWorkProtobuf.ReadUInt64(buffer, offset)));

    private static bool IsFinite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);

    private static bool TryReadDateTime(double seconds, out DateTime value) {
        long epochTicks = new DateTime(2001, 1, 1, 0, 0, 0, DateTimeKind.Utc).Ticks;
        double roundedDeltaTicks = Math.Round(seconds * TimeSpan.TicksPerSecond,
            MidpointRounding.AwayFromZero);
        if (!IsFinite(roundedDeltaTicks)
            || roundedDeltaTicks < -epochTicks
            || roundedDeltaTicks > DateTime.MaxValue.Ticks - epochTicks) {
            value = default;
            return false;
        }
        long absoluteTicks = epochTicks + (long)roundedDeltaTicks;
        if (absoluteTicks < DateTime.MinValue.Ticks
            || absoluteTicks > DateTime.MaxValue.Ticks) {
            value = default;
            return false;
        }
        value = new DateTime(absoluteTicks, DateTimeKind.Utc);
        return true;
    }

    private static double? ReadDecimal128(byte[] buffer, int offset) {
        int exponent = (((buffer[offset + 15] & 0x7f) << 7) | (buffer[offset + 14] >> 1)) - 0x1820;
        BigInteger coefficient = BigInteger.Zero;
        for (int index = 13; index >= 0; index--) {
            coefficient = coefficient * 256 + buffer[offset + index];
        }
        if ((buffer[offset + 14] & 1) != 0) coefficient += BigInteger.One << 112;
        if (coefficient.IsZero) return 0d;
        while (coefficient % 10 == 0) {
            coefficient /= 10;
            exponent++;
        }
        if (coefficient.ToString(CultureInfo.InvariantCulture).Length > 15) return null;
        string text = coefficient.ToString(CultureInfo.InvariantCulture)
            + "E" + exponent.ToString(CultureInfo.InvariantCulture);
        if (!double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture,
                out double value) || !IsFinite(value) || value == 0d) return null;
        return (buffer[offset + 15] & 0x80) != 0 ? -value : value;
    }
}
