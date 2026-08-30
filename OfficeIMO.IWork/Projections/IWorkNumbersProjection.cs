using OfficeIMO.IWork.Internal;
using System.Globalization;

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

    internal static IWorkNumbersProjection Read(IWorkSourceDocument source) {
        var diagnostics = new List<IWorkDiagnostic>();
        var sheets = new List<IWorkNumbersSheet>();
        IWorkObjectIndex index = source.Index;
        IWorkArchiveRecord? document = index.FirstOfType(DocumentArchive);
        if (document == null) {
            diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning, "IWORK_NUMBERS_DOCUMENT_MISSING",
                "No supported Numbers document root was found; editable reconstruction is unavailable."));
            return new IWorkNumbersProjection(source, sheets, diagnostics, supportsEditableReconstruction: false);
        }
        int materializedCellCount = 0;
        var projectionBudget = new IWorkProjectionBudget(source.Options);
        var projectedDrawableIdentifiers = new HashSet<ulong>();
        bool supportsEditableReconstruction = true;
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
                    IWorkTable? table = IWorkTableReader.Read(source, drawable, diagnostics,
                        ref materializedCellCount, ref supportsEditableReconstruction);
                    if (table != null) tables.Add(table);
                } else if (drawable.MessageType == TextShapeArchive) {
                    IWorkArchiveRecord? storage = index.Dereference(index.Message(drawable), 2);
                    if (storage != null && storage.MessageType == TextStorageArchive) {
                        string text = IWorkPagesReader.StorageText(index.Message(storage), projectionBudget,
                            out bool textComplete).Trim();
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
                }
            }
            string? sheetName = sheetMessage.GetString(1, out bool sheetNameComplete);
            if (!sheetNameComplete) {
                MarkTextMetadataUnsupported(sheetRecord, diagnostics, ref supportsEditableReconstruction);
            }
            sheets.Add(new IWorkNumbersSheet(sheetName ?? string.Empty, tables, textBoxes));
        }
        return new IWorkNumbersProjection(source, sheets, diagnostics, supportsEditableReconstruction);
    }

    internal static IWorkTable? ReadTableInfo(IWorkSourceDocument source, IWorkArchiveRecord tableRecord,
        List<IWorkDiagnostic> diagnostics, ref int materializedCellCount,
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
        IWorkArchiveRecord? model = source.Index.Dereference(tableInfo, 2);
        if (model == null || model.MessageType != TableModelArchive) {
            supportsEditableReconstruction = false;
            diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                "IWORK_TABLE_MODEL_UNSUPPORTED",
                "An iWork table does not reference a supported table model; editable reconstruction is incomplete.",
                tableRecord.EntryPath, tableRecord.Identifier));
            return null;
        }
        IWorkWireMessage? drawable = IWorkObjectIndex.TryGetMessage(tableInfo, 1, out bool malformedDrawable);
        if (malformedDrawable || tableInfo.HasBytes(1) && drawable == null) {
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
        return ReadTable(source, source.Index, model, geometry, diagnostics,
            ref materializedCellCount, ref supportsEditableReconstruction);
    }

    private static IWorkTable ReadTable(IWorkSourceDocument source, IWorkObjectIndex index,
        IWorkArchiveRecord model, IWorkGeometry? geometry, List<IWorkDiagnostic> diagnostics,
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
        string name = tableName ?? string.Empty;
        int headerRows = CheckedSubDimension(message.GetUnsigned(9), rows, "header row", model);
        int headerColumns = CheckedSubDimension(message.GetUnsigned(10), columns, "header column", model);
        int footerRows = CheckedSubDimension(message.GetUnsigned(11), rows, "footer row", model);
        double? defaultRowHeight = ValidDimension(message.GetDouble(16));
        double? defaultColumnWidth = ValidDimension(message.GetDouble(17));
        IReadOnlyList<IWorkTableMergeRange> mergedRanges = ReadMergedRanges(message, rows, columns,
            source.Options.MaximumTableMergedRanges, model, diagnostics, ref supportsEditableReconstruction);
        var cells = new List<IWorkTableCell>();
        var coordinates = new HashSet<long>();
        IWorkWireMessage? store = IWorkObjectIndex.TryGetMessage(message, 4);
        if (store == null) {
            MarkTableStorageUnsupported(model, diagnostics, ref supportsEditableReconstruction);
            return CreateTable();
        }

        IReadOnlyDictionary<uint, string> strings = ReadStrings(index, store, out bool stringStorageComplete);
        IReadOnlyDictionary<uint, IWorkWireMessage> formulas = ReadFormulas(index, store,
            out bool formulaStorageComplete);
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
        IWorkWireMessage? tileStorage = IWorkObjectIndex.TryGetMessage(store, 3);
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
                if ((currentBuffer == null) != (currentOffsets == null)) {
                    supportsEditableReconstruction = false;
                    if (!diagnostics.Any(diagnostic => diagnostic.Code == "IWORK_TABLE_CELL_STORAGE_UNSUPPORTED")) {
                        diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                            "IWORK_TABLE_CELL_STORAGE_UNSUPPORTED",
                            "An iWork table row declares incomplete modern cell storage; editable reconstruction is incomplete.",
                            tile.EntryPath, tile.Identifier));
                    }
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
                int availableColumns = Math.Min(columns, offsets.Length / 2);
                for (int column = 0; column < availableColumns; column++) {
                    int encodedOffset = offsets[column * 2] | offsets[column * 2 + 1] << 8;
                    if (encodedOffset == ushort.MaxValue) continue;
                    int offset = hasWideOffsets ? checked(encodedOffset * 4) : encodedOffset;
                    IWorkTableCell cell = DecodeCell(buffer, offset, checked((int)zeroBasedRow + 1), column + 1,
                        strings, formulas, source.Options);
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
            mergedRanges, geometry);
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

    private static IReadOnlyDictionary<uint, string> ReadStrings(IWorkObjectIndex index, IWorkWireMessage store,
        out bool fullyReconstructed) {
        var strings = new Dictionary<uint, string>();
        fullyReconstructed = true;
        IWorkArchiveRecord? list = index.Dereference(store, 4);
        if (list == null) return strings;
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
            uint normalizedKey = (uint)key.Value;
            if (strings.ContainsKey(normalizedKey)) fullyReconstructed = false;
            else strings.Add(normalizedKey, value);
        }
        return strings;
    }

    private static IReadOnlyDictionary<uint, IWorkWireMessage> ReadFormulas(IWorkObjectIndex index,
        IWorkWireMessage store, out bool fullyReconstructed) {
        var formulas = new Dictionary<uint, IWorkWireMessage>();
        var ambiguousIdentifiers = new HashSet<uint>();
        fullyReconstructed = true;
        IWorkArchiveRecord? list = index.Dereference(store, 6);
        if (list == null) return formulas;
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

    private static bool HasUnsupportedTableScalarEncoding(IWorkWireMessage message) =>
        message.LacksWireKind(6, IWorkWireKind.Varint)
        || message.LacksWireKind(7, IWorkWireKind.Varint)
        || message.LacksWireKind(9, IWorkWireKind.Varint)
        || message.LacksWireKind(10, IWorkWireKind.Varint)
        || message.LacksWireKind(11, IWorkWireKind.Varint)
        || message.LacksWireKind(16, IWorkWireKind.Fixed64)
        || message.LacksWireKind(17, IWorkWireKind.Fixed64);

    private static IReadOnlyList<IWorkTableMergeRange> ReadMergedRanges(IWorkWireMessage table,
        int rowCount, int columnCount, int maximumRanges, IWorkArchiveRecord model,
        List<IWorkDiagnostic> diagnostics,
        ref bool supportsEditableReconstruction) {
        IWorkWireMessage? mergeOwner = IWorkObjectIndex.TryGetMessage(table, 47, out bool malformedOwner);
        if (malformedOwner) {
            MarkMergeStorageUnsupported(model, diagnostics, ref supportsEditableReconstruction);
            return Array.Empty<IWorkTableMergeRange>();
        }
        if (mergeOwner == null) return Array.Empty<IWorkTableMergeRange>();
        IWorkWireMessage? formulaStore = IWorkObjectIndex.TryGetMessage(mergeOwner, 2, out bool malformedStore);
        if (malformedStore || formulaStore == null) {
            MarkMergeStorageUnsupported(model, diagnostics, ref supportsEditableReconstruction);
            return Array.Empty<IWorkTableMergeRange>();
        }
        IReadOnlyList<IWorkWireMessage> pairs = IWorkObjectIndex.TryGetMessages(formulaStore, 3, out bool malformedPairs);
        if (malformedPairs) {
            MarkMergeStorageUnsupported(model, diagnostics, ref supportsEditableReconstruction);
            return Array.Empty<IWorkTableMergeRange>();
        }
        if (pairs.Count > maximumRanges) {
            throw new InvalidDataException($"iWork table merged-range count {pairs.Count} in object {model.Identifier} exceeds the configured limit of {maximumRanges}.");
        }
        var result = new List<IWorkTableMergeRange>();
        var seen = new HashSet<string>(StringComparer.Ordinal);
        foreach (IWorkWireMessage pair in pairs) {
            IWorkWireMessage? formula = IWorkObjectIndex.TryGetMessage(pair, 2, out bool malformedFormula);
            if (malformedFormula || formula == null
                || !IWorkFormulaReader.TryReadAbsoluteRange(formula,
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

    private static int CheckedDimension(ulong? value, int maximum, string label, IWorkArchiveRecord record) {
        ulong resolved = value ?? 0;
        if (resolved > (ulong)maximum || resolved > int.MaxValue) {
            throw new InvalidDataException($"iWork table {label} count {resolved} in object {record.Identifier} exceeds the configured limit of {maximum}.");
        }
        return (int)resolved;
    }

    private static IWorkTableCell DecodeCell(byte[] buffer, int offset, int row, int column,
        IReadOnlyDictionary<uint, string> strings, IReadOnlyDictionary<uint, IWorkWireMessage> formulas,
        IWorkReadOptions options) {
        if (offset < 0 || offset > buffer.Length - 12) return Error(row, column, "Truncated cell record.");
        int version = buffer[offset];
        int type = buffer[offset + 1];
        if (version != 5) return Error(row, column, $"Unsupported cell storage version {version}.");
        uint flags = IWorkProtobuf.ReadUInt32(buffer, offset + 8);
        int position = offset + 12;
        double decimalValue = 0;
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
            if (position < 0 || position > buffer.Length - size) return Error(row, column, "Truncated cell value field.");
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
                if (hasDecimal) return FiniteNumber(row, column, decimalValue, hasFormula, formulaIdentifier, formulas, options);
                if (hasDouble) return FiniteNumber(row, column, doubleValue, hasFormula, formulaIdentifier, formulas, options);
                return hasFormula ? Formula(row, column, formulaIdentifier, formulas, options) : Error(row, column, "Number cell has no value field.");
            case 3:
                if (hasString && strings.TryGetValue(stringIdentifier, out string? text)) {
                    return hasFormula
                        ? Formula(row, column, formulaIdentifier, formulas, options, text, IWorkCellKind.Text)
                        : new IWorkTableCell(row, column, IWorkCellKind.Text, text);
                }
                return hasFormula ? Formula(row, column, formulaIdentifier, formulas, options) : Error(row, column, $"Unresolved shared string {stringIdentifier}.");
            case 5:
                if (!hasDate) return Error(row, column, "Date cell has no date value field.");
                if (!IsFinite(dateValue)) return Error(row, column, "Date cell has a non-finite value.");
                try {
                    DateTime value = new DateTime(2001, 1, 1, 0, 0, 0, DateTimeKind.Utc).AddSeconds(dateValue);
                    return hasFormula
                        ? Formula(row, column, formulaIdentifier, formulas, options, value, IWorkCellKind.DateTime)
                        : new IWorkTableCell(row, column, IWorkCellKind.DateTime, value);
                } catch (ArgumentOutOfRangeException) {
                    return Error(row, column, "Date cell is outside the supported DateTime range.");
                }
            case 6:
                if (!hasDouble) return Error(row, column, "Boolean cell has no value field.");
                if (!IsFinite(doubleValue)) return Error(row, column, "Boolean cell has a non-finite value.");
                bool booleanValue = doubleValue != 0;
                return hasFormula
                    ? Formula(row, column, formulaIdentifier, formulas, options, booleanValue, IWorkCellKind.Boolean)
                    : new IWorkTableCell(row, column, IWorkCellKind.Boolean, booleanValue);
            case 7:
                if (!hasDouble) return Error(row, column, "Duration cell has no value field.");
                if (!IsFinite(doubleValue)) return Error(row, column, "Duration cell has a non-finite value.");
                return hasFormula
                    ? Formula(row, column, formulaIdentifier, formulas, options, doubleValue, IWorkCellKind.Duration)
                    : new IWorkTableCell(row, column, IWorkCellKind.Duration, doubleValue);
            case 8:
                return hasFormula
                    ? Formula(row, column, formulaIdentifier, formulas, options, "#ERROR", IWorkCellKind.Error)
                    : Error(row, column, "#ERROR");
            case 9:
                return hasFormula ? Formula(row, column, formulaIdentifier, formulas, options) : new IWorkTableCell(row, column, IWorkCellKind.Text, string.Empty);
            default:
                return Error(row, column, $"Unknown cell type {type}.");
        }
    }

    private static IWorkTableCell Formula(int row, int column, uint formulaIdentifier,
        IReadOnlyDictionary<uint, IWorkWireMessage> formulas, IWorkReadOptions options,
        object? cachedValue = null,
        IWorkCellKind? cachedValueKind = null) {
        IWorkFormulaResult result = formulas.TryGetValue(formulaIdentifier, out IWorkWireMessage? formula)
            ? IWorkFormulaReader.Render(formula, row - 1, column - 1,
                options.MaximumFormulaNodes, options.MaximumFormulaCharacters)
            : new IWorkFormulaResult("=?", false);
        return new IWorkTableCell(row, column, IWorkCellKind.Formula, cachedValue,
            formula: result.Text.Length == 0 ? "=?" : result.Text, valueKind: cachedValueKind,
            formulaIsComplete: result.IsComplete);
    }

    private static IWorkTableCell FiniteNumber(int row, int column, double value, bool hasFormula,
        uint formulaIdentifier, IReadOnlyDictionary<uint, IWorkWireMessage> formulas,
        IWorkReadOptions options) =>
        IsFinite(value)
            ? hasFormula
                ? Formula(row, column, formulaIdentifier, formulas, options, value, IWorkCellKind.Number)
                : new IWorkTableCell(row, column, IWorkCellKind.Number, value)
            : Error(row, column, "Number cell has a non-finite value.");

    private static IWorkTableCell Error(int row, int column, string message) =>
        new(row, column, IWorkCellKind.Error, null, error: message);

    private static double ReadDouble(byte[] buffer, int offset) =>
        BitConverter.Int64BitsToDouble(unchecked((long)IWorkProtobuf.ReadUInt64(buffer, offset)));

    private static bool IsFinite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);

    private static double ReadDecimal128(byte[] buffer, int offset) {
        int exponent = (((buffer[offset + 15] & 0x7f) << 7) | (buffer[offset + 14] >> 1)) - 0x1820;
        double coefficient = 0;
        for (int index = 13; index >= 0; index--) coefficient = coefficient * 256 + buffer[offset + index];
        if ((buffer[offset + 14] & 1) != 0) coefficient += 5.192296858534828e33;
        double value = coefficient * Math.Pow(10, exponent);
        return (buffer[offset + 15] & 0x80) != 0 ? -value : value;
    }
}
