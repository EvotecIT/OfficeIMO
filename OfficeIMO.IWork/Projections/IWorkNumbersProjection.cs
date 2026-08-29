using OfficeIMO.IWork.Internal;
using System.Globalization;

namespace OfficeIMO.IWork;

/// <summary>One materialized, non-empty Numbers cell.</summary>
public sealed class IWorkNumbersCell {
    internal IWorkNumbersCell(int row, int column, IWorkCellKind kind, object? value,
        string? formula = null, string? error = null) {
        Row = row;
        Column = column;
        Kind = kind;
        Value = value;
        Formula = formula;
        Error = error;
    }

    /// <summary>Gets the one-based row position.</summary>
    public int Row { get; }
    /// <summary>Gets the one-based column position.</summary>
    public int Column { get; }
    /// <summary>Gets the recovered value kind.</summary>
    public IWorkCellKind Kind { get; }
    /// <summary>Gets the typed cached value, when one was recovered.</summary>
    public object? Value { get; }
    /// <summary>Gets a formula marker when the source formula cannot yet be reconstructed.</summary>
    public string? Formula { get; }
    /// <summary>Gets a cell-level decode error without failing the surrounding table.</summary>
    public string? Error { get; }
    /// <summary>Gets a culture-invariant display representation of the recovered value.</summary>
    public string DisplayText => Kind switch {
        IWorkCellKind.Boolean => Convert.ToBoolean(Value, CultureInfo.InvariantCulture) ? "TRUE" : "FALSE",
        IWorkCellKind.DateTime when Value is DateTime date => date.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture),
        IWorkCellKind.Duration when Value is double seconds => seconds.ToString("R", CultureInfo.InvariantCulture) + "s",
        IWorkCellKind.Formula => Formula ?? "=?",
        IWorkCellKind.Error => Error ?? "#ERROR",
        _ => Convert.ToString(Value, CultureInfo.InvariantCulture) ?? string.Empty
    };
}

/// <summary>A sparse Numbers table projection with declared dimensions.</summary>
public sealed class IWorkNumbersTable {
    private readonly Dictionary<long, IWorkNumbersCell> _cells;

    internal IWorkNumbersTable(string name, int rowCount, int columnCount,
        IReadOnlyList<IWorkNumbersCell> cells) {
        Name = name;
        RowCount = rowCount;
        ColumnCount = columnCount;
        _cells = new Dictionary<long, IWorkNumbersCell>();
        foreach (IWorkNumbersCell cell in cells) _cells[Key(cell.Row, cell.Column)] = cell;
        Cells = _cells.Values.OrderBy(cell => cell.Row).ThenBy(cell => cell.Column).ToArray();
    }

    /// <summary>Gets the source table name.</summary>
    public string Name { get; }
    /// <summary>Gets the declared row count without allocating an equivalent dense grid.</summary>
    public int RowCount { get; }
    /// <summary>Gets the declared column count without allocating an equivalent dense grid.</summary>
    public int ColumnCount { get; }
    /// <summary>Gets materialized non-empty or diagnostic cells.</summary>
    public IReadOnlyList<IWorkNumbersCell> Cells { get; }

    /// <summary>Returns a materialized cell at a one-based position, or null when the source cell is empty.</summary>
    public IWorkNumbersCell? GetCell(int row, int column) {
        if (row < 1 || row > RowCount) throw new ArgumentOutOfRangeException(nameof(row));
        if (column < 1 || column > ColumnCount) throw new ArgumentOutOfRangeException(nameof(column));
        return _cells.TryGetValue(Key(row, column), out IWorkNumbersCell? cell) ? cell : null;
    }

    private static long Key(int row, int column) => ((long)row << 32) | (uint)column;
}

/// <summary>One Numbers sheet and its semantic drawables.</summary>
public sealed class IWorkNumbersSheet {
    internal IWorkNumbersSheet(string name, IReadOnlyList<IWorkNumbersTable> tables,
        IReadOnlyList<string> textBoxes) {
        Name = name;
        Tables = tables;
        TextBoxes = textBoxes;
    }

    /// <summary>Gets the source sheet name.</summary>
    public string Name { get; }
    /// <summary>Gets tables in drawable order.</summary>
    public IReadOnlyList<IWorkNumbersTable> Tables { get; }
    /// <summary>Gets text-box content in drawable order.</summary>
    public IReadOnlyList<string> TextBoxes { get; }
}

/// <summary>Read-only Numbers structure recovered from a shared IWA object graph.</summary>
public sealed class IWorkNumbersProjection {
    private readonly IWorkSourceDocument _source;
    private readonly IReadOnlyCollection<ulong> _recognizedIdentifiers;
    private readonly bool _supportsEditableReconstruction;

    internal IWorkNumbersProjection(IWorkSourceDocument source, IReadOnlyList<IWorkNumbersSheet> sheets,
        IReadOnlyCollection<ulong> recognizedIdentifiers, IReadOnlyList<IWorkDiagnostic> diagnostics,
        bool supportsEditableReconstruction) {
        _source = source;
        Sheets = sheets;
        _recognizedIdentifiers = recognizedIdentifiers;
        Diagnostics = diagnostics;
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
        return _source.CreateReport(kind, _recognizedIdentifiers, Diagnostics, preview,
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
            return new IWorkNumbersProjection(this, Array.Empty<IWorkNumbersSheet>(), Array.Empty<ulong>(),
                new[] { IWorkProjectionDiagnostics.SemanticProjectionSkipped }, supportsEditableReconstruction: false);
        }
        return IWorkNumbersReader.Read(this);
    }
}

internal static class IWorkNumbersReader {
    private const uint DocumentArchive = 1;
    private const uint TableInfoArchive = 6000;
    private const uint TableModelArchive = 6001;
    private const uint TableTileArchive = 6002;
    private const uint TextStorageArchive = 2001;
    private const uint TextShapeArchive = 2011;
    private const int TileRowStride = 256;

    internal static IWorkNumbersProjection Read(IWorkSourceDocument source) {
        var recognized = new HashSet<ulong>();
        var diagnostics = new List<IWorkDiagnostic>();
        var sheets = new List<IWorkNumbersSheet>();
        IWorkObjectIndex index = source.Index;
        IWorkArchiveRecord? document = index.FirstOfType(DocumentArchive);
        if (document == null) {
            diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning, "IWORK_NUMBERS_DOCUMENT_MISSING",
                "No supported Numbers document root was found; editable reconstruction is unavailable."));
            return new IWorkNumbersProjection(source, sheets, recognized, diagnostics, supportsEditableReconstruction: false);
        }
        recognized.Add(document.Identifier);
        int materializedCellCount = 0;
        bool supportsEditableReconstruction = true;
        IReadOnlyList<IWorkArchiveRecord> sheetRecords = index.DereferenceAll(
            index.Message(document), 1, out int unresolvedSheetCount);
        if (unresolvedSheetCount > 0) {
            supportsEditableReconstruction = false;
            recognized.Remove(document.Identifier);
            diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                "IWORK_NUMBERS_SHEET_UNSUPPORTED",
                "The Numbers document references a missing sheet; editable reconstruction is incomplete.",
                document.EntryPath, document.Identifier));
        }

        foreach (IWorkArchiveRecord sheetRecord in sheetRecords) {
            recognized.Add(sheetRecord.Identifier);
            IWorkWireMessage sheetMessage = index.Message(sheetRecord);
            var tables = new List<IWorkNumbersTable>();
            var textBoxes = new List<string>();
            IReadOnlyList<IWorkArchiveRecord> drawables = index.DereferenceAll(
                sheetMessage, 2, out int unresolvedDrawableCount);
            if (unresolvedDrawableCount > 0) {
                supportsEditableReconstruction = false;
                recognized.Remove(sheetRecord.Identifier);
                diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                    "IWORK_NUMBERS_DRAWABLE_UNSUPPORTED",
                    "A Numbers sheet references a missing drawable; editable reconstruction is incomplete.",
                    sheetRecord.EntryPath, sheetRecord.Identifier));
            }
            foreach (IWorkArchiveRecord drawable in drawables) {
                if (drawable.MessageType == TableInfoArchive) {
                    IWorkArchiveRecord? model = index.Dereference(index.Message(drawable), 2);
                    if (model != null && model.MessageType == TableModelArchive) {
                        recognized.Add(drawable.Identifier);
                        recognized.Add(model.Identifier);
                        tables.Add(ReadTable(source, index, model, recognized, diagnostics,
                            ref materializedCellCount, ref supportsEditableReconstruction));
                    } else {
                        supportsEditableReconstruction = false;
                        diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                            "IWORK_NUMBERS_TABLE_MODEL_UNSUPPORTED",
                            "A Numbers table does not reference a supported table model; editable reconstruction is incomplete.",
                            drawable.EntryPath, drawable.Identifier));
                    }
                } else if (drawable.MessageType == TextShapeArchive) {
                    IWorkArchiveRecord? storage = index.Dereference(index.Message(drawable), 2);
                    if (storage != null && storage.MessageType == TextStorageArchive) {
                        string text = IWorkPagesReader.StorageText(index.Message(storage)).Trim();
                        if (text.Length > 0) textBoxes.Add(text);
                        recognized.Add(drawable.Identifier);
                        recognized.Add(storage.Identifier);
                    } else {
                        supportsEditableReconstruction = false;
                        diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                            "IWORK_NUMBERS_TEXT_STORAGE_UNSUPPORTED",
                            "A Numbers text shape does not reference supported text storage; editable reconstruction is incomplete.",
                            drawable.EntryPath, drawable.Identifier));
                    }
                }
            }
            sheets.Add(new IWorkNumbersSheet(sheetMessage.GetString(1) ?? string.Empty, tables, textBoxes));
        }
        return new IWorkNumbersProjection(source, sheets, recognized, diagnostics, supportsEditableReconstruction);
    }

    private static IWorkNumbersTable ReadTable(IWorkSourceDocument source, IWorkObjectIndex index,
        IWorkArchiveRecord model, HashSet<ulong> recognized, List<IWorkDiagnostic> diagnostics,
        ref int materializedCellCount, ref bool supportsEditableReconstruction) {
        IWorkWireMessage message = index.Message(model);
        int rows = CheckedDimension(message.GetUnsigned(6), source.Options.MaximumTableRows, "row", model);
        int columns = CheckedDimension(message.GetUnsigned(7), source.Options.MaximumTableColumns, "column", model);
        string name = message.GetString(8) ?? string.Empty;
        var cells = new List<IWorkNumbersCell>();
        IWorkWireMessage? store = IWorkObjectIndex.TryGetMessage(message, 4);
        if (store == null) {
            MarkTableStorageUnsupported(model, recognized, diagnostics, ref supportsEditableReconstruction);
            return new IWorkNumbersTable(name, rows, columns, cells);
        }

        IReadOnlyDictionary<uint, string> strings = ReadStrings(index, store, recognized);
        IWorkWireMessage? tileStorage = IWorkObjectIndex.TryGetMessage(store, 3);
        if (tileStorage == null) {
            MarkTableStorageUnsupported(model, recognized, diagnostics, ref supportsEditableReconstruction);
            return new IWorkNumbersTable(name, rows, columns, cells);
        }
        foreach (IWorkWireMessage tileEntry in IWorkObjectIndex.TryGetMessages(tileStorage, 1)) {
            ulong rawTileId = tileEntry.GetUnsigned(1) ?? 0;
            if (rawTileId > int.MaxValue) {
                supportsEditableReconstruction = false;
                recognized.Remove(model.Identifier);
                diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                    "IWORK_NUMBERS_TILE_INDEX_UNSUPPORTED",
                    "A Numbers table tile index exceeds the supported range; editable reconstruction is incomplete.",
                    model.EntryPath, model.Identifier));
                continue;
            }
            IWorkArchiveRecord? tile = index.Dereference(tileEntry, 2);
            if (tile == null || tile.MessageType != TableTileArchive) {
                supportsEditableReconstruction = false;
                recognized.Remove(model.Identifier);
                diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                    "IWORK_NUMBERS_TILE_UNSUPPORTED",
                    "A Numbers table references a missing or unsupported tile object; editable reconstruction is incomplete.",
                    model.EntryPath, model.Identifier));
                continue;
            }
            bool tileFullyReconstructed = true;
            foreach (IWorkWireMessage rowInfo in IWorkObjectIndex.TryGetMessages(index.Message(tile), 5)) {
                byte[]? currentBuffer = rowInfo.GetBytes(6);
                byte[]? currentOffsets = rowInfo.GetBytes(7);
                bool hasPreBncStorage = (rowInfo.GetBytes(3)?.Length ?? 0) > 0
                    || (rowInfo.GetBytes(4)?.Length ?? 0) > 0;
                if ((currentBuffer == null || currentOffsets == null) && hasPreBncStorage) {
                    supportsEditableReconstruction = false;
                    tileFullyReconstructed = false;
                    if (!diagnostics.Any(diagnostic => diagnostic.Code == "IWORK_NUMBERS_LEGACY_CELL_STORAGE")) {
                        diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                            "IWORK_NUMBERS_LEGACY_CELL_STORAGE",
                            "The source uses pre-BNC Numbers cell storage. Records are preserved, but editable reconstruction is unavailable.",
                            tile.EntryPath, tile.Identifier));
                    }
                    continue;
                }
                ulong rawRow = rowInfo.GetUnsigned(1) ?? 0;
                long zeroBasedRow = checked((long)rawTileId * TileRowStride + (long)rawRow);
                if (zeroBasedRow < 0 || zeroBasedRow >= rows) {
                    supportsEditableReconstruction = false;
                    tileFullyReconstructed = false;
                    if (!diagnostics.Any(diagnostic => diagnostic.Code == "IWORK_NUMBERS_TILE_ROW_UNSUPPORTED")) {
                        diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                            "IWORK_NUMBERS_TILE_ROW_UNSUPPORTED",
                            "A Numbers tile contains a row outside the declared table bounds; editable reconstruction is incomplete.",
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
                    IWorkNumbersCell cell = DecodeCell(buffer, offset, checked((int)zeroBasedRow + 1), column + 1, strings);
                    if (cell.Kind == IWorkCellKind.Empty) continue;
                    if (materializedCellCount >= source.Options.MaximumMaterializedCells) {
                        throw new InvalidDataException($"Numbers cell count exceeds the configured source-wide limit of {source.Options.MaximumMaterializedCells}.");
                    }
                    cells.Add(cell);
                    materializedCellCount++;
                }
            }
            if (tileFullyReconstructed) recognized.Add(tile.Identifier);
        }

        int errorCount = cells.Count(cell => cell.Kind == IWorkCellKind.Error);
        if (errorCount > 0) {
            diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning, "IWORK_NUMBERS_CELL_DECODE",
                $"{errorCount} cells in table '{name}' could not be decoded completely.", model.EntryPath, model.Identifier));
        }
        return new IWorkNumbersTable(name, rows, columns, cells);
    }

    private static void MarkTableStorageUnsupported(IWorkArchiveRecord model, HashSet<ulong> recognized,
        List<IWorkDiagnostic> diagnostics, ref bool supportsEditableReconstruction) {
        supportsEditableReconstruction = false;
        recognized.Remove(model.Identifier);
        diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
            "IWORK_NUMBERS_TABLE_STORAGE_UNSUPPORTED",
            "A Numbers table has no supported tile storage; editable reconstruction is incomplete.",
            model.EntryPath, model.Identifier));
    }

    private static IReadOnlyDictionary<uint, string> ReadStrings(IWorkObjectIndex index, IWorkWireMessage store,
        HashSet<ulong> recognized) {
        var strings = new Dictionary<uint, string>();
        IWorkArchiveRecord? list = index.Dereference(store, 4);
        if (list == null) return strings;
        recognized.Add(list.Identifier);
        foreach (IWorkWireMessage entry in IWorkObjectIndex.TryGetMessages(index.Message(list), 3)) {
            ulong? key = entry.GetUnsigned(1);
            string? value = entry.GetString(3);
            if (key.HasValue && key.Value <= uint.MaxValue && value != null) strings[(uint)key.Value] = value;
        }
        return strings;
    }

    private static int CheckedDimension(ulong? value, int maximum, string label, IWorkArchiveRecord record) {
        ulong resolved = value ?? 0;
        if (resolved > (ulong)maximum || resolved > int.MaxValue) {
            throw new InvalidDataException($"Numbers table {label} count {resolved} in object {record.Identifier} exceeds the configured limit of {maximum}.");
        }
        return (int)resolved;
    }

    private static IWorkNumbersCell DecodeCell(byte[] buffer, int offset, int row, int column,
        IReadOnlyDictionary<uint, string> strings) {
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
                    hasFormula = true;
                    break;
            }
            position += size;
        }

        switch (type) {
            case 0:
                return new IWorkNumbersCell(row, column, IWorkCellKind.Empty, null);
            case 2:
            case 10:
                if (hasDecimal) return FiniteNumber(row, column, decimalValue, hasFormula);
                if (hasDouble) return FiniteNumber(row, column, doubleValue, hasFormula);
                return hasFormula ? Formula(row, column) : Error(row, column, "Number cell has no value field.");
            case 3:
                if (hasString && strings.TryGetValue(stringIdentifier, out string? text)) {
                    return hasFormula
                        ? Formula(row, column, text)
                        : new IWorkNumbersCell(row, column, IWorkCellKind.Text, text);
                }
                return hasFormula ? Formula(row, column) : Error(row, column, $"Unresolved shared string {stringIdentifier}.");
            case 5:
                if (!hasDate) return Error(row, column, "Date cell has no date value field.");
                if (!IsFinite(dateValue)) return Error(row, column, "Date cell has a non-finite value.");
                try {
                    DateTime value = new DateTime(2001, 1, 1, 0, 0, 0, DateTimeKind.Utc).AddSeconds(dateValue);
                    return hasFormula
                        ? Formula(row, column, value)
                        : new IWorkNumbersCell(row, column, IWorkCellKind.DateTime, value);
                } catch (ArgumentOutOfRangeException) {
                    return Error(row, column, "Date cell is outside the supported DateTime range.");
                }
            case 6:
                if (!hasDouble) return Error(row, column, "Boolean cell has no value field.");
                if (!IsFinite(doubleValue)) return Error(row, column, "Boolean cell has a non-finite value.");
                bool booleanValue = doubleValue != 0;
                return hasFormula
                    ? Formula(row, column, booleanValue)
                    : new IWorkNumbersCell(row, column, IWorkCellKind.Boolean, booleanValue);
            case 7:
                if (!hasDouble) return Error(row, column, "Duration cell has no value field.");
                if (!IsFinite(doubleValue)) return Error(row, column, "Duration cell has a non-finite value.");
                return hasFormula
                    ? Formula(row, column, doubleValue)
                    : new IWorkNumbersCell(row, column, IWorkCellKind.Duration, doubleValue);
            case 8:
                return Error(row, column, "#ERROR");
            case 9:
                return hasFormula ? Formula(row, column) : new IWorkNumbersCell(row, column, IWorkCellKind.Text, string.Empty);
            default:
                return Error(row, column, $"Unknown cell type {type}.");
        }
    }

    private static IWorkNumbersCell Formula(int row, int column, object? cachedValue = null) =>
        new(row, column, IWorkCellKind.Formula, cachedValue, formula: "=?");

    private static IWorkNumbersCell FiniteNumber(int row, int column, double value, bool hasFormula) =>
        IsFinite(value)
            ? hasFormula
                ? Formula(row, column, value)
                : new IWorkNumbersCell(row, column, IWorkCellKind.Number, value)
            : Error(row, column, "Number cell has a non-finite value.");

    private static IWorkNumbersCell Error(int row, int column, string message) =>
        new(row, column, IWorkCellKind.Error, null, error: message);

    private static double ReadDouble(byte[] buffer, int offset) =>
        BitConverter.Int64BitsToDouble(unchecked((long)IWorkProtobuf.ReadUInt64(buffer, offset)));

    private static bool IsFinite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);

    private static double ReadDecimal128(byte[] buffer, int offset) {
        int exponent = (((buffer[offset + 15] & 0x7f) << 7) | (buffer[offset + 14] >> 1)) - 0x1820;
        double coefficient = 0;
        for (int index = 13; index >= 0; index--) coefficient = coefficient * 256 + buffer[offset + index];
        if ((buffer[offset + 14] & 1) != 0) coefficient += 1.208925819614629e24;
        double value = coefficient * Math.Pow(10, exponent);
        return (buffer[offset + 15] & 0x80) != 0 ? -value : value;
    }
}
