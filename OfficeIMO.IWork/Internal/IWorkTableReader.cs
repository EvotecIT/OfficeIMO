namespace OfficeIMO.IWork.Internal;

/// <summary>Shared bounded table reader for Pages, Numbers, and Keynote.</summary>
internal static partial class IWorkTableReader {
    private const uint TableInfoArchive = 6000;
    private const uint WordProcessingTableInfoArchive = 6007;
    private const uint TableModelArchive = 6001;
    private const uint TableTileArchive = 6002;
    private const int TileRowStride = 256;
    private const int MaximumTileMetadataFields = 7;
    private const uint RecognizedCellValueMask = (1u << 21) - 1;

    internal static IWorkTable? Read(IWorkSourceDocument source, IWorkArchiveRecord tableRecord,
        IWorkProjectionBudget projectionBudget, IWorkSourceReferenceIssueCollector references,
        List<IWorkDiagnostic> diagnostics,
        ref int materializedCellCount,
        ref bool supportsEditableReconstruction) {
        IWorkWireMessage recordMessage;
        try {
            recordMessage = source.Index.Message(tableRecord);
        } catch (InvalidDataException exception) when (!IWorkProtobuf.IsLimitException(exception)) {
            references.Declarations.Record(tableRecord, "$", null);
            supportsEditableReconstruction = false;
            diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                "IWORK_TABLE_INFO_UNSUPPORTED",
                "An iWork table-info record is malformed; editable reconstruction is incomplete.",
                tableRecord.EntryPath, tableRecord.Identifier));
            return null;
        }
        IWorkWireMessage? tableInfo = tableRecord.MessageType switch {
            TableInfoArchive => recordMessage,
            WordProcessingTableInfoArchive => IWorkObjectIndex.TryGetMessage(recordMessage, 1),
            _ => null
        };
        if (tableInfo == null) {
            if (tableRecord.MessageType == WordProcessingTableInfoArchive && recordMessage.HasField(1))
                references.Declarations.Record(tableRecord, "1", recordMessage.FieldCount(1));
            supportsEditableReconstruction = false;
            diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                "IWORK_TABLE_INFO_UNSUPPORTED",
                "An iWork table has no supported table-info payload; editable reconstruction is incomplete.",
                tableRecord.EntryPath, tableRecord.Identifier));
            return null;
        }
        bool modelReferenceComplete = tableInfo.FieldCount(2) == 1
            && !tableInfo.HasUnexpectedWireKind(2, IWorkWireKind.Bytes);
        IWorkArchiveRecord? model = references.ReadOne(tableRecord, tableInfo, 2,
            tableRecord.MessageType == WordProcessingTableInfoArchive ? "1/2" : "2");
        if (!modelReferenceComplete || model == null || model.MessageType != TableModelArchive) {
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
        bool requirePositiveGeometry = source.Kind == IWorkDocumentKind.Keynote;
        bool geometryComplete = !requirePositiveGeometry || drawable != null;
        IWorkGeometry? geometry = drawable == null
            ? null
            : IWorkDrawingReader.ReadGeometry(drawable, out geometryComplete,
                requirePositiveSize: requirePositiveGeometry);
        if (!geometryComplete) {
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
        IWorkWireMessage modelMessage;
        try {
            modelMessage = source.Index.Message(model);
        } catch (InvalidDataException exception) when (!IWorkProtobuf.IsLimitException(exception)) {
            references.Declarations.Record(model, "$", null);
            supportsEditableReconstruction = false;
            diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                "IWORK_TABLE_MODEL_UNSUPPORTED",
                "An iWork table references a malformed table model; editable reconstruction is incomplete.",
                model.EntryPath, model.Identifier));
            return null;
        }
        return ReadTable(source, source.Index, model, modelMessage, geometry, projectionBudget, references,
            diagnostics, accessibilityDescription, ref materializedCellCount,
            ref supportsEditableReconstruction, new IWorkObjectIdentity(tableRecord));
    }

    private static IWorkTable ReadTable(IWorkSourceDocument source, IWorkObjectIndex index,
        IWorkArchiveRecord model, IWorkWireMessage message, IWorkGeometry? geometry,
        IWorkProjectionBudget projectionBudget,
        IWorkSourceReferenceIssueCollector references,
        List<IWorkDiagnostic> diagnostics,
        string? accessibilityDescription,
        ref int materializedCellCount, ref bool supportsEditableReconstruction, IWorkObjectIdentity sourceIdentity) {
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
        IReadOnlyList<IWorkTableMergeRange> mergedRanges = ReadMergedRanges(source, message, rows, columns,
            source.Options.MaximumTableMergedRanges, source.Options.MaximumFormulaNodes,
            model, references, diagnostics, ref supportsEditableReconstruction);
        var rowHeights = new Dictionary<int, double>();
        var columnWidths = new Dictionary<int, double>();
        var cells = new List<IWorkTableCell>();
        var omittedTextUnits = new List<IWorkObjectIdentity>();
        var coordinates = new HashSet<long>();
        var formulaRichStringIdentifiers = new HashSet<uint>();
        var nonFormulaRichStringIdentifiers = new HashSet<uint>();
        IWorkWireMessage? store = IWorkObjectIndex.TryGetMessage(message, 4);
        if (store == null) {
            if (message.HasField(4)) references.Declarations.Record(model, "4", message.FieldCount(4));
            MarkTableStorageUnsupported(model, diagnostics, ref supportsEditableReconstruction);
            return CreateTable();
        }

        ReadHeaderDimensions(source, store, model, rows, columns, projectionBudget,
            references, rowHeights, columnWidths, diagnostics, ref supportsEditableReconstruction);

        IReadOnlyDictionary<uint, string> strings = ReadStrings(source, store, model, references,
            projectionBudget,
            out bool stringStorageComplete);
        IWorkTableRichTextCatalog richStrings = IWorkTableRichTextCatalog.Create(source, store, model,
            projectionBudget, references);
        IReadOnlyDictionary<uint, IWorkWireMessage> formulas = ReadFormulas(source, store, model, references,
            projectionBudget,
            out bool formulaStorageComplete, out bool formulaCatalogEnvelopeComplete);
        if (!stringStorageComplete) {
            supportsEditableReconstruction = false;
            diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                "IWORK_TABLE_STRING_STORAGE_UNSUPPORTED",
                "An iWork string catalog is unresolved or contains malformed or duplicate entries; editable reconstruction is incomplete.",
                model.EntryPath, model.Identifier));
        }
        if (!formulaStorageComplete) {
            if (!formulaCatalogEnvelopeComplete) supportsEditableReconstruction = false;
            diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                "IWORK_TABLE_FORMULA_STORAGE_UNSUPPORTED",
                "An iWork formula catalog is unresolved or contains malformed or duplicate entries; affected formulas retain cached values only.",
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
        } catch (InvalidDataException exception) when (!IWorkProtobuf.IsLimitException(exception)) {
            declaredTileCount = -1;
        }
        int maximumTileCount = checked((rows + TileRowStride - 1) / TileRowStride);
        if (declaredTileCount > maximumTileCount) {
            references.Declarations.Record(model, "4/3/1", declaredTileCount,
                IWorkSourceDeclarationIssueKind.RejectedMessageSet);
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
        } catch (InvalidDataException exception) when (!IWorkProtobuf.IsLimitException(exception)) {
            tileStorage = null;
        }
        if (tileStorage == null) {
            if (store.HasField(3)) references.Declarations.Record(model, "4/3", store.FieldCount(3));
            MarkTableStorageUnsupported(model, diagnostics, ref supportsEditableReconstruction);
            return CreateTable();
        }
        IReadOnlyList<IWorkWireMessage> tileEntries = IWorkObjectIndex.TryGetMessages(
            tileStorage, 1, out bool malformedTileEntries);
        if (malformedTileEntries) {
            references.Declarations.Record(model, "4/3/1", tileStorage.FieldCount(1),
                IWorkSourceDeclarationIssueKind.RejectedMessageSet);
            MarkTableStorageUnsupported(model, diagnostics, ref supportsEditableReconstruction);
            return CreateTable();
        }
        var tileIndexes = new HashSet<ulong>();
        var tileIdentifiers = new HashSet<ulong>();
        int tileEntryPosition = 0;
        foreach (IWorkWireMessage tileEntry in tileEntries) {
            tileEntryPosition++;
            ulong? declaredTileId = tileEntry.GetUnsigned(1);
            if (tileEntry.FieldCount(1) != 1
                || tileEntry.HasUnexpectedWireKind(1, IWorkWireKind.Varint)
                || !declaredTileId.HasValue || declaredTileId.Value >= (ulong)maximumTileCount) {
                RecordInvalidTileEntry(model, tileEntryPosition, references);
                supportsEditableReconstruction = false;
                diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                    "IWORK_TABLE_TILE_INDEX_UNSUPPORTED",
                    "An iWork table tile index is missing or exceeds the supported range; editable reconstruction is incomplete.",
                    model.EntryPath, model.Identifier));
                continue;
            }
            ulong rawTileId = declaredTileId.Value;
            if (!tileIndexes.Add(rawTileId)) {
                RecordInvalidTileEntry(model, tileEntryPosition, references);
                MarkDuplicateTile(model, diagnostics, ref supportsEditableReconstruction);
                continue;
            }
            IWorkArchiveRecord? tile = references.ReadOne(model, tileEntry, 2,
                "4/3/1[" + tileEntryPosition.ToString(System.Globalization.CultureInfo.InvariantCulture) + "]/2");
            if (tile == null || tile.MessageType != TableTileArchive) {
                supportsEditableReconstruction = false;
                diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                    "IWORK_TABLE_TILE_UNSUPPORTED",
                    "An iWork table references a missing or unsupported tile object; editable reconstruction is incomplete.",
                    model.EntryPath, model.Identifier));
                continue;
            }
            if (!tileIdentifiers.Add(tile.Identifier)) {
                RecordInvalidTileEntry(model, tileEntryPosition, references);
                MarkDuplicateTile(model, diagnostics, ref supportsEditableReconstruction);
                continue;
            }
            long tileStartRow = checked((long)rawTileId * TileRowStride);
            long remainingRows = rows - tileStartRow;
            int maximumRowsInTile = remainingRows <= 0
                ? 0
                : (int)Math.Min(TileRowStride, remainingRows);
            IWorkWireMessage? tileMessage = ReadTileMessage(source, tile, maximumRowsInTile,
                references, diagnostics, ref supportsEditableReconstruction);
            if (tileMessage == null) continue;
            IReadOnlyList<(IWorkWireMessage Message, int Position)> rowsInTile = ReadTileRows(source,
                tile, tileMessage, references, diagnostics, ref supportsEditableReconstruction);
            var rowIndexes = new HashSet<ulong>();
            foreach ((IWorkWireMessage rowInfo, int rowPosition) in rowsInTile) {
                source.CancellationToken.ThrowIfCancellationRequested();
                byte[]? currentBuffer = rowInfo.GetBytes(6);
                byte[]? currentOffsets = rowInfo.GetBytes(7);
                if (rowInfo.FieldCount(1) != 1
                    || rowInfo.HasUnexpectedWireKind(1, IWorkWireKind.Varint)
                    || rowInfo.HasUnexpectedWireKind(3, IWorkWireKind.Bytes)
                    || rowInfo.HasUnexpectedWireKind(4, IWorkWireKind.Bytes)
                    || rowInfo.HasUnexpectedWireKind(6, IWorkWireKind.Bytes)
                    || rowInfo.HasUnexpectedWireKind(7, IWorkWireKind.Bytes)
                    || rowInfo.HasUnexpectedWireKind(8, IWorkWireKind.Varint)
                    || rowInfo.FieldCount(3) > 1
                    || rowInfo.FieldCount(4) > 1
                    || rowInfo.FieldCount(6) > 1
                    || rowInfo.FieldCount(7) > 1
                    || rowInfo.FieldCount(8) > 1
                    || rowInfo.GetUnsigned(8) > 1
                    || (currentBuffer == null) != (currentOffsets == null)
                    || currentOffsets != null && currentOffsets.Length % 2 != 0) {
                    RecordInvalidTileRow(tile, rowPosition, references);
                    MarkCellStorageUnsupported(tile, diagnostics, ref supportsEditableReconstruction);
                    continue;
                }
                bool hasPreBncStorage = (rowInfo.GetBytes(3)?.Length ?? 0) > 0
                    || (rowInfo.GetBytes(4)?.Length ?? 0) > 0;
                if ((currentBuffer == null || currentOffsets == null) && hasPreBncStorage) {
                    references.Declarations.Record(tile, TileRowPath(rowPosition), 1,
                        IWorkSourceDeclarationIssueKind.RejectedMessageSet);
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
                    RecordInvalidTileRow(tile, rowPosition, references);
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
                    RecordInvalidTileRow(tile, rowPosition, references);
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
                    RecordInvalidTileRow(tile, rowPosition, references);
                    MarkCellStorageUnsupported(tile, diagnostics, ref supportsEditableReconstruction);
                }
                int[] populatedOffsets = Enumerable.Range(0, availableColumns)
                    .Select(column => offsets[column * 2] | offsets[column * 2 + 1] << 8)
                    .Where(encodedOffset => encodedOffset != ushort.MaxValue)
                    .Select(encodedOffset => hasWideOffsets ? checked(encodedOffset * 4) : encodedOffset)
                    .ToArray();
                if (populatedOffsets.Length != populatedOffsets.Distinct().Count()) {
                    RecordInvalidTileRow(tile, rowPosition, references);
                    MarkCellStorageUnsupported(tile, diagnostics, ref supportsEditableReconstruction);
                    continue;
                }
                Array.Sort(populatedOffsets);
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
                        strings, richStrings, formulas, source.Options, projectionBudget,
                        formulaRichStringIdentifiers, nonFormulaRichStringIdentifiers);
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

        foreach (var entry in richStrings.OmittedStorages) {
            if (formulaRichStringIdentifiers.Contains(entry.Key) || nonFormulaRichStringIdentifiers.Contains(entry.Key))
                omittedTextUnits.Add(entry.Value);
        }

        bool blockingRichText = !richStrings.StructureComplete || richStrings.Materialized.Any(entry =>
            !entry.Value.IsTextComplete
            && nonFormulaRichStringIdentifiers.Contains(entry.Key));
        if (!richStrings.FullyReconstructed && blockingRichText) {
            supportsEditableReconstruction = false;
            diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                "IWORK_TABLE_RICH_TEXT_STORAGE_UNSUPPORTED",
                "An iWork rich-text table catalog contains malformed or unresolved entries; affected cell text may be incomplete.",
                model.EntryPath, model.Identifier));
        }
        if (richStrings.Materialized.Any(entry => nonFormulaRichStringIdentifiers.Contains(entry.Key)
                && !entry.Value.IsComplete && entry.Value.IsTextComplete)) {
            supportsEditableReconstruction = false;
            diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                "IWORK_TABLE_RICH_TEXT_STYLE_UNSUPPORTED",
                "An iWork rich-text table catalog contains formatting that could not be reconstructed; editable formatting is incomplete.",
                model.EntryPath, model.Identifier));
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
        int incompleteFormulaCacheCount = cells.Count(cell => cell.Kind == IWorkCellKind.Formula
            && !cell.CachedValueIsComplete && cell.Value != null);
        if (incompleteFormulaCacheCount > 0) {
            diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                "IWORK_TABLE_FORMULA_CACHE_PARTIAL",
                $"{incompleteFormulaCacheCount} formula cached values in table '{name}' are partial; only complete expressions can be reconstructed as editable formulas.",
                model.EntryPath, model.Identifier));
            if (cells.Any(cell => cell.Kind == IWorkCellKind.Formula
                && !cell.CachedValueIsComplete && !cell.FormulaIsComplete)) {
                supportsEditableReconstruction = false;
            }
        }
        int incompleteUncachedFormulaCount = cells.Count(cell => cell.Kind == IWorkCellKind.Formula
            && (!cell.FormulaIsComplete || !cell.CachedValueIsComplete) && cell.Value == null);
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
            mergedRanges, geometry, accessibilityDescription, sourceIdentity, omittedTextUnits, rowHeights, columnWidths);
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

}
