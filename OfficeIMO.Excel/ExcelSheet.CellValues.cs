using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private const int DirectSequentialCellWriteLimit = 16;
        private const int DirectCellValuesLinearHeaderDuplicateCheckLimit = 32;

        /// <summary>
        /// Writes multiple cell values efficiently, using parallelization when beneficial.
        /// </summary>
        /// <param name="cells">Collection of cell coordinates and values.</param>
        /// <param name="mode">Optional execution mode override.</param>
        /// <param name="ct">Cancellation token.</param>
        /// <remarks>
        /// This is the canonical API for batch cell writes. Use this in place of the older
        /// <see cref="SetCellValues(IEnumerable{ValueTuple{int, int, object}}, ExecutionMode?, CancellationToken)"/>
        /// method, which will be removed in a future release.
        /// </remarks>
        public void CellValues(IEnumerable<(int Row, int Column, object Value)> cells, ExecutionMode? mode = null, CancellationToken ct = default) {
            if (cells is null) {
                throw new ArgumentNullException(nameof(cells));
            }
            var list = cells as IReadOnlyList<(int Row, int Column, object Value)> ?? cells.ToList();
            if (list.Count == 0) return;
            DirectCellValuesSaveCandidate? appendSaveCandidate = null;

            if (!_excelDocument.IsMaterializingDeferredDataSetImport) {
                _excelDocument.MaterializeDeferredDataSetImport();
            }

            DirectCellValuesSaveCandidate? directSaveCandidate = null;
            if (TryCreateDirectCellValuesSaveCandidate(list, mode, out DirectCellValuesSaveCandidate? candidate)
                && candidate != null
                && CanRegisterDirectTabularSaveCandidate(1, 1, candidate.ColumnNames.Length)) {
                directSaveCandidate = candidate;
            }

            if (directSaveCandidate != null
                && RegisterDeferredDirectCellValuesSaveCandidateIfPossible(directSaveCandidate)) {
                return;
            }

            if (appendSaveCandidate == null
                && mode == ExecutionMode.Parallel
                && TryCreateDirectCellValuesAppendCandidate(list, out DirectCellValuesSaveCandidate? appendCandidate)) {
                appendSaveCandidate = appendCandidate;
            }

            if (appendSaveCandidate != null
                && _excelDocument.CanDeferDirectCellValuesAppendCandidate
                && RegisterDeferredDirectCellValuesSaveCandidateIfPossible(appendSaveCandidate)) {
                return;
            }

            // Single cell: trivially sequential
            if (list.Count == 1) {
                var single = list[0];
                CellValue(single.Row, single.Column, single.Value);
                RegisterDirectCellValuesSaveCandidateIfPossible(directSaveCandidate);
                return;
            }

            if (list.Count > DirectSequentialCellWriteLimit && TryApplyPlainCellsByAppendingRows(list, ct)) {
                RegisterDirectCellValuesSaveCandidateIfPossible(appendSaveCandidate ?? directSaveCandidate);
                return;
            }

            // Prepared buffers for parallel scenario
            var prepared = new (int Row, int Col, CellValue Val, EnumValue<DocumentFormat.OpenXml.Spreadsheet.CellValues> Type)[list.Count];
            var ssPlanner = new SharedStringPlanner();

            ExecuteWithPolicy(
                opName: "CellValues",
                itemCount: list.Count,
                overrideMode: mode,
                sequentialCore: () => {
                    if (list.Count <= DirectSequentialCellWriteLimit) {
                        for (int i = 0; i < list.Count; i++) {
                            ct.ThrowIfCancellationRequested();
                            var (r, c, v) = list[i];
                            CellValueCore(r, c, v);
                        }

                        return;
                    }

                    // Sequential path - keep the fast prepared/apply writer so row-major
                    // batches can append rows instead of falling back to GetCell per cell.
                    for (int i = 0; i < list.Count; i++) {
                        ct.ThrowIfCancellationRequested();
                        var (r, c, v) = list[i];
                        var (val, type) = CoerceForCellNoDom(v, ssPlanner);
                        prepared[i] = (r, c, val!, type!);
                    }

                    ssPlanner.ApplyAndFixup(prepared, _excelDocument);
                    ApplyPreparedCells(prepared, list);
                },
                computeParallel: () => {
                    // Parallel compute phase - prepare values without DOM mutation
                    Parallel.For(0, list.Count, new ParallelOptions {
                        CancellationToken = ct,
                        MaxDegreeOfParallelism = EffectiveExecution.MaxDegreeOfParallelism ?? -1
                    }, i => {
                        var (r, c, obj) = list[i];
                        var (val, type) = CoerceForCellNoDom(obj, ssPlanner);
                        prepared[i] = (r, c, val!, type!);
                    });
                },
                applySequential: () => {
                    // Apply phase - first fix shared strings, then write all values to DOM
                    ssPlanner.ApplyAndFixup(prepared, _excelDocument);
                    ApplyPreparedCells(prepared, list);
                },
                ct: ct
            );

            RegisterDirectCellValuesSaveCandidateIfPossible(appendSaveCandidate ?? directSaveCandidate);
        }

        /// <summary>
        /// Compute-only coercion for parallel scenarios. Does not mutate DOM.
        /// Uses <see cref="SharedStringPlanner"/> for string values.
        /// </summary>
        private (CellValue cellValue, EnumValue<DocumentFormat.OpenXml.Spreadsheet.CellValues> dataType) CoerceForCellNoDom(object? value, SharedStringPlanner planner) {
            var dateTimeOffsetStrategy = _excelDocument.DateTimeOffsetWriteStrategy;
            var (cellValue, cellType) = CoerceValueHelper.Coerce(
                value,
                s => {
                    var sanitized = planner.Note(s);
                    return new CellValue(sanitized);
                },
                dateTimeOffsetStrategy);
            return (cellValue, GetCachedDataTableCellType(cellType));
        }

        private void RegisterDirectCellValuesSaveCandidateIfPossible(DirectCellValuesSaveCandidate? candidate) {
            if (candidate == null || string.IsNullOrEmpty(candidate.Range)) {
                return;
            }

            if (candidate.Rows != null) {
                _excelDocument.RegisterDirectTabularSaveCandidate(
                    this,
                    "Cells",
                    candidate.ColumnNames,
                    candidate.ColumnTypes,
                    candidate.Rows,
                    candidate.IncludeHeaders,
                    candidate.Range);
            } else {
                _excelDocument.RegisterDirectCellValuesSaveCandidate(
                    this,
                    "Cells",
                    candidate.ColumnNames,
                    candidate.ColumnTypes,
                    candidate.Values!,
                    candidate.ColumnCount,
                    candidate.RowCount,
                    candidate.ValuesMatchColumnTypes,
                    candidate.IncludeHeaders,
                    candidate.Range);
            }
        }

        private bool RegisterDeferredDirectCellValuesSaveCandidateIfPossible(DirectCellValuesSaveCandidate candidate) {
            if (string.IsNullOrEmpty(candidate.Range)) {
                return false;
            }

            return candidate.Rows != null
                ? _excelDocument.RegisterDeferredDirectTabularSaveCandidate(
                    this,
                    "Cells",
                    candidate.ColumnNames,
                    candidate.ColumnTypes,
                    candidate.Rows,
                    candidate.IncludeHeaders,
                    candidate.Range)
                : _excelDocument.RegisterDeferredDirectCellValuesSaveCandidate(
                    this,
                    "Cells",
                    candidate.ColumnNames,
                    candidate.ColumnTypes,
                    candidate.Values!,
                    candidate.ColumnCount,
                    candidate.RowCount,
                    candidate.ValuesMatchColumnTypes,
                    candidate.IncludeHeaders,
                    candidate.Range);
        }

        private bool TryCreateDirectCellValuesAppendCandidate(IReadOnlyList<(int Row, int Column, object Value)> cells, out DirectCellValuesSaveCandidate? candidate) {
            candidate = null;
            if (!TryGetColumnOneRectangleShape(cells, out int firstRow, out int columnCount)
                || firstRow != 2
                || !TryReadExistingHeadersForDirectAppend(columnCount, out string[] headers)
                || !TryCreateDirectAppendCellValuesSaveCandidate(cells, headers, firstRow, columnCount, out candidate)
                || candidate == null) {
                return false;
            }

            string range = A1.CellReference(1, 1) + ":" + A1.CellReference(firstRow + candidate.RowCount - 1, columnCount);
            candidate = candidate.WithRange(range);
            return true;
        }

        private bool TryReadExistingHeadersForDirectAppend(int columnCount, out string[] headers) {
            headers = Array.Empty<string>();
            if (columnCount <= 0 || _excelDocument.HasPackagePropertiesDirty) {
                return false;
            }

            var sheets = WorkbookRoot.Sheets;
            if (sheets == null) {
                return false;
            }

            Sheet? onlySheet = null;
            foreach (var sheet in sheets.Elements<Sheet>()) {
                if (onlySheet != null) {
                    return false;
                }

                onlySheet = sheet;
            }

            if (onlySheet == null || !ReferenceEquals(onlySheet, SheetElement)) {
                return false;
            }

            if (onlySheet.State != null && onlySheet.State.Value != SheetStateValues.Visible) {
                return false;
            }

            if (WorksheetPart.DrawingsPart != null || WorksheetPart.WorksheetCommentsPart != null || WorksheetPart.ExternalRelationships.Any()) {
                return false;
            }

            if (WorksheetPart.TableDefinitionParts.Any()) {
                return false;
            }

            var worksheet = WorksheetRoot;
            bool foundSheetData = false;
            foreach (var child in worksheet.ChildElements) {
                if (child is SheetDimension) {
                    continue;
                }

                if (child is not SheetData sheetData) {
                    return false;
                }

                if (foundSheetData) {
                    return false;
                }

                if (!TryReadHeadersIfSheetDataContainsOnlyHeaderRow(sheetData, columnCount, out headers)) {
                    return false;
                }

                foundSheetData = true;
            }

            return foundSheetData;
        }

        private bool TryReadHeadersIfSheetDataContainsOnlyHeaderRow(SheetData sheetData, int columnCount, out string[] headers) {
            headers = new string[columnCount];
            Row? headerRow = null;
            foreach (var row in sheetData.Elements<Row>()) {
                if (row.RowIndex == null || row.RowIndex.Value != 1U) {
                    if (row.Elements<Cell>().Any()) {
                        return false;
                    }

                    continue;
                }

                if (headerRow != null) {
                    return false;
                }

                headerRow = row;
            }

            if (headerRow == null) {
                return false;
            }

            HashSet<string>? usedHeaders = columnCount > DirectCellValuesLinearHeaderDuplicateCheckLimit
                ? new HashSet<string>(StringComparer.OrdinalIgnoreCase)
                : null;
            int cellCount = 0;
            foreach (var cell in headerRow.Elements<Cell>()) {
                cellCount++;
                if (cellCount > columnCount) {
                    return false;
                }

                string? reference = cell.CellReference?.Value;
                if (string.IsNullOrEmpty(reference)
                    || A1.ParseColumnIndexFromCellReferenceFast(reference) != cellCount) {
                    return false;
                }

                string header = GetCellText(cell);
                if (string.IsNullOrWhiteSpace(header)) {
                    return false;
                }

                if (usedHeaders != null) {
                    if (!usedHeaders.Add(header)) {
                        return false;
                    }
                } else {
                    for (int headerIndex = 0; headerIndex < headers.Length; headerIndex++) {
                        if (string.Equals(headers[headerIndex], header, StringComparison.OrdinalIgnoreCase)) {
                            return false;
                        }
                    }
                }

                headers[cellCount - 1] = header;
            }

            return cellCount == columnCount;
        }

        private static bool TryCreateDirectAppendCellValuesSaveCandidate(
            IReadOnlyList<(int Row, int Column, object Value)> cells,
            string[] headers,
            int firstRow,
            int columnCount,
            out DirectCellValuesSaveCandidate? candidate) {
            candidate = null;

            if (!TrySnapshotDirectAppendCellValues(cells, firstRow, columnCount, out Type[] columnTypes, out object?[] values, out int rowCount, out bool valuesMatchColumnTypes)) {
                return false;
            }

            candidate = new DirectCellValuesSaveCandidate(headers, columnTypes, values, columnCount, rowCount, valuesMatchColumnTypes, includeHeaders: true, range: string.Empty);
            return true;
        }

        private static bool TryCreateDirectCellValuesSaveCandidate(
            IReadOnlyList<(int Row, int Column, object Value)> cells,
            ExecutionMode? mode,
            out DirectCellValuesSaveCandidate? candidate) {
            candidate = null;

            if (!TryGetA1RectangleShape(cells, out int rowCount, out int columnCount)) {
                return false;
            }

            if (mode == ExecutionMode.Parallel && (rowCount <= 1 || columnCount <= 1)) {
                return false;
            }

            bool includeHeaders = CanTreatFirstCellValuesRowAsHeaders(cells, columnCount, rowCount);
            if (!includeHeaders
                && mode != ExecutionMode.Parallel
                && FirstCellValuesRowLooksLikeHeaderText(cells, columnCount)) {
                return false;
            }

            int dataStartIndex = includeHeaders ? columnCount : 0;
            bool useFlatSnapshot = !includeHeaders && columnCount <= 3;
            Type[] columnTypes;
            object?[]? values = null;
            object?[][]? rows = null;
            bool valuesMatchColumnTypes = false;
            int dataRowCount;
            if (useFlatSnapshot) {
                if (!TrySnapshotDirectCellValues(cells, dataStartIndex, columnCount, validateA1Rectangle: true, out columnTypes, out values, out dataRowCount, out valuesMatchColumnTypes)) {
                    return false;
                }
            } else if (TryCreateDirectCellValuesRowsAndColumnTypes(cells, dataStartIndex, columnCount, validateA1Rectangle: true, out columnTypes, out rows)) {
                dataRowCount = rows.Length;
            } else {
                return false;
            }

            var columnNames = new string[columnCount];
            for (int column = 0; column < columnCount; column++) {
                columnNames[column] = includeHeaders
                    ? Convert.ToString(cells[column].Value, CultureInfo.InvariantCulture) ?? string.Empty
                    : "Column" + (column + 1).ToString(CultureInfo.InvariantCulture);
            }

            string range = A1.CellReference(1, 1) + ":" + A1.CellReference(rowCount, columnCount);
            candidate = values != null
                ? new DirectCellValuesSaveCandidate(columnNames, columnTypes, values, columnCount, dataRowCount, valuesMatchColumnTypes, includeHeaders, range)
                : new DirectCellValuesSaveCandidate(columnNames, columnTypes, rows!, includeHeaders, range);
            return true;
        }

        private sealed class DirectCellValuesSaveCandidate {
            internal DirectCellValuesSaveCandidate(string[] columnNames, Type[] columnTypes, object?[][] rows, bool includeHeaders, string range) {
                ColumnNames = columnNames;
                ColumnTypes = columnTypes;
                Rows = rows;
                ColumnCount = columnNames.Length;
                RowCount = rows.Length;
                ValuesMatchColumnTypes = false;
                IncludeHeaders = includeHeaders;
                Range = range;
            }

            internal DirectCellValuesSaveCandidate(
                string[] columnNames,
                Type[] columnTypes,
                object?[] values,
                int columnCount,
                int rowCount,
                bool valuesMatchColumnTypes,
                bool includeHeaders,
                string range) {
                ColumnNames = columnNames;
                ColumnTypes = columnTypes;
                Values = values;
                ColumnCount = columnCount;
                RowCount = rowCount;
                ValuesMatchColumnTypes = valuesMatchColumnTypes;
                IncludeHeaders = includeHeaders;
                Range = range;
            }

            internal string[] ColumnNames { get; }

            internal Type[] ColumnTypes { get; }

            internal object?[][]? Rows { get; }

            internal object?[]? Values { get; }

            internal int ColumnCount { get; }

            internal int RowCount { get; }

            internal bool ValuesMatchColumnTypes { get; }

            internal bool IncludeHeaders { get; }

            internal string Range { get; }

            internal DirectCellValuesSaveCandidate WithRange(string range) {
                return Rows != null
                    ? new DirectCellValuesSaveCandidate(ColumnNames, ColumnTypes, Rows, IncludeHeaders, range)
                    : new DirectCellValuesSaveCandidate(ColumnNames, ColumnTypes, Values!, ColumnCount, RowCount, ValuesMatchColumnTypes, IncludeHeaders, range);
            }
        }

        private static bool TryGetA1RectangleShape(IReadOnlyList<(int Row, int Column, object Value)> cells, out int rowCount, out int columnCount) {
            rowCount = 0;
            columnCount = 0;
            if (cells.Count == 0 || cells[0].Row != 1 || cells[0].Column != 1) {
                return false;
            }

            while (columnCount < cells.Count && cells[columnCount].Row == 1) {
                if (cells[columnCount].Column != columnCount + 1) {
                    return false;
                }

                columnCount++;
            }

            if (columnCount == 0 || cells.Count % columnCount != 0) {
                return false;
            }

            rowCount = cells.Count / columnCount;
            return true;
        }

        private static bool TryGetColumnOneRectangleShape(
            IReadOnlyList<(int Row, int Column, object Value)> cells,
            out int firstRow,
            out int columnCount) {
            firstRow = 0;
            columnCount = 0;
            if (cells.Count == 0 || cells[0].Row <= 1 || cells[0].Column != 1) {
                return false;
            }

            firstRow = cells[0].Row;
            while (columnCount < cells.Count && cells[columnCount].Row == firstRow) {
                if (cells[columnCount].Column != columnCount + 1) {
                    return false;
                }

                columnCount++;
            }

            if (columnCount == 0 || cells.Count % columnCount != 0) {
                return false;
            }

            return true;
        }

        private static bool CanTreatFirstCellValuesRowAsHeaders(IReadOnlyList<(int Row, int Column, object Value)> cells, int columnCount, int rowCount) {
            if (rowCount < 2) {
                return false;
            }

            if (columnCount <= DirectCellValuesLinearHeaderDuplicateCheckLimit) {
                for (int column = 0; column < columnCount; column++) {
                    if (cells[column].Value is not string header
                        || string.IsNullOrWhiteSpace(header)
                        || IsDirectCellValuesAutomaticFormattingText(header)) {
                        return false;
                    }

                    for (int previousColumn = 0; previousColumn < column; previousColumn++) {
                        if (string.Equals((string)cells[previousColumn].Value, header, StringComparison.OrdinalIgnoreCase)) {
                            return false;
                        }
                    }
                }

                return true;
            }

            var headers = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (int column = 0; column < columnCount; column++) {
                if (cells[column].Value is not string header
                    || string.IsNullOrWhiteSpace(header)
                    || IsDirectCellValuesAutomaticFormattingText(header)
                    || !headers.Add(header)) {
                    return false;
                }
            }

            return true;
        }

        private static bool FirstCellValuesRowLooksLikeHeaderText(IReadOnlyList<(int Row, int Column, object Value)> cells, int columnCount) {
            if (columnCount <= 0 || cells.Count < columnCount) {
                return false;
            }

            for (int column = 0; column < columnCount; column++) {
                if (cells[column].Value is not string header || string.IsNullOrWhiteSpace(header)) {
                    return false;
                }
            }

            return true;
        }

        private static bool TryCreateDirectCellValuesRowsAndColumnTypes(
            IReadOnlyList<(int Row, int Column, object Value)> cells,
            int dataStartIndex,
            int columnCount,
            bool validateA1Rectangle,
            out Type[] columnTypes,
            out object?[][] rows) {
            columnTypes = new Type[columnCount];
            int rowCount = (cells.Count - dataStartIndex) / columnCount;
            rows = new object?[rowCount][];
            var inferredTypes = new Type?[columnCount];

            int rowOffset = 0;
            for (int index = dataStartIndex; index < cells.Count; index += columnCount) {
                var row = new object?[columnCount];
                for (int column = 0; column < columnCount; column++) {
                    int cellIndex = index + column;
                    if (validateA1Rectangle) {
                        var cell = cells[cellIndex];
                        int expectedRow = (cellIndex / columnCount) + 1;
                        int expectedColumn = (cellIndex % columnCount) + 1;
                        if (cell.Row != expectedRow || cell.Column != expectedColumn) {
                            return false;
                        }
                    }

                    object? value = cells[cellIndex].Value;
                    if (value is string text && IsDirectCellValuesAutomaticFormattingText(text)) {
                        return false;
                    }

                    if (IsDirectCellValuesBlankValue(value)) {
                        row[column] = value;
                        continue;
                    }

                    Type valueType = GetDirectCellValuesColumnType(value!);
                    Type? inferred = inferredTypes[column];
                    if (inferred == null) {
                        inferredTypes[column] = valueType;
                        row[column] = value;
                        continue;
                    }

                    if (inferred == typeof(object)) {
                        row[column] = value;
                        continue;
                    }

                    if (inferred != valueType) {
                        if (IsDirectCellValuesStyleSensitiveType(inferred) || IsDirectCellValuesStyleSensitiveType(valueType)) {
                            return false;
                        }

                        inferredTypes[column] = typeof(object);
                    }

                    row[column] = value;
                }

                rows[rowOffset++] = row;
            }

            for (int column = 0; column < columnCount; column++) {
                columnTypes[column] = inferredTypes[column] ?? typeof(string);
            }

            return true;
        }

        private static bool TrySnapshotDirectCellValues(
            IReadOnlyList<(int Row, int Column, object Value)> cells,
            int dataStartIndex,
            int columnCount,
            bool validateA1Rectangle,
            out Type[] columnTypes,
            out object?[] values,
            out int rowCount,
            out bool valuesMatchColumnTypes) {
            columnTypes = new Type[columnCount];
            rowCount = (cells.Count - dataStartIndex) / columnCount;
            values = new object?[rowCount * columnCount];
            var inferredTypes = new Type?[columnCount];
            valuesMatchColumnTypes = true;

            int valueIndex = 0;
            for (int index = dataStartIndex; index < cells.Count; index += columnCount) {
                for (int column = 0; column < columnCount; column++) {
                    int cellIndex = index + column;
                    if (validateA1Rectangle) {
                        var cell = cells[cellIndex];
                        int expectedRow = (cellIndex / columnCount) + 1;
                        int expectedColumn = (cellIndex % columnCount) + 1;
                        if (cell.Row != expectedRow || cell.Column != expectedColumn) {
                            return false;
                        }
                    }

                    object? value = cells[cellIndex].Value;
                    if (value is string text && IsDirectCellValuesAutomaticFormattingText(text)) {
                        return false;
                    }

                    if (IsDirectCellValuesBlankValue(value)) {
                        valuesMatchColumnTypes = false;
                        values[valueIndex++] = value;
                        continue;
                    }

                    Type valueType = GetDirectCellValuesColumnType(value!);
                    Type? inferred = inferredTypes[column];
                    if (inferred == null) {
                        inferredTypes[column] = valueType;
                        values[valueIndex++] = value;
                        continue;
                    }

                    if (inferred == typeof(object)) {
                        values[valueIndex++] = value;
                        continue;
                    }

                    if (inferred != valueType) {
                        if (IsDirectCellValuesStyleSensitiveType(inferred) || IsDirectCellValuesStyleSensitiveType(valueType)) {
                            return false;
                        }

                        inferredTypes[column] = typeof(object);
                        valuesMatchColumnTypes = false;
                    }

                    values[valueIndex++] = value;
                }
            }

            for (int column = 0; column < columnCount; column++) {
                if (inferredTypes[column] == null || inferredTypes[column] == typeof(object)) {
                    valuesMatchColumnTypes = false;
                }

                columnTypes[column] = inferredTypes[column] ?? typeof(string);
            }

            return true;
        }

        private static bool TrySnapshotDirectAppendCellValues(
            IReadOnlyList<(int Row, int Column, object Value)> cells,
            int firstRow,
            int columnCount,
            out Type[] columnTypes,
            out object?[] values,
            out int rowCount,
            out bool valuesMatchColumnTypes) {
            columnTypes = new Type[columnCount];
            rowCount = cells.Count / columnCount;
            values = new object?[cells.Count];
            var inferredTypes = new Type?[columnCount];
            valuesMatchColumnTypes = true;

            int index = 0;
            for (int rowOffset = 0; rowOffset < rowCount; rowOffset++) {
                int expectedRow = firstRow + rowOffset;
                for (int column = 0; column < columnCount; column++, index++) {
                    var cell = cells[index];
                    if (cell.Row != expectedRow || cell.Column != column + 1) {
                        return false;
                    }

                    object? value = cell.Value;
                    if (value is string text && IsDirectCellValuesAutomaticFormattingText(text)) {
                        return false;
                    }

                    if (IsDirectCellValuesBlankValue(value)) {
                        valuesMatchColumnTypes = false;
                        values[index] = value;
                        continue;
                    }

                    Type valueType = GetDirectCellValuesColumnType(value!);
                    Type? inferred = inferredTypes[column];
                    if (inferred == null) {
                        inferredTypes[column] = valueType;
                        values[index] = value;
                        continue;
                    }

                    if (inferred == typeof(object)) {
                        values[index] = value;
                        continue;
                    }

                    if (inferred != valueType) {
                        if (IsDirectCellValuesStyleSensitiveType(inferred) || IsDirectCellValuesStyleSensitiveType(valueType)) {
                            return false;
                        }

                        inferredTypes[column] = typeof(object);
                        valuesMatchColumnTypes = false;
                    }

                    values[index] = value;
                }
            }

            for (int column = 0; column < columnCount; column++) {
                if (inferredTypes[column] == null || inferredTypes[column] == typeof(object)) {
                    valuesMatchColumnTypes = false;
                }

                columnTypes[column] = inferredTypes[column] ?? typeof(string);
            }

            return true;
        }

        private static bool IsDirectCellValuesAutomaticFormattingText(string text)
            => text.IndexOf('\r') >= 0 || text.IndexOf('\n') >= 0;

        private static Type NormalizeDirectCellValuesColumnType(Type type) {
            if (type == typeof(DBNull) || type == typeof(void)) {
                return typeof(object);
            }

            return type;
        }

        private static Type GetDirectCellValuesColumnType(object value) {
            switch (value) {
                case string:
                    return typeof(string);
                case bool:
                    return typeof(bool);
                case DateTime:
                    return typeof(DateTime);
                case DateTimeOffset:
                    return typeof(DateTimeOffset);
                case TimeSpan:
                    return typeof(TimeSpan);
                case double:
                    return typeof(double);
                case float:
                    return typeof(float);
                case decimal:
                    return typeof(decimal);
                case sbyte:
                    return typeof(sbyte);
                case byte:
                    return typeof(byte);
                case short:
                    return typeof(short);
                case ushort:
                    return typeof(ushort);
                case int:
                    return typeof(int);
                case uint:
                    return typeof(uint);
                case long:
                    return typeof(long);
                case ulong:
                    return typeof(ulong);
#if NET6_0_OR_GREATER
                case DateOnly:
                    return typeof(DateOnly);
                case TimeOnly:
                    return typeof(TimeOnly);
#endif
                default:
                    return NormalizeDirectCellValuesColumnType(value.GetType());
            }
        }

        private static bool IsDirectCellValuesStyleSensitiveType(Type type) {
            return type == typeof(DateTime)
                || type == typeof(DateTimeOffset)
                || type == typeof(TimeSpan)
#if NET6_0_OR_GREATER
                || type == typeof(DateOnly)
                || type == typeof(TimeOnly)
#endif
                ;
        }

        private static bool IsDirectCellValuesBlankValue(object? value) {
            return value == null || value == DBNull.Value;
        }

        private bool TryApplyPlainCellsByAppendingRows(IReadOnlyList<(int Row, int Column, object Value)> source, CancellationToken ct) {
            bool applied = false;
            System.Threading.ReaderWriterLockSlim? lck = _excelDocument._lock;
            if (lck == null) {
                try { lck = _excelDocument.EnsureLock(); } catch { lck = null; }
            }

            Locking.ExecuteWrite(lck, () => applied = TryApplyPlainCellsByAppendingRowsCore(source, ct));
            return applied;
        }

        private bool TryApplyPlainCellsByAppendingRowsCore(IReadOnlyList<(int Row, int Column, object Value)> source, CancellationToken ct) {
            if (!TryGetPlainAppendLayout(source, out int firstRow, out int minColumn, out int maxColumn)) {
                return false;
            }

            var sheetData = GetOrCreateSheetData();
            int minExistingRow = int.MaxValue;
            int minExistingColumn = int.MaxValue;
            int maxExistingRow = 0;
            int maxExistingColumn = 0;
            foreach (var existingRow in sheetData.Elements<Row>()) {
                if (existingRow.RowIndex == null) {
                    return false;
                }

                if (existingRow.RowIndex != null && existingRow.RowIndex.Value >= (uint)firstRow) {
                    return false;
                }

                if (!existingRow.HasChildren) {
                    continue;
                }

                int existingRowIndex = checked((int)(existingRow.RowIndex?.Value ?? 0U));
                if (existingRowIndex <= 0) {
                    continue;
                }

                foreach (var existingCell in existingRow.Elements<Cell>()) {
                    int existingColumnIndex = 0;
                    string? reference = existingCell.CellReference?.Value;
                    if (!string.IsNullOrEmpty(reference)) {
                        existingColumnIndex = A1.ParseColumnIndexFromCellReference(reference!);
                    }

                    if (existingColumnIndex <= 0) {
                        continue;
                    }

                    if (existingRowIndex < minExistingRow) minExistingRow = existingRowIndex;
                    if (existingRowIndex > maxExistingRow) maxExistingRow = existingRowIndex;
                    if (existingColumnIndex < minExistingColumn) minExistingColumn = existingColumnIndex;
                    if (existingColumnIndex > maxExistingColumn) maxExistingColumn = existingColumnIndex;
                }
            }

            var columnNames = new string[maxColumn + 1];
            for (int column = 1; column <= maxColumn; column++) {
                columnNames[column] = GetColumnName(column);
            }

            Dictionary<string, int>? sharedStringIndexes = null;
            bool useDirectStringCells = source.Count >= 4096 && maxColumn > 1;
            List<Row> pendingRows = new List<Row>();
            Row? row = null;
            int rowIndex = 0;
            string rowReference = string.Empty;
            bool canCancel = ct.CanBeCanceled;

            for (int i = 0; i < source.Count; i++) {
                if (canCancel) {
                    ct.ThrowIfCancellationRequested();
                }

                var item = source[i];

                if (item.Row != rowIndex) {
                    if (row != null) {
                        pendingRows.Add(row);
                    }

                    rowIndex = item.Row;
                    rowReference = InvariantNumberText.Get(rowIndex);
                    row = new Row { RowIndex = (uint)rowIndex };
                }

                var (cellValue, cellType) = CoercePlainAppendValue(item.Value, ref sharedStringIndexes, useDirectStringCells);
                row!.Append(CreateTabularAppendCell(columnNames[item.Column] + rowReference, cellValue, cellType));
            }

            if (row != null) {
                pendingRows.Add(row);
            }

            foreach (var pendingRow in pendingRows) {
                sheetData.Append(pendingRow);
            }

            ClearHeaderCacheForPreparedAppend();
            int lastRow = source[source.Count - 1].Row;
            int dimensionMinRow = minExistingRow == int.MaxValue ? firstRow : Math.Min(minExistingRow, firstRow);
            int dimensionMinColumn = minExistingColumn == int.MaxValue ? minColumn : Math.Min(minExistingColumn, minColumn);
            int dimensionMaxRow = Math.Max(maxExistingRow, lastRow);
            int dimensionMaxColumn = Math.Max(maxExistingColumn, maxColumn);
            SetSheetDimensionReference(dimensionMinRow, dimensionMinColumn, dimensionMaxRow, dimensionMaxColumn);
            _requiresSavePreparation = false;
            return true;
        }

        private void ClearHeaderCacheForPreparedAppend() {
            _hasWorksheetMutations = true;
            _excelDocument.MarkPackageDirty();
            ClearCellTextSharedStringCache();
            lock (_headerMapLock) {
                _headerMapCache = null;
                _headerMapSourceA1 = null;
            }
        }

        private bool TryGetPlainAppendLayout(
            IReadOnlyList<(int Row, int Column, object Value)> source,
            out int firstRow,
            out int minColumn,
            out int maxColumn) {
            firstRow = source[0].Row;
            minColumn = int.MaxValue;
            maxColumn = 0;
            int currentRow = 0;
            int currentColumn = 0;

            for (int i = 0; i < source.Count; i++) {
                var item = source[i];
                if (item.Row <= 0 || item.Column <= 0 || item.Column > A1.MaxColumns || item.Row < currentRow) {
                    return false;
                }

                if (!CanAppendPlainValueDirectly(item.Value)) {
                    return false;
                }

                if (item.Row != currentRow) {
                    currentRow = item.Row;
                    currentColumn = 0;
                }

                if (item.Column <= currentColumn) {
                    return false;
                }

                currentColumn = item.Column;
                if (item.Column < minColumn) {
                    minColumn = item.Column;
                }

                if (item.Column > maxColumn) {
                    maxColumn = item.Column;
                }
            }

            if (minColumn == int.MaxValue) {
                minColumn = 1;
            }

            return true;
        }

        private void SetSheetDimensionReference(int minRow, int minColumn, int maxRow, int maxColumn) {
            var worksheet = WorksheetRoot;
            SheetDimension? dimension = null;
            List<SheetDimension>? extraDimensions = null;
            foreach (var currentDimension in worksheet.Elements<SheetDimension>()) {
                if (dimension == null) {
                    dimension = currentDimension;
                    continue;
                }

                extraDimensions ??= new List<SheetDimension>();
                extraDimensions.Add(currentDimension);
            }

            if (extraDimensions != null) {
                for (int i = 0; i < extraDimensions.Count; i++) {
                    extraDimensions[i].Remove();
                }
            }

            string start = A1.CellReference(minRow, minColumn);
            string end = A1.CellReference(maxRow, maxColumn);
            string reference = start == end ? start : start + ":" + end;
            if (dimension == null) {
                InsertSheetDimensionInSchemaOrder(worksheet, new SheetDimension { Reference = reference });
            } else {
                dimension.Reference = reference;
            }
        }

        private static bool CanAppendPlainValueDirectly(object? value) {
            switch (value) {
                case null:
                case DBNull:
                case double:
                case float:
                case decimal:
                case int:
                case long:
                case bool:
                case uint:
                case ulong:
                case ushort:
                case byte:
                case sbyte:
                case short:
                case Guid:
                case Enum:
                case char:
                case Uri:
                    return true;
                case string text:
                    if (text.IndexOf('\r') >= 0 || text.IndexOf('\n') >= 0) {
                        return false;
                    }

                    CoerceValueHelper.ValidateSharedStringLength(text, nameof(value));
                    return true;
                default:
                    return false;
            }
        }

        private (CellValue cellValue, DocumentFormat.OpenXml.Spreadsheet.CellValues cellType) CoercePlainAppendValue(
            object? value,
            ref Dictionary<string, int>? sharedStringIndexes,
            bool useDirectStringCells) {
            (CellValue cellValue, DocumentFormat.OpenXml.Spreadsheet.CellValues cellType) = value switch {
                null => CoerceValueHelper.HandleEmptyString(),
                DBNull => CoerceValueHelper.HandleEmptyString(),
                string text when text.Length == 0 => CoerceValueHelper.HandleEmptyString(),
                string text => useDirectStringCells
                    ? (CreatePrevalidatedPlainAppendStringValue(text), DocumentFormat.OpenXml.Spreadsheet.CellValues.String)
                    : (CreatePlainAppendSharedStringValue(text, ref sharedStringIndexes), DocumentFormat.OpenXml.Spreadsheet.CellValues.SharedString),
                double number => CoerceValueHelper.HandleNumber(number),
                float number => CoerceValueHelper.HandleNumber((double)number),
                decimal number => CoerceValueHelper.HandleDecimal(number),
                int number => CoerceValueHelper.HandleSignedInteger(number),
                long number => CoerceValueHelper.HandleSignedInteger(number),
                bool flag => CoerceValueHelper.HandleBoolean(flag),
                uint number => CoerceValueHelper.HandleUnsignedInteger(number),
                ulong number => CoerceValueHelper.HandleUnsignedInteger(number),
                ushort number => CoerceValueHelper.HandleUnsignedInteger(number),
                byte number => CoerceValueHelper.HandleUnsignedInteger(number),
                sbyte number => CoerceValueHelper.HandleSignedInteger(number),
                short number => CoerceValueHelper.HandleSignedInteger(number),
                Guid guid => useDirectStringCells
                    ? (CreatePlainAppendStringValue(guid.ToString()), DocumentFormat.OpenXml.Spreadsheet.CellValues.String)
                    : (CreatePlainAppendSharedStringValue(guid.ToString(), ref sharedStringIndexes), DocumentFormat.OpenXml.Spreadsheet.CellValues.SharedString),
                Enum enumValue => useDirectStringCells
                    ? (CreatePlainAppendStringValue(enumValue.ToString()), DocumentFormat.OpenXml.Spreadsheet.CellValues.String)
                    : (CreatePlainAppendSharedStringValue(enumValue.ToString(), ref sharedStringIndexes), DocumentFormat.OpenXml.Spreadsheet.CellValues.SharedString),
                char character => useDirectStringCells
                    ? (CreatePlainAppendStringValue(character.ToString()), DocumentFormat.OpenXml.Spreadsheet.CellValues.String)
                    : (CreatePlainAppendSharedStringValue(character.ToString(), ref sharedStringIndexes), DocumentFormat.OpenXml.Spreadsheet.CellValues.SharedString),
                Uri uri => useDirectStringCells
                    ? (CreatePlainAppendStringValue(uri.ToString()), DocumentFormat.OpenXml.Spreadsheet.CellValues.String)
                    : (CreatePlainAppendSharedStringValue(uri.ToString(), ref sharedStringIndexes), DocumentFormat.OpenXml.Spreadsheet.CellValues.SharedString),
                _ => throw new InvalidOperationException("Unsupported direct append value.")
            };

            return (cellValue, cellType);
        }

        private static CellValue CreatePlainAppendStringValue(string text) {
            CoerceValueHelper.ValidateSharedStringLength(text, nameof(text));
            return new CellValue(Utilities.ExcelSanitizer.SanitizeString(text));
        }

        private static CellValue CreatePrevalidatedPlainAppendStringValue(string text) {
            return new CellValue(Utilities.ExcelSanitizer.SanitizeString(text));
        }

        private CellValue CreatePlainAppendSharedStringValue(string text, ref Dictionary<string, int>? sharedStringIndexes) {
            string sanitized = Utilities.ExcelSanitizer.SanitizeString(text);
            sharedStringIndexes ??= new Dictionary<string, int>(StringComparer.Ordinal);
            if (!sharedStringIndexes.TryGetValue(sanitized, out int index)) {
                index = _excelDocument.GetSharedStringIndex(sanitized);
                sharedStringIndexes[sanitized] = index;
            }

            return new CellValue(SharedStringIndexText.Get(index));
        }

        /// <summary>
        /// Obsolete. Use <see cref="CellValues(IEnumerable{ValueTuple{int, int, object}}, ExecutionMode?, CancellationToken)"/> instead.
        /// </summary>
        [Obsolete("Use CellValues(...) instead.")]
        public void SetCellValues(IEnumerable<(int Row, int Column, object Value)> cells, ExecutionMode? mode = null, CancellationToken ct = default) {
            CellValues(cells, mode, ct);
        }

        private void ApplyPreparedCells(
            (int Row, int Col, CellValue Val, EnumValue<DocumentFormat.OpenXml.Spreadsheet.CellValues> Type)[] prepared,
            IReadOnlyList<(int Row, int Column, object Value)> source) {
            if (TryApplyPreparedCellsByAppendingRows(prepared, source)) {
                return;
            }

            var writer = new BatchCellWriter(this);

            for (int i = 0; i < prepared.Length; i++) {
                var p = prepared[i];
                var originalValue = source[i].Value;
                var cell = writer.GetOrCreateCell(p.Row, p.Col);
                cell.CellValue = p.Val;
                cell.DataType = p.Type;
                ApplyAutomaticCellFormatting(cell, originalValue, p.Type);
            }

            ClearHeaderCache();
        }

        private bool TryApplyPreparedCellsByAppendingRows(
            (int Row, int Col, CellValue Val, EnumValue<DocumentFormat.OpenXml.Spreadsheet.CellValues> Type)[] prepared,
            IReadOnlyList<(int Row, int Column, object Value)> source) {
            if (prepared.Length != source.Count) {
                return false;
            }

            if (prepared.Length == 0) {
                ClearHeaderCache();
                return true;
            }

            int firstRow = prepared[0].Row;
            int currentRow = 0;
            int currentColumn = 0;
            int maxColumn = 0;

            for (int i = 0; i < prepared.Length; i++) {
                var p = prepared[i];
                if (p.Row <= 0 || p.Col <= 0 || p.Col > A1.MaxColumns || p.Row < currentRow) {
                    return false;
                }

                if (p.Row != currentRow) {
                    currentRow = p.Row;
                    currentColumn = 0;
                }

                if (p.Col <= currentColumn) {
                    return false;
                }

                currentColumn = p.Col;
                if (p.Col > maxColumn) {
                    maxColumn = p.Col;
                }
            }

            var sheetData = GetOrCreateSheetData();
            foreach (var existingRow in sheetData.Elements<Row>()) {
                if (existingRow.RowIndex == null) {
                    return false;
                }

                if (existingRow.RowIndex != null && existingRow.RowIndex.Value >= (uint)firstRow) {
                    return false;
                }
            }

            bool needsAutomaticFormatting = false;
            for (int i = 0; i < source.Count; i++) {
                if (RequiresAutomaticCellFormatting(source[i].Value, prepared[i].Type)) {
                    needsAutomaticFormatting = true;
                    break;
                }
            }

            var columnNames = new string[maxColumn + 1];
            for (int column = 1; column <= maxColumn; column++) {
                columnNames[column] = GetColumnName(column);
            }

            var baseStyleIndexes = needsAutomaticFormatting
                ? GetAppendBaseStyleIndexes(sheetData, firstRow, maxColumn)
                : null;
            Row? row = null;
            int rowIndex = 0;
            string rowReference = string.Empty;
            Dictionary<uint, uint>? appendedDateStyleIndexes = null;
            Dictionary<uint, uint>? appendedDurationStyleIndexes = null;
            Dictionary<uint, uint>? appendedWrapStyleIndexes = null;
            for (int i = 0; i < prepared.Length; i++) {
                var p = prepared[i];
                if (p.Row != rowIndex) {
                    rowIndex = p.Row;
                    rowReference = InvariantNumberText.Get(rowIndex);
                    row = new Row { RowIndex = (uint)rowIndex };
                    sheetData.Append(row);
                }

                var cell = new Cell {
                    CellReference = columnNames[p.Col] + rowReference,
                    CellValue = p.Val,
                    DataType = p.Type
                };

                row!.Append(cell);
                if (needsAutomaticFormatting) {
                    ApplyAutomaticCellFormattingForAppendedCell(
                        cell,
                        source[i].Value,
                        p.Type,
                        baseStyleIndexes![p.Col] ?? 0U,
                        ref appendedDateStyleIndexes,
                        ref appendedDurationStyleIndexes,
                        ref appendedWrapStyleIndexes);
                }
            }

            ClearHeaderCache();
            return true;
        }

        private static uint?[] GetAppendBaseStyleIndexes(SheetData sheetData, int firstRow, int maxColumn) {
            var baseStyleIndexes = new uint?[maxColumn + 1];
            var baseStyleRows = new int[maxColumn + 1];

            foreach (var existingRow in sheetData.Elements<Row>()) {
                if (existingRow.RowIndex == null) {
                    continue;
                }

                int existingRowIndex = (int)existingRow.RowIndex.Value;
                if (existingRowIndex >= firstRow) {
                    continue;
                }

                foreach (var existingCell in existingRow.Elements<Cell>()) {
                    if (existingCell.CellReference == null || existingCell.StyleIndex == null) {
                        continue;
                    }

                    int columnIndex = A1.ParseColumnIndexFromCellReference(existingCell.CellReference.Value);
                    if (columnIndex <= 0 || columnIndex > maxColumn) {
                        continue;
                    }

                    if (existingRowIndex >= baseStyleRows[columnIndex]) {
                        baseStyleRows[columnIndex] = existingRowIndex;
                        baseStyleIndexes[columnIndex] = existingCell.StyleIndex.Value;
                    }
                }
            }

            return baseStyleIndexes;
        }

        private sealed class BatchCellWriter {
            private readonly ExcelSheet _sheet;
            private readonly SheetData _sheetData;
            private readonly Dictionary<int, BatchRowState> _rows;
            private Row? _lastRow;
            private int _lastRowIndex;

            internal BatchCellWriter(ExcelSheet sheet) {
                _sheet = sheet;
                _sheetData = sheet.GetOrCreateSheetData();
                _rows = new Dictionary<int, BatchRowState>();

                foreach (var row in _sheetData.Elements<Row>()) {
                    if (row.RowIndex == null) {
                        continue;
                    }

                    _rows[(int)row.RowIndex.Value] = new BatchRowState(_sheet, row);
                    if (row.RowIndex.Value >= _lastRowIndex) {
                        _lastRowIndex = (int)row.RowIndex.Value;
                        _lastRow = row;
                    }
                }
            }

            internal Cell GetOrCreateCell(int rowIndex, int columnIndex) {
                if (!_rows.TryGetValue(rowIndex, out BatchRowState? rowState)) {
                    var row = GetOrCreateRow(rowIndex);
                    rowState = new BatchRowState(_sheet, row);
                    _rows[rowIndex] = rowState;
                }

                return rowState.GetOrCreateCell(columnIndex, rowIndex);
            }

            private Row GetOrCreateRow(int rowIndex) {
                if (_lastRow != null && rowIndex > _lastRowIndex) {
                    var appended = new Row { RowIndex = (uint)rowIndex };
                    _sheetData.Append(appended);
                    _lastRow = appended;
                    _lastRowIndex = rowIndex;
                    return appended;
                }

                var row = _sheet.GetOrCreateRowElement(_sheetData, rowIndex);
                if (row.RowIndex != null && row.RowIndex.Value >= _lastRowIndex) {
                    _lastRow = row;
                    _lastRowIndex = (int)row.RowIndex.Value;
                }

                return row;
            }

            private sealed class BatchRowState {
                private readonly ExcelSheet _sheet;
                private readonly Row _row;
                private readonly Dictionary<int, Cell> _cells;
                private Cell? _lastCell;
                private int _lastColumnIndex;

                internal BatchRowState(ExcelSheet sheet, Row row) {
                    _sheet = sheet;
                    _row = row;
                    _cells = new Dictionary<int, Cell>();

                    foreach (var cell in row.Elements<Cell>()) {
                        var reference = cell.CellReference?.Value;
                        if (string.IsNullOrEmpty(reference)) {
                            continue;
                        }

                        int columnIndex = GetColumnIndex(reference!);
                        _cells[columnIndex] = cell;

                        if (columnIndex >= _lastColumnIndex) {
                            _lastColumnIndex = columnIndex;
                            _lastCell = cell;
                        }
                    }
                }

                internal Cell GetOrCreateCell(int columnIndex, int rowIndex) {
                    if (_cells.TryGetValue(columnIndex, out Cell? existing)) {
                        return existing;
                    }

                    string cellReference = _sheet.BuildCellReference(rowIndex, columnIndex);
                    var cell = new Cell { CellReference = cellReference };

                    if (_lastCell == null) {
                        var firstCell = _row.Elements<Cell>().FirstOrDefault();
                        if (firstCell != null) {
                            _row.InsertBefore(cell, firstCell);
                        } else {
                            _row.Append(cell);
                        }
                    } else if (columnIndex > _lastColumnIndex) {
                        _row.InsertAfter(cell, _lastCell);
                    } else {
                        Cell? insertAfter = null;
                        foreach (var existingCell in _row.Elements<Cell>()) {
                            var existingReference = existingCell.CellReference?.Value;
                            if (string.IsNullOrEmpty(existingReference)) {
                                continue;
                            }

                            int existingColumnIndex = GetColumnIndex(existingReference!);
                            if (existingColumnIndex > columnIndex) {
                                _row.InsertBefore(cell, existingCell);
                                _cells[columnIndex] = cell;
                                return cell;
                            }

                            insertAfter = existingCell;
                        }

                        if (insertAfter != null) {
                            _row.InsertAfter(cell, insertAfter);
                        } else {
                            _row.Append(cell);
                        }
                    }

                    _cells[columnIndex] = cell;
                    if (columnIndex >= _lastColumnIndex) {
                        _lastColumnIndex = columnIndex;
                        _lastCell = cell;
                    }

                    return cell;
                }
            }
        }
    }
}
