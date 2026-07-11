using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
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
            bool useFlatSnapshot = cells.Count - dataStartIndex <= DirectCellValuesFlatSnapshotCellLimit;
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

    }
}
