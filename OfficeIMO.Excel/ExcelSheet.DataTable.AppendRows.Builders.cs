using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Data;
using System.Globalization;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private Row CreateDataTableHeaderRow(
            int rowIndex,
            string[] columnReferencePrefixes,
            DataTable table,
            bool useDirectStringCells,
            ref Dictionary<string, int>? sharedStringIndexes,
            bool canCancel,
            CancellationToken ct) {
            string rowReference = InvariantNumberText.Get(rowIndex);
            var row = new Row { RowIndex = (uint)rowIndex };
            for (int offset = 0; offset < table.Columns.Count; offset++) {
                if (canCancel) {
                    ct.ThrowIfCancellationRequested();
                }

                var (cellValue, cellType) = CoerceDataTableAppendValue(table.Columns[offset].ColumnName, useDirectStringCells, ref sharedStringIndexes);
                var cell = CreateTabularAppendCell(columnReferencePrefixes[offset] + rowReference, cellValue, cellType);

                row.Append(cell);
            }

            return row;
        }

        private Row CreateTabularRowSourceHeaderRow(
            int rowIndex,
            string[] columnReferencePrefixes,
            IExcelSheetTabularRowSource source,
            bool useDirectStringCells,
            ref Dictionary<string, int>? sharedStringIndexes,
            bool canCancel,
            CancellationToken ct) {
            string rowReference = InvariantNumberText.Get(rowIndex);
            var row = new Row { RowIndex = (uint)rowIndex };
            for (int offset = 0; offset < source.ColumnCount; offset++) {
                if (canCancel) {
                    ct.ThrowIfCancellationRequested();
                }

                var (cellValue, cellType) = CoerceDataTableAppendValue(source.GetColumnName(offset), useDirectStringCells, ref sharedStringIndexes);
                var cell = CreateTabularAppendCell(columnReferencePrefixes[offset] + rowReference, cellValue, cellType);

                row.Append(cell);
            }

            return row;
        }

        private Row CreateDataTableValueRow(
            int rowIndex,
            string[] columnReferencePrefixes,
            DataRow dataRow,
            TabularAppendColumnKind[] columnKinds,
            uint?[] styleIndexes,
            uint? objectDateTimeStyleIndex,
            uint? objectTimeSpanStyleIndex,
            bool useDirectStringCells,
            ref Dictionary<string, int>? sharedStringIndexes,
            bool canCancel,
            CancellationToken ct) {
            string rowReference = InvariantNumberText.Get(rowIndex);
            int columnCount = dataRow.Table.Columns.Count;
            bool hasObjectValueStyles = objectDateTimeStyleIndex.HasValue || objectTimeSpanStyleIndex.HasValue;
            var row = new Row { RowIndex = (uint)rowIndex };
            for (int offset = 0; offset < columnCount; offset++) {
                if (canCancel) {
                    ct.ThrowIfCancellationRequested();
                }

                object? value = dataRow[offset];
                if (value == DBNull.Value) {
                    value = null;
                }

                var (cellValue, cellType) = CoerceTabularAppendValue(value, columnKinds[offset], useDirectStringCells, ref sharedStringIndexes);
                var cell = CreateTabularAppendCell(columnReferencePrefixes[offset] + rowReference, cellValue, cellType);

                if (styleIndexes[offset] is uint styleIndex) {
                    cell.StyleIndex = styleIndex;
                } else if (hasObjectValueStyles && TryGetObjectDataTableValueStyleIndex(value, objectDateTimeStyleIndex, objectTimeSpanStyleIndex, out uint objectValueStyleIndex)) {
                    cell.StyleIndex = objectValueStyleIndex;
                }

                row.Append(cell);
            }

            return row;
        }

        private Row CreateDataTableValueRow(
            int rowIndex,
            string[] columnReferencePrefixes,
            DataRow dataRow,
            TabularAppendColumnKind[] columnKinds,
            uint?[] styleIndexes,
            uint? objectDateTimeStyleIndex,
            uint? objectTimeSpanStyleIndex,
            bool useDirectStringCells,
            ref Dictionary<string, int>? sharedStringIndexes) {
            string rowReference = InvariantNumberText.Get(rowIndex);
            int columnCount = dataRow.Table.Columns.Count;
            bool hasObjectValueStyles = objectDateTimeStyleIndex.HasValue || objectTimeSpanStyleIndex.HasValue;
            var row = new Row { RowIndex = (uint)rowIndex };
            for (int offset = 0; offset < columnCount; offset++) {
                object? value = dataRow[offset];
                if (value == DBNull.Value) {
                    value = null;
                }

                var (cellValue, cellType) = CoerceTabularAppendValue(value, columnKinds[offset], useDirectStringCells, ref sharedStringIndexes);
                var cell = CreateTabularAppendCell(columnReferencePrefixes[offset] + rowReference, cellValue, cellType);

                if (styleIndexes[offset] is uint styleIndex) {
                    cell.StyleIndex = styleIndex;
                } else if (hasObjectValueStyles && TryGetObjectDataTableValueStyleIndex(value, objectDateTimeStyleIndex, objectTimeSpanStyleIndex, out uint objectValueStyleIndex)) {
                    cell.StyleIndex = objectValueStyleIndex;
                }

                row.Append(cell);
            }

            return row;
        }

        private Row CreateTabularRowSourceValueRow(
            int rowIndex,
            string[] columnReferencePrefixes,
            object?[] values,
            int valueOffset,
            int columnCount,
            TabularAppendColumnKind[] columnKinds,
            uint?[] styleIndexes,
            uint? objectDateTimeStyleIndex,
            uint? objectTimeSpanStyleIndex,
            bool useDirectStringCells,
            ref Dictionary<string, int>? sharedStringIndexes) {
            string rowReference = InvariantNumberText.Get(rowIndex);
            bool hasObjectValueStyles = objectDateTimeStyleIndex.HasValue || objectTimeSpanStyleIndex.HasValue;
            var row = new Row { RowIndex = (uint)rowIndex };
            for (int offset = 0; offset < columnCount; offset++) {
                object? value = values[valueOffset + offset];
                if (value == DBNull.Value) {
                    value = null;
                }

                var (cellValue, cellType) = CoerceTabularAppendValue(value, columnKinds[offset], useDirectStringCells, ref sharedStringIndexes);
                var cell = CreateTabularAppendCell(columnReferencePrefixes[offset] + rowReference, cellValue, cellType);

                if (styleIndexes[offset] is uint styleIndex) {
                    cell.StyleIndex = styleIndex;
                } else if (hasObjectValueStyles && TryGetObjectDataTableValueStyleIndex(value, objectDateTimeStyleIndex, objectTimeSpanStyleIndex, out uint objectValueStyleIndex)) {
                    cell.StyleIndex = objectValueStyleIndex;
                }

                row.Append(cell);
            }

            return row;
        }

        private Row CreateTabularRowSourceValueRow(
            int rowIndex,
            string[] columnReferencePrefixes,
            object?[] values,
            int valueOffset,
            int columnCount,
            TabularAppendColumnKind[] columnKinds,
            uint?[] styleIndexes,
            uint? objectDateTimeStyleIndex,
            uint? objectTimeSpanStyleIndex,
            bool useDirectStringCells,
            ref Dictionary<string, int>? sharedStringIndexes,
            bool canCancel,
            CancellationToken ct) {
            string rowReference = InvariantNumberText.Get(rowIndex);
            bool hasObjectValueStyles = objectDateTimeStyleIndex.HasValue || objectTimeSpanStyleIndex.HasValue;
            var row = new Row { RowIndex = (uint)rowIndex };
            for (int offset = 0; offset < columnCount; offset++) {
                if (canCancel) {
                    ct.ThrowIfCancellationRequested();
                }

                object? value = values[valueOffset + offset];
                if (value == DBNull.Value) {
                    value = null;
                }

                var (cellValue, cellType) = CoerceTabularAppendValue(value, columnKinds[offset], useDirectStringCells, ref sharedStringIndexes);
                var cell = CreateTabularAppendCell(columnReferencePrefixes[offset] + rowReference, cellValue, cellType);

                if (styleIndexes[offset] is uint styleIndex) {
                    cell.StyleIndex = styleIndex;
                } else if (hasObjectValueStyles && TryGetObjectDataTableValueStyleIndex(value, objectDateTimeStyleIndex, objectTimeSpanStyleIndex, out uint objectValueStyleIndex)) {
                    cell.StyleIndex = objectValueStyleIndex;
                }

                row.Append(cell);
            }

            return row;
        }

        private Row CreateTabularRowSourceValueRow(
            int rowIndex,
            string[] columnReferencePrefixes,
            IExcelSheetTabularRowSource source,
            int sourceRowIndex,
            TabularAppendColumnKind[] columnKinds,
            uint?[] styleIndexes,
            uint? objectDateTimeStyleIndex,
            uint? objectTimeSpanStyleIndex,
            bool useDirectStringCells,
            ref Dictionary<string, int>? sharedStringIndexes,
            bool canCancel,
            CancellationToken ct) {
            string rowReference = InvariantNumberText.Get(rowIndex);
            int columnCount = source.ColumnCount;
            bool hasObjectValueStyles = objectDateTimeStyleIndex.HasValue || objectTimeSpanStyleIndex.HasValue;
            var row = new Row { RowIndex = (uint)rowIndex };
            for (int offset = 0; offset < columnCount; offset++) {
                if (canCancel) {
                    ct.ThrowIfCancellationRequested();
                }

                object? value = source.GetValue(sourceRowIndex, offset);
                if (value == DBNull.Value) {
                    value = null;
                }

                var (cellValue, cellType) = CoerceTabularAppendValue(value, columnKinds[offset], useDirectStringCells, ref sharedStringIndexes);
                var cell = CreateTabularAppendCell(columnReferencePrefixes[offset] + rowReference, cellValue, cellType);

                if (styleIndexes[offset] is uint styleIndex) {
                    cell.StyleIndex = styleIndex;
                } else if (hasObjectValueStyles && TryGetObjectDataTableValueStyleIndex(value, objectDateTimeStyleIndex, objectTimeSpanStyleIndex, out uint objectValueStyleIndex)) {
                    cell.StyleIndex = objectValueStyleIndex;
                }

                row.Append(cell);
            }

            return row;
        }

        private static string[] BuildColumnReferencePrefixes(int startColumn, int columnCount) {
            var columnReferencePrefixes = new string[columnCount];
            for (int offset = 0; offset < columnReferencePrefixes.Length; offset++) {
                columnReferencePrefixes[offset] = GetColumnName(startColumn + offset);
            }

            return columnReferencePrefixes;
        }

        private static EnumValue<DocumentFormat.OpenXml.Spreadsheet.CellValues> GetCachedDataTableCellType(DocumentFormat.OpenXml.Spreadsheet.CellValues cellType) {
            if (cellType == DocumentFormat.OpenXml.Spreadsheet.CellValues.String) return DataTableStringCellType;
            if (cellType == DocumentFormat.OpenXml.Spreadsheet.CellValues.SharedString) return DataTableSharedStringCellType;
            if (cellType == DocumentFormat.OpenXml.Spreadsheet.CellValues.Number) return DataTableNumberCellType;
            if (cellType == DocumentFormat.OpenXml.Spreadsheet.CellValues.Boolean) return DataTableBooleanCellType;
            return new EnumValue<DocumentFormat.OpenXml.Spreadsheet.CellValues>(cellType);
        }

        private static Cell CreateTabularAppendCell(
            string cellReference,
            CellValue cellValue,
            DocumentFormat.OpenXml.Spreadsheet.CellValues cellType) {
            var cell = new Cell {
                CellReference = cellReference
            };

            if (cellType != DocumentFormat.OpenXml.Spreadsheet.CellValues.Number) {
                cell.DataType = GetCachedDataTableCellType(cellType);
            }

            cell.AppendChild(cellValue);
            return cell;
        }

        private (CellValue cellValue, DocumentFormat.OpenXml.Spreadsheet.CellValues cellType) CoerceTabularAppendValue(
            object? value,
            TabularAppendColumnKind columnKind,
            bool useDirectStringCells,
            ref Dictionary<string, int>? sharedStringIndexes) {
            switch (columnKind) {
                case TabularAppendColumnKind.String:
                    if (value is string text) {
                        return CoerceTabularAppendStringValue(text, useDirectStringCells, ref sharedStringIndexes);
                    }
                    break;
                case TabularAppendColumnKind.Double:
                    if (value is double doubleValue) return CoerceValueHelper.HandleNumber(doubleValue);
                    break;
                case TabularAppendColumnKind.Float:
                    if (value is float floatValue) return CoerceValueHelper.HandleNumber(floatValue);
                    break;
                case TabularAppendColumnKind.Decimal:
                    if (value is decimal decimalValue) return CoerceValueHelper.HandleDecimal(decimalValue);
                    break;
                case TabularAppendColumnKind.SignedInteger:
                    if (value is int intValue) return CoerceValueHelper.HandleSignedInteger(intValue);
                    if (value is long longValue) return CoerceValueHelper.HandleSignedInteger(longValue);
                    if (value is short shortValue) return CoerceValueHelper.HandleSignedInteger(shortValue);
                    if (value is sbyte sbyteValue) return CoerceValueHelper.HandleSignedInteger(sbyteValue);
                    break;
                case TabularAppendColumnKind.UnsignedInteger:
                    if (value is uint uintValue) return CoerceValueHelper.HandleUnsignedInteger(uintValue);
                    if (value is ulong ulongValue) return CoerceValueHelper.HandleUnsignedInteger(ulongValue);
                    if (value is ushort ushortValue) return CoerceValueHelper.HandleUnsignedInteger(ushortValue);
                    if (value is byte byteValue) return CoerceValueHelper.HandleUnsignedInteger(byteValue);
                    break;
                case TabularAppendColumnKind.Boolean:
                    if (value is bool boolValue) return CoerceValueHelper.HandleBoolean(boolValue);
                    break;
                case TabularAppendColumnKind.DateTime:
                    if (value is DateTime dateTimeValue) return CoerceValueHelper.HandleNumber(dateTimeValue.ToOADate());
                    break;
                case TabularAppendColumnKind.DateTimeOffset:
                    break;
#if NET6_0_OR_GREATER
                case TabularAppendColumnKind.DateOnly:
                    if (value is DateOnly dateOnlyValue) return CoerceValueHelper.HandleNumber(dateOnlyValue.ToDateTime(TimeOnly.MinValue).ToOADate());
                    break;
                case TabularAppendColumnKind.TimeOnly:
                    if (value is TimeOnly timeOnlyValue) return CoerceValueHelper.HandleNumber(timeOnlyValue.ToTimeSpan().TotalDays);
                    break;
#endif
                case TabularAppendColumnKind.TimeSpan:
                    if (value is TimeSpan timeSpanValue) return CoerceValueHelper.HandleNumber(timeSpanValue.TotalDays);
                    break;
            }

            return CoerceDataTableAppendValue(value, useDirectStringCells, ref sharedStringIndexes);
        }

        private (CellValue cellValue, DocumentFormat.OpenXml.Spreadsheet.CellValues cellType) CoerceTabularAppendStringValue(
            string text,
            bool useDirectStringCells,
            ref Dictionary<string, int>? sharedStringIndexes) {
            return useDirectStringCells
                ? (CreatePlainAppendStringValue(text), DocumentFormat.OpenXml.Spreadsheet.CellValues.String)
                : (CreatePlainAppendSharedStringValue(text, ref sharedStringIndexes), DocumentFormat.OpenXml.Spreadsheet.CellValues.SharedString);
        }

        private (CellValue cellValue, DocumentFormat.OpenXml.Spreadsheet.CellValues cellType) CoerceDataTableAppendValue(
            object? value,
            bool useDirectStringCells,
            ref Dictionary<string, int>? sharedStringIndexes) {
            var indexes = sharedStringIndexes;
            CellValue HandleString(string text) {
                return useDirectStringCells
                    ? CreatePlainAppendStringValue(text)
                    : CreatePlainAppendSharedStringValue(text, ref indexes);
            }

            var (cellValue, cellType) = CoerceValueHelper.Coerce(
                value,
                HandleString,
                _excelDocument.DateTimeOffsetWriteStrategy);
            sharedStringIndexes = indexes;

            if (useDirectStringCells && cellType == DocumentFormat.OpenXml.Spreadsheet.CellValues.SharedString) {
                cellType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String;
            }

            return (cellValue, cellType);
        }

        private static string?[] BuildDataTableNumberFormats(DataTable table) {
            var formats = new string?[table.Columns.Count];
            for (int i = 0; i < table.Columns.Count; i++) {
                formats[i] = GetDataTableNumberFormat(table.Columns[i].DataType, value: null);
            }

            return formats;
        }

        private static string?[] BuildTabularRowSourceNumberFormats(IExcelSheetTabularRowSource source) {
            var formats = new string?[source.ColumnCount];
            for (int i = 0; i < source.ColumnCount; i++) {
                formats[i] = GetDataTableNumberFormat(source.GetColumnType(i), value: null);
            }

            return formats;
        }

        private static TabularAppendColumnKind[] BuildDataTableAppendColumnKinds(DataTable table) {
            var kinds = new TabularAppendColumnKind[table.Columns.Count];
            for (int i = 0; i < kinds.Length; i++) {
                kinds[i] = GetTabularAppendColumnKind(table.Columns[i].DataType);
            }

            return kinds;
        }

        private static TabularAppendColumnKind[] BuildTabularAppendColumnKinds(IExcelSheetTabularRowSource source) {
            var kinds = new TabularAppendColumnKind[source.ColumnCount];
            for (int i = 0; i < kinds.Length; i++) {
                kinds[i] = GetTabularAppendColumnKind(source.GetColumnType(i));
            }

            return kinds;
        }

        private static TabularAppendColumnKind GetTabularAppendColumnKind(Type type) {
            if (type == typeof(string)) return TabularAppendColumnKind.String;
            if (type == typeof(double)) return TabularAppendColumnKind.Double;
            if (type == typeof(float)) return TabularAppendColumnKind.Float;
            if (type == typeof(decimal)) return TabularAppendColumnKind.Decimal;
            if (type == typeof(int) || type == typeof(long) || type == typeof(short) || type == typeof(sbyte)) return TabularAppendColumnKind.SignedInteger;
            if (type == typeof(uint) || type == typeof(ulong) || type == typeof(ushort) || type == typeof(byte)) return TabularAppendColumnKind.UnsignedInteger;
            if (type == typeof(bool)) return TabularAppendColumnKind.Boolean;
            if (type == typeof(DateTime)) return TabularAppendColumnKind.DateTime;
            if (type == typeof(DateTimeOffset)) return TabularAppendColumnKind.DateTimeOffset;
#if NET6_0_OR_GREATER
            if (type == typeof(DateOnly)) return TabularAppendColumnKind.DateOnly;
            if (type == typeof(TimeOnly)) return TabularAppendColumnKind.TimeOnly;
#endif
            if (type == typeof(TimeSpan)) return TabularAppendColumnKind.TimeSpan;
            return TabularAppendColumnKind.General;
        }

        private static string? GetDataTableNumberFormat(Type type, object? value) {
            if (type == typeof(DateTime) || type == typeof(DateTimeOffset) || value is DateTime || value is DateTimeOffset) {
                return DataTableDateTimeNumberFormat;
            }

            if (type == typeof(TimeSpan) || value is TimeSpan) {
                return DataTableTimeSpanNumberFormat;
            }

#if NET6_0_OR_GREATER
            if (type == typeof(DateOnly) || value is DateOnly) {
                return DataTableDateTimeNumberFormat;
            }

            if (type == typeof(TimeOnly) || value is TimeOnly) {
                return DataTableTimeSpanNumberFormat;
            }
#endif

            return null;
        }

        private static bool TryGetObjectDataTableValueStyleIndex(object? value, uint? dateTimeStyleIndex, uint? timeSpanStyleIndex, out uint styleIndex) {
            styleIndex = 0U;
            if (value is DateTime || value is DateTimeOffset
#if NET6_0_OR_GREATER
                || value is DateOnly
#endif
                ) {
                if (dateTimeStyleIndex.HasValue) {
                    styleIndex = dateTimeStyleIndex.Value;
                    return true;
                }

                return false;
            }

            if (value is TimeSpan
#if NET6_0_OR_GREATER
                || value is TimeOnly
#endif
                ) {
                if (timeSpanStyleIndex.HasValue) {
                    styleIndex = timeSpanStyleIndex.Value;
                    return true;
                }
            }

            return false;
        }
    }
}
