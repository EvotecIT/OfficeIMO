#if NET6_0_OR_GREATER
using System.Buffers;
#endif
using System.Data;
using System.Globalization;
using System.Text;
using System.Text.Json;
using System.Threading;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private static readonly char[] CsvSpecialCharacters = { ',', '"', '\r', '\n' };

        /// <summary>
        /// Reads the worksheet used range as a <see cref="DataTable"/>.
        /// </summary>
        public DataTable ToDataTable(bool headersInFirstRow = true, ExcelReadOptions? options = null, ExecutionMode? mode = null, CancellationToken ct = default) {
            return ToDataTable(GetUsedRangeA1(), headersInFirstRow, options, mode, ct);
        }

        /// <summary>
        /// Reads an A1 range as a <see cref="DataTable"/>.
        /// </summary>
        public DataTable ToDataTable(string a1Range, bool headersInFirstRow = true, ExcelReadOptions? options = null, ExecutionMode? mode = null, CancellationToken ct = default) {
            if (string.IsNullOrWhiteSpace(a1Range)) throw new ArgumentNullException(nameof(a1Range));
            using var reader = _excelDocument.CreateReader(options);
            return reader.GetSheet(Name).ReadRangeAsDataTable(a1Range, headersInFirstRow, mode, ct);
        }

        /// <summary>
        /// Reads an Excel table as a <see cref="DataTable"/>.
        /// </summary>
        public DataTable TableToDataTable(string tableName, bool? headersInFirstRow = null, ExcelReadOptions? options = null, ExecutionMode? mode = null, CancellationToken ct = default) {
            if (string.IsNullOrWhiteSpace(tableName)) throw new ArgumentNullException(nameof(tableName));
            using var reader = _excelDocument.CreateReader(options);
            return reader.ReadTableAsDataTable(tableName, headersInFirstRow, mode, ct);
        }

        /// <summary>
        /// Reads an A1 range and returns CSV text.
        /// </summary>
        public string ToCsv(string a1Range, bool headersInFirstRow = true, ExcelReadOptions? options = null, ExecutionMode? mode = null, CancellationToken ct = default) {
            return DataTableToCsv(ToDataTable(a1Range, headersInFirstRow, options, mode, ct), includeHeaders: headersInFirstRow);
        }

        /// <summary>
        /// Reads the worksheet used range and returns CSV text.
        /// </summary>
        public string ToCsv(bool headersInFirstRow = true, ExcelReadOptions? options = null, ExecutionMode? mode = null, CancellationToken ct = default) {
            return ToCsv(GetUsedRangeA1(), headersInFirstRow, options, mode, ct);
        }

        /// <summary>
        /// Reads an Excel table and returns CSV text.
        /// </summary>
        public string TableToCsv(string tableName, bool? headersInFirstRow = null, ExcelReadOptions? options = null, ExecutionMode? mode = null, CancellationToken ct = default) {
            bool includeHeaders = headersInFirstRow ?? true;
            return DataTableToCsv(TableToDataTable(tableName, headersInFirstRow, options, mode, ct), includeHeaders);
        }

        /// <summary>
        /// Reads an A1 range and returns JSON as an array of objects.
        /// </summary>
        public string ToJson(string a1Range, bool headersInFirstRow = true, ExcelReadOptions? options = null, ExecutionMode? mode = null, JsonSerializerOptions? jsonOptions = null, CancellationToken ct = default) {
            return DataTableToJson(ToDataTable(a1Range, headersInFirstRow, options, mode, ct), jsonOptions);
        }

        /// <summary>
        /// Reads the worksheet used range and returns JSON as an array of objects.
        /// </summary>
        public string ToJson(bool headersInFirstRow = true, ExcelReadOptions? options = null, ExecutionMode? mode = null, JsonSerializerOptions? jsonOptions = null, CancellationToken ct = default) {
            return ToJson(GetUsedRangeA1(), headersInFirstRow, options, mode, jsonOptions, ct);
        }

        /// <summary>
        /// Reads an Excel table and returns JSON as an array of objects.
        /// </summary>
        public string TableToJson(string tableName, bool? headersInFirstRow = null, ExcelReadOptions? options = null, ExecutionMode? mode = null, JsonSerializerOptions? jsonOptions = null, CancellationToken ct = default) {
            return DataTableToJson(TableToDataTable(tableName, headersInFirstRow, options, mode, ct), jsonOptions);
        }

        /// <summary>
        /// Inserts CSV text into the worksheet and returns the inserted range.
        /// </summary>
        public string FromCsv(string csv, int startRow = 1, int startColumn = 1, bool firstRowIsHeader = true, bool includeHeaders = true, ExecutionMode? mode = null, CancellationToken ct = default) {
            DataTable table = CsvToDataTable(csv, firstRowIsHeader);
            InsertDataTable(table, startRow, startColumn, includeHeaders, mode, ct);
            return BuildInsertedRange(table, startRow, startColumn, includeHeaders);
        }

        /// <summary>
        /// Inserts JSON array data into the worksheet and returns the inserted range.
        /// </summary>
        public string FromJson(string json, int startRow = 1, int startColumn = 1, bool includeHeaders = true, ExecutionMode? mode = null, CancellationToken ct = default) {
            DataTable table = JsonToDataTable(json);
            InsertDataTable(table, startRow, startColumn, includeHeaders, mode, ct);
            return BuildInsertedRange(table, startRow, startColumn, includeHeaders);
        }

        private static string DataTableToCsv(DataTable table, bool includeHeaders) {
            if (table == null) throw new ArgumentNullException(nameof(table));
            var builder = new StringBuilder(EstimateCsvCapacity(table, includeHeaders));
            int columnCount = table.Columns.Count;
            if (includeHeaders) {
                for (int column = 0; column < columnCount; column++) {
                    if (column > 0) builder.Append(',');
                    AppendCsvField(builder, table.Columns[column].ColumnName);
                }
                builder.AppendLine();
            }

            foreach (DataRow row in table.Rows) {
                for (int column = 0; column < columnCount; column++) {
                    if (column > 0) builder.Append(',');
                    object? value = row.IsNull(column) ? null : row[column];
                    AppendCsvField(builder, value);
                }
                builder.AppendLine();
            }

            return builder.ToString();
        }

        private static string DataTableToJson(DataTable table, JsonSerializerOptions? options) {
            if (table == null) throw new ArgumentNullException(nameof(table));
            if (options == null) {
                return DataTableToJsonStreaming(table);
            }

            int columnCount = table.Columns.Count;
            string[] columnNames = new string[columnCount];
            for (int column = 0; column < columnCount; column++) {
                columnNames[column] = table.Columns[column].ColumnName;
            }

            var rows = new List<Dictionary<string, object?>>(table.Rows.Count);
            foreach (DataRow row in table.Rows) {
                var item = new Dictionary<string, object?>(columnCount, StringComparer.OrdinalIgnoreCase);
                for (int column = 0; column < columnCount; column++) {
                    object? value = row.IsNull(column) ? null : row[column];
                    item[columnNames[column]] = value;
                }

                rows.Add(item);
            }

            return JsonSerializer.Serialize(rows, options);
        }

        private static string DataTableToJsonStreaming(DataTable table) {
            JsonEncodedText[] propertyNames = CreateJsonPropertyNames(table.Columns);
            int estimatedCapacity = EstimateJsonCapacity(table, propertyNames);
#if NET6_0_OR_GREATER
            var buffer = new ArrayBufferWriter<byte>(estimatedCapacity);
            using (var writer = new Utf8JsonWriter(buffer)) {
                WriteDataTableJson(table, propertyNames, writer);
            }

            return Encoding.UTF8.GetString(buffer.WrittenSpan);
#else
            using var stream = new MemoryStream(estimatedCapacity);
            using (var writer = new Utf8JsonWriter(stream)) {
                WriteDataTableJson(table, propertyNames, writer);
            }

            byte[] jsonBytes = stream.ToArray();
            return Encoding.UTF8.GetString(jsonBytes, 0, jsonBytes.Length);
#endif
        }

        private static void WriteDataTableJson(DataTable table, JsonEncodedText[] propertyNames, Utf8JsonWriter writer) {
            writer.WriteStartArray();
            foreach (DataRow row in table.Rows) {
                writer.WriteStartObject();
                for (int i = 0; i < propertyNames.Length; i++) {
                    writer.WritePropertyName(propertyNames[i]);
                    object? value = row.IsNull(i) ? null : row[i];
                    if (value == null) {
                        writer.WriteNullValue();
                    } else {
                        WriteJsonValue(writer, value);
                    }
                }

                writer.WriteEndObject();
            }

            writer.WriteEndArray();
        }

        private static void WriteJsonValue(Utf8JsonWriter writer, object value) {
            switch (value) {
                case string stringValue:
                    writer.WriteStringValue(stringValue);
                    return;
                case bool boolValue:
                    writer.WriteBooleanValue(boolValue);
                    return;
                case int intValue:
                    writer.WriteNumberValue(intValue);
                    return;
                case long longValue:
                    writer.WriteNumberValue(longValue);
                    return;
                case double doubleValue:
                    writer.WriteNumberValue(doubleValue);
                    return;
                case decimal decimalValue:
                    writer.WriteNumberValue(decimalValue);
                    return;
                case DateTime dateTime:
                    writer.WriteStringValue(dateTime);
                    return;
                case DateTimeOffset dateTimeOffset:
                    writer.WriteStringValue(dateTimeOffset);
                    return;
                case Guid guid:
                    writer.WriteStringValue(guid);
                    return;
                case float floatValue:
                    writer.WriteNumberValue(floatValue);
                    return;
                case short shortValue:
                    writer.WriteNumberValue(shortValue);
                    return;
                case byte byteValue:
                    writer.WriteNumberValue(byteValue);
                    return;
                case sbyte sbyteValue:
                    writer.WriteNumberValue(sbyteValue);
                    return;
                case ushort ushortValue:
                    writer.WriteNumberValue(ushortValue);
                    return;
                case uint uintValue:
                    writer.WriteNumberValue(uintValue);
                    return;
                case ulong ulongValue:
                    writer.WriteNumberValue(ulongValue);
                    return;
                default:
                    JsonSerializer.Serialize(writer, value, value.GetType());
                    return;
            }
        }

        private static JsonEncodedText[] CreateJsonPropertyNames(DataColumnCollection columns) {
            var propertyNames = new JsonEncodedText[columns.Count];
            for (int i = 0; i < propertyNames.Length; i++) {
                propertyNames[i] = JsonEncodedText.Encode(columns[i].ColumnName);
            }

            return propertyNames;
        }

        private static int EstimateJsonCapacity(DataTable table, JsonEncodedText[] propertyNames) {
            const int MaxInitialJsonBufferBytes = 16 * 1024 * 1024;
            int rowCount = table.Rows.Count;
            int columnCount = propertyNames.Length;
            if (rowCount <= 0 || columnCount <= 0) {
                return 256;
            }

            long perRow = 2;
            for (int i = 0; i < propertyNames.Length; i++) {
                perRow += propertyNames[i].EncodedUtf8Bytes.Length + 8L;
            }

            perRow += (long)columnCount * 16L;
            long estimated = 2L + (long)rowCount * perRow;
            if (estimated < 256L) {
                return 256;
            }

            return estimated > MaxInitialJsonBufferBytes
                ? MaxInitialJsonBufferBytes
                : (int)estimated;
        }

        private static DataTable CsvToDataTable(string csv, bool firstRowIsHeader) {
            if (csv == null) throw new ArgumentNullException(nameof(csv));
            if (csv.IndexOf('"') < 0) {
                return UnquotedCsvToDataTable(csv, firstRowIsHeader);
            }

            List<List<string?>> records = ParseCsv(csv, out int columnCount);
            var table = new DataTable { Locale = CultureInfo.InvariantCulture };
            if (records.Count == 0) {
                return table;
            }

            int firstDataRow = 0;
            if (firstRowIsHeader) {
                var headers = ExcelHeaderNameHelper.BuildUniqueHeaders(columnCount, c => c < records[0].Count ? records[0][c] : null, true);
                foreach (string header in headers) {
                    table.Columns.Add(header, typeof(string));
                }
                firstDataRow = 1;
            } else {
                for (int column = 0; column < columnCount; column++) {
                    table.Columns.Add($"Column{column + 1}", typeof(string));
                }
            }

            table.MinimumCapacity = Math.Max(0, records.Count - firstDataRow);
            table.BeginLoadData();
            try {
                for (int rowIndex = firstDataRow; rowIndex < records.Count; rowIndex++) {
                    var record = records[rowIndex];
                    DataRow row = table.NewRow();
                    int columnLimit = Math.Min(record.Count, table.Columns.Count);
                    for (int column = 0; column < columnLimit; column++) {
                        string? value = record[column];
                        if (value != null) {
                            row[column] = value;
                        }
                    }

                    table.Rows.Add(row);
                }
            } finally {
                table.EndLoadData();
            }

            return table;
        }

        private static DataTable UnquotedCsvToDataTable(string csv, bool firstRowIsHeader) {
            CountUnquotedCsvShape(csv, out int recordCount, out int columnCount);
            var table = new DataTable { Locale = CultureInfo.InvariantCulture };
            if (recordCount == 0) {
                return table;
            }

            int firstDataRow = 0;
            if (firstRowIsHeader) {
                List<string?> firstRecord = ParseFirstUnquotedCsvRecord(csv);
                var headers = ExcelHeaderNameHelper.BuildUniqueHeaders(columnCount, c => c < firstRecord.Count ? firstRecord[c] : null, true);
                foreach (string header in headers) {
                    table.Columns.Add(header, typeof(string));
                }

                firstDataRow = 1;
            } else {
                for (int column = 0; column < columnCount; column++) {
                    table.Columns.Add($"Column{column + 1}", typeof(string));
                }
            }

            table.MinimumCapacity = Math.Max(0, recordCount - firstDataRow);
            table.BeginLoadData();
            try {
                AddUnquotedCsvRows(csv, firstDataRow, table);
            } finally {
                table.EndLoadData();
            }

            return table;
        }

        private static DataTable JsonToDataTable(string json) {
            if (string.IsNullOrWhiteSpace(json)) throw new ArgumentNullException(nameof(json));

            using JsonDocument document = JsonDocument.Parse(json);
            if (document.RootElement.ValueKind != JsonValueKind.Array) {
                throw new ArgumentException("JSON input must be an array of objects.", nameof(json));
            }

            var table = new DataTable { Locale = CultureInfo.InvariantCulture };
            int expectedRowCount = document.RootElement.GetArrayLength();
            table.MinimumCapacity = expectedRowCount;
            var rows = new List<JsonTableRow>(expectedRowCount);
            var columnIndexes = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            foreach (JsonElement element in document.RootElement.EnumerateArray()) {
                if (element.ValueKind != JsonValueKind.Object) {
                    throw new ArgumentException("JSON input must be an array of objects.", nameof(json));
                }

                rows.Add(CreateJsonTableRow(element, table, columnIndexes));
            }

            table.BeginLoadData();
            try {
                foreach (JsonTableRow source in rows) {
                    DataRow row = table.NewRow();
                    if (source.Values != null) {
                        object?[] values = source.Values;
                        for (int column = 0; column < values.Length; column++) {
                            object? value = values[column];
                            if (value != null) {
                                row[column] = value;
                            }
                        }
                    } else if (source.SparseValues != null) {
                        JsonCellValue[] values = source.SparseValues;
                        for (int i = 0; i < values.Length; i++) {
                            JsonCellValue cell = values[i];
                            row[cell.ColumnIndex] = cell.Value ?? DBNull.Value;
                        }
                    }

                    table.Rows.Add(row);
                }
            } finally {
                table.EndLoadData();
            }

            return table;
        }

        private static JsonTableRow CreateJsonTableRow(JsonElement element, DataTable table, Dictionary<string, int> columnIndexes) {
            int existingColumnCount = table.Columns.Count;
            int propertyCount = CountJsonProperties(element);
            if (UseSparseJsonRow(propertyCount, existingColumnCount)) {
                var values = new JsonCellValue[propertyCount];
                int valueIndex = 0;
                foreach (JsonProperty property in element.EnumerateObject()) {
                    int columnIndex = GetOrAddJsonColumn(table, columnIndexes, property.Name);
                    values[valueIndex++] = new JsonCellValue(columnIndex, JsonElementToValue(property.Value));
                }

                return new JsonTableRow(values);
            }

            object?[] row = new object?[existingColumnCount];
            foreach (JsonProperty property in element.EnumerateObject()) {
                int columnIndex = GetOrAddJsonColumn(table, columnIndexes, property.Name);
                EnsureJsonRowCapacity(ref row, columnIndex + 1);
                row[columnIndex] = JsonElementToValue(property.Value);
            }

            return new JsonTableRow(row);
        }

        private static int GetOrAddJsonColumn(DataTable table, Dictionary<string, int> columnIndexes, string columnName) {
            if (!columnIndexes.TryGetValue(columnName, out int columnIndex)) {
                columnIndex = table.Columns.Count;
                table.Columns.Add(columnName, typeof(object));
                columnIndexes.Add(columnName, columnIndex);
            }

            return columnIndex;
        }

        private static bool UseSparseJsonRow(int propertyCount, int existingColumnCount) {
            return propertyCount > 0 && existingColumnCount >= 32 && propertyCount * 4 <= existingColumnCount;
        }

        private static int CountJsonProperties(JsonElement element) {
            int count = 0;
            foreach (JsonProperty _ in element.EnumerateObject()) {
                count++;
            }

            return count;
        }

        private readonly struct JsonTableRow {
            internal JsonTableRow(object?[] values) {
                Values = values;
                SparseValues = null;
            }

            internal JsonTableRow(JsonCellValue[] sparseValues) {
                Values = null;
                SparseValues = sparseValues;
            }

            internal object?[]? Values { get; }

            internal JsonCellValue[]? SparseValues { get; }
        }

        private readonly struct JsonCellValue {
            internal JsonCellValue(int columnIndex, object? value) {
                ColumnIndex = columnIndex;
                Value = value;
            }

            internal int ColumnIndex { get; }

            internal object? Value { get; }
        }

        private static void EnsureJsonRowCapacity(ref object?[] row, int requiredLength) {
            if (row.Length >= requiredLength) {
                return;
            }

            int newLength = row.Length == 0 ? 4 : row.Length * 2;
            if (newLength < requiredLength) {
                newLength = requiredLength;
            }

            Array.Resize(ref row, newLength);
        }

        private static object? JsonElementToValue(JsonElement element) {
            switch (element.ValueKind) {
                case JsonValueKind.Null:
                case JsonValueKind.Undefined:
                    return null;
                case JsonValueKind.True:
                    return true;
                case JsonValueKind.False:
                    return false;
                case JsonValueKind.Number:
                    if (element.TryGetInt64(out long integer)) return integer;
                    if (element.TryGetDecimal(out decimal dec)) return dec;
                    return element.GetDouble();
                case JsonValueKind.String:
                    if (element.TryGetDateTime(out DateTime dateTime)) return dateTime;
                    return element.GetString();
                default:
                    return element.GetRawText();
            }
        }

        private static List<List<string?>> ParseCsv(string csv, out int maxColumnCount) {
            var records = new List<List<string?>>();
            var record = new List<string?>();
            var field = new StringBuilder();
            bool inQuotes = false;
            bool quoted = false;
            maxColumnCount = 0;

            for (int i = 0; i < csv.Length; i++) {
                char ch = csv[i];
                if (inQuotes) {
                    if (ch == '"') {
                        if (i + 1 < csv.Length && csv[i + 1] == '"') {
                            field.Append('"');
                            i++;
                        } else {
                            inQuotes = false;
                        }
                    } else {
                        field.Append(ch);
                    }
                    continue;
                }

                if (ch == '"' && field.Length == 0) {
                    inQuotes = true;
                    quoted = true;
                    continue;
                }

                if (ch == ',') {
                    record.Add(FieldValue(field, quoted));
                    field.Clear();
                    quoted = false;
                    continue;
                }

                if (ch == '\r' || ch == '\n') {
                    if (ch == '\r' && i + 1 < csv.Length && csv[i + 1] == '\n') {
                        i++;
                    }
                    record.Add(FieldValue(field, quoted));
                    field.Clear();
                    quoted = false;
                    AddCsvRecord(records, record, ref maxColumnCount);
                    record = new List<string?>();
                    continue;
                }

                field.Append(ch);
            }

            if (field.Length > 0 || quoted || record.Count > 0) {
                record.Add(FieldValue(field, quoted));
                AddCsvRecord(records, record, ref maxColumnCount);
            }

            return records;
        }

        private static void CountUnquotedCsvShape(string csv, out int recordCount, out int maxColumnCount) {
            recordCount = 0;
            maxColumnCount = 0;
            int columnCount = 0;
            int fieldStart = 0;

            for (int i = 0; i < csv.Length; i++) {
                char ch = csv[i];
                if (ch == ',') {
                    columnCount++;
                    fieldStart = i + 1;
                    continue;
                }

                if (ch == '\r' || ch == '\n') {
                    columnCount++;
                    if (columnCount > maxColumnCount) {
                        maxColumnCount = columnCount;
                    }

                    recordCount++;
                    columnCount = 0;
                    if (ch == '\r' && i + 1 < csv.Length && csv[i + 1] == '\n') {
                        i++;
                    }

                    fieldStart = i + 1;
                }
            }

            if (fieldStart < csv.Length || columnCount > 0) {
                columnCount++;
                if (columnCount > maxColumnCount) {
                    maxColumnCount = columnCount;
                }

                recordCount++;
            }
        }

        private static List<string?> ParseFirstUnquotedCsvRecord(string csv) {
            var record = new List<string?>();
            int fieldStart = 0;

            for (int i = 0; i < csv.Length; i++) {
                char ch = csv[i];
                if (ch == ',') {
                    record.Add(UnquotedFieldValue(csv, fieldStart, i - fieldStart));
                    fieldStart = i + 1;
                    continue;
                }

                if (ch == '\r' || ch == '\n') {
                    record.Add(UnquotedFieldValue(csv, fieldStart, i - fieldStart));
                    return record;
                }
            }

            if (fieldStart < csv.Length || record.Count > 0) {
                record.Add(UnquotedFieldValue(csv, fieldStart, csv.Length - fieldStart));
            }

            return record;
        }

        private static void AddUnquotedCsvRows(string csv, int firstDataRow, DataTable table) {
            int recordIndex = 0;
            int columnIndex = 0;
            int fieldStart = 0;
            DataRow? row = null;

            for (int i = 0; i < csv.Length; i++) {
                char ch = csv[i];
                if (ch == ',') {
                    AddField(i - fieldStart);
                    fieldStart = i + 1;
                    continue;
                }

                if (ch == '\r' || ch == '\n') {
                    AddField(i - fieldStart);
                    AddRecord();
                    if (ch == '\r' && i + 1 < csv.Length && csv[i + 1] == '\n') {
                        i++;
                    }

                    fieldStart = i + 1;
                }
            }

            if (fieldStart < csv.Length || columnIndex > 0) {
                AddField(csv.Length - fieldStart);
                AddRecord();
            }

            void AddField(int length) {
                if (recordIndex >= firstDataRow && columnIndex < table.Columns.Count) {
                    row ??= table.NewRow();
                    string? value = UnquotedFieldValue(csv, fieldStart, length);
                    if (value != null) {
                        row[columnIndex] = value;
                    }
                }

                columnIndex++;
            }

            void AddRecord() {
                if (recordIndex >= firstDataRow) {
                    row ??= table.NewRow();
                    table.Rows.Add(row);
                    row = null;
                }

                recordIndex++;
                columnIndex = 0;
            }
        }

        private static void AddCsvRecord(List<List<string?>> records, List<string?> record, ref int maxColumnCount) {
            if (record.Count > maxColumnCount) {
                maxColumnCount = record.Count;
            }

            records.Add(record);
        }

        private static string? UnquotedFieldValue(string csv, int start, int length) {
            return length == 0 ? null : csv.Substring(start, length);
        }

        private static string? FieldValue(StringBuilder field, bool quoted) {
            string value = field.ToString();
            return !quoted && value.Length == 0 ? null : value;
        }

        private static void AppendCsvField(StringBuilder builder, object? value) {
            if (value == null || value == DBNull.Value) {
                return;
            }

            string text = value as string ?? Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty;
            if (text.Length == 0) {
                return;
            }

            if (text.IndexOfAny(CsvSpecialCharacters) < 0) {
                builder.Append(text);
                return;
            }

            builder.Append('"');
            if (text.IndexOf('"') < 0) {
                builder.Append(text);
            } else {
                foreach (char ch in text) {
                    if (ch == '"') {
                        builder.Append("\"\"");
                    } else {
                        builder.Append(ch);
                    }
                }
            }

            builder.Append('"');
        }

        private static int EstimateCsvCapacity(DataTable table, bool includeHeaders) {
            int rowCount = table.Rows.Count + (includeHeaders ? 1 : 0);
            if (rowCount <= 0 || table.Columns.Count == 0) {
                return 256;
            }

            long estimated = ((long)rowCount * table.Columns.Count * 12) + ((long)rowCount * Environment.NewLine.Length);
            if (estimated < 256) {
                return 256;
            }

            return estimated > int.MaxValue ? int.MaxValue : (int)estimated;
        }

        private static string BuildInsertedRange(DataTable table, int startRow, int startColumn, bool includeHeaders) {
            int rowCount = table.Rows.Count + (includeHeaders ? 1 : 0);
            if (table.Columns.Count == 0 || rowCount == 0) {
                return string.Empty;
            }

            return A1.CellReference(startRow, startColumn) + ":" + A1.CellReference(startRow + rowCount - 1, startColumn + table.Columns.Count - 1);
        }
    }
}
