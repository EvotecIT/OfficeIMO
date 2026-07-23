#if NET6_0_OR_GREATER
using System.Buffers;
#endif
using System.Data;
using System.Globalization;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
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
        /// Inserts JSON array data into the worksheet and returns the inserted range.
        /// </summary>
        public string FromJson(string json, int startRow = 1, int startColumn = 1, bool includeHeaders = true, ExecutionMode? mode = null, CancellationToken ct = default) {
            DataTable table = JsonToDataTable(json);
            InsertDataTable(table, startRow, startColumn, includeHeaders, mode, ct);
            return BuildInsertedRange(table, startRow, startColumn, includeHeaders);
        }

        private static string DataTableToJson(DataTable table, JsonSerializerOptions? options) {
            if (table == null) throw new ArgumentNullException(nameof(table));
            if (options == null) {
                return DataTableToJsonStreaming(table);
            }

            return DataTableToJsonWithOptions(table, options);
        }

        [System.Diagnostics.CodeAnalysis.UnconditionalSuppressMessage("Trimming", "IL2026",
            Justification = "The JsonSerializerOptions compatibility path intentionally honors caller-provided row and cell converters and metadata. The default NativeAOT path uses the typed streaming writer.")]
        [System.Diagnostics.CodeAnalysis.UnconditionalSuppressMessage("AOT", "IL3050",
            Justification = "The JsonSerializerOptions compatibility path intentionally honors caller-provided row and cell converters and metadata. The default NativeAOT path uses the typed streaming writer.")]
        private static string DataTableToJsonWithOptions(DataTable table, JsonSerializerOptions options) {
            EnsureStreamingJsonOptionsSupported(options);
            JsonEncodedText[] propertyNames = CreateJsonPropertyNames(table.Columns, options);
            int estimatedCapacity = EstimateJsonCapacity(table, propertyNames);
            var writerOptions = new JsonWriterOptions {
                Encoder = options.Encoder,
                Indented = options.WriteIndented
            };
#if NET6_0_OR_GREATER
            var buffer = new ArrayBufferWriter<byte>(estimatedCapacity);
            using (var writer = new Utf8JsonWriter(buffer, writerOptions)) {
                WriteDataTableJson(table, propertyNames, writer, options);
            }

            return Encoding.UTF8.GetString(buffer.WrittenSpan);
#else
            using var stream = new MemoryStream(estimatedCapacity);
            using (var writer = new Utf8JsonWriter(stream, writerOptions)) {
                WriteDataTableJson(table, propertyNames, writer, options);
            }

            byte[] jsonBytes = stream.ToArray();
            return Encoding.UTF8.GetString(jsonBytes, 0, jsonBytes.Length);
#endif
        }

        private static void EnsureStreamingJsonOptionsSupported(JsonSerializerOptions options) {
            if (options.ReferenceHandler != null) {
                throw new NotSupportedException(
                    "ReferenceHandler is not supported for bounded DataTable JSON export because reference metadata requires whole-table materialization.");
            }

            Type rootType = typeof(List<Dictionary<string, object?>>);
            foreach (JsonConverter converter in options.Converters) {
                if (converter.CanConvert(rootType)) {
                    throw new NotSupportedException(
                        "Converters for the complete DataTable row collection are not supported by bounded JSON export. Register row or value converters instead.");
                }
            }
        }

        private static string DataTableToJsonStreaming(DataTable table) {
            JsonEncodedText[] propertyNames = CreateJsonPropertyNames(table.Columns, options: null);
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

        private static void WriteDataTableJson(
            DataTable table,
            JsonEncodedText[] propertyNames,
            Utf8JsonWriter writer,
            JsonSerializerOptions options) {
            writer.WriteStartArray();
            foreach (DataRow row in table.Rows) {
                var item = new Dictionary<string, object?>(propertyNames.Length, StringComparer.OrdinalIgnoreCase);
                for (int i = 0; i < propertyNames.Length; i++) {
                    item[table.Columns[i].ColumnName] = row.IsNull(i) ? null : row[i];
                }

                JsonSerializer.Serialize(writer, item, options);
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
                case TimeSpan timeSpan:
                    writer.WriteStringValue(timeSpan.ToString("c", CultureInfo.InvariantCulture));
                    return;
                case char character:
                    writer.WriteStringValue(character.ToString());
                    return;
                case byte[] bytes:
                    writer.WriteBase64StringValue(bytes);
                    return;
                case JsonElement element:
                    element.WriteTo(writer);
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
#if NET6_0_OR_GREATER
                case DateOnly dateOnly:
                    writer.WriteStringValue(dateOnly.ToString("O", CultureInfo.InvariantCulture));
                    return;
                case TimeOnly timeOnly:
                    writer.WriteStringValue(timeOnly.ToString("O", CultureInfo.InvariantCulture));
                    return;
#endif
                default:
                    writer.WriteStringValue(Convert.ToString(value, CultureInfo.InvariantCulture));
                    return;
            }
        }

        private static JsonEncodedText[] CreateJsonPropertyNames(DataColumnCollection columns, JsonSerializerOptions? options) {
            var propertyNames = new JsonEncodedText[columns.Count];
            for (int i = 0; i < propertyNames.Length; i++) {
                string columnName = columns[i].ColumnName;
                if (options?.DictionaryKeyPolicy != null) {
                    columnName = options.DictionaryKeyPolicy.ConvertName(columnName);
                }

                propertyNames[i] = JsonEncodedText.Encode(columnName, options?.Encoder);
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

        private static string BuildInsertedRange(DataTable table, int startRow, int startColumn, bool includeHeaders) {
            int rowCount = table.Rows.Count + (includeHeaders ? 1 : 0);
            if (table.Columns.Count == 0 || rowCount == 0) {
                return string.Empty;
            }

            return A1.CellReference(startRow, startColumn) + ":" + A1.CellReference(startRow + rowCount - 1, startColumn + table.Columns.Count - 1);
        }
    }
}
