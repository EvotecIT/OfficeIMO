using System.Data;
using System.Globalization;
using System.Text;
using System.Text.Json;
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
            var builder = new StringBuilder();
            if (includeHeaders) {
                for (int column = 0; column < table.Columns.Count; column++) {
                    if (column > 0) builder.Append(',');
                    AppendCsvField(builder, table.Columns[column].ColumnName);
                }
                builder.AppendLine();
            }

            foreach (DataRow row in table.Rows) {
                for (int column = 0; column < table.Columns.Count; column++) {
                    if (column > 0) builder.Append(',');
                    AppendCsvField(builder, row.IsNull(column) ? null : row[column]);
                }
                builder.AppendLine();
            }

            return builder.ToString();
        }

        private static string DataTableToJson(DataTable table, JsonSerializerOptions? options) {
            if (table == null) throw new ArgumentNullException(nameof(table));
            var rows = new List<Dictionary<string, object?>>(table.Rows.Count);
            foreach (DataRow row in table.Rows) {
                var item = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
                foreach (DataColumn column in table.Columns) {
                    object? value = row.IsNull(column) ? null : row[column];
                    item[column.ColumnName] = value;
                }
                rows.Add(item);
            }

            return JsonSerializer.Serialize(rows, options);
        }

        private static DataTable CsvToDataTable(string csv, bool firstRowIsHeader) {
            if (csv == null) throw new ArgumentNullException(nameof(csv));

            var records = ParseCsv(csv).ToList();
            var table = new DataTable { Locale = CultureInfo.InvariantCulture };
            if (records.Count == 0) {
                return table;
            }

            int columnCount = records.Max(record => record.Count);
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

            for (int rowIndex = firstDataRow; rowIndex < records.Count; rowIndex++) {
                var record = records[rowIndex];
                DataRow row = table.NewRow();
                for (int column = 0; column < table.Columns.Count; column++) {
                    row[column] = column < record.Count && record[column] != null ? record[column] : DBNull.Value;
                }
                table.Rows.Add(row);
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
            var rows = new List<Dictionary<string, object?>>();
            foreach (JsonElement element in document.RootElement.EnumerateArray()) {
                if (element.ValueKind != JsonValueKind.Object) {
                    throw new ArgumentException("JSON input must be an array of objects.", nameof(json));
                }

                var row = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
                foreach (JsonProperty property in element.EnumerateObject()) {
                    if (!table.Columns.Contains(property.Name)) {
                        table.Columns.Add(property.Name, typeof(object));
                    }
                    row[property.Name] = JsonElementToValue(property.Value);
                }
                rows.Add(row);
            }

            foreach (var source in rows) {
                DataRow row = table.NewRow();
                foreach (DataColumn column in table.Columns) {
                    row[column] = source.TryGetValue(column.ColumnName, out object? value) && value != null ? value : DBNull.Value;
                }
                table.Rows.Add(row);
            }

            return table;
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

        private static IEnumerable<List<string?>> ParseCsv(string csv) {
            var record = new List<string?>();
            var field = new StringBuilder();
            bool inQuotes = false;
            bool quoted = false;

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
                    yield return record;
                    record = new List<string?>();
                    continue;
                }

                field.Append(ch);
            }

            if (field.Length > 0 || quoted || record.Count > 0) {
                record.Add(FieldValue(field, quoted));
                yield return record;
            }
        }

        private static string? FieldValue(StringBuilder field, bool quoted) {
            string value = field.ToString();
            return !quoted && value.Length == 0 ? null : value;
        }

        private static void AppendCsvField(StringBuilder builder, object? value) {
            if (value == null || value == DBNull.Value) {
                return;
            }

            string text = Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty;
            bool quote = text.IndexOfAny(new[] { ',', '"', '\r', '\n' }) >= 0;
            if (!quote) {
                builder.Append(text);
                return;
            }

            builder.Append('"');
            builder.Append(text.Replace("\"", "\"\""));
            builder.Append('"');
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
