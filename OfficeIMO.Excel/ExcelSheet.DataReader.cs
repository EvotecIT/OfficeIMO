using System.Data;
using System.Globalization;
using System.Threading;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Streams rows from an <see cref="IDataReader"/> (including provider-owned DbDataReader implementations) into the worksheet and optionally creates an Excel table.
        /// The caller owns the connection, command, query, and provider.
        /// </summary>
        /// <param name="reader">Open data reader positioned before the first row.</param>
        /// <param name="startRow">1-based start row.</param>
        /// <param name="startColumn">1-based start column.</param>
        /// <param name="includeHeaders">Write field names as the first row.</param>
        /// <param name="tableName">Optional Excel table name.</param>
        /// <param name="style">Excel table style to use when <paramref name="createTable"/> is true.</param>
        /// <param name="includeAutoFilter">Include table AutoFilter dropdowns when creating a table.</param>
        /// <param name="createTable">Create an Excel table over the imported range.</param>
        /// <param name="autoFit">Auto-fit imported columns after rows are written.</param>
        /// <param name="ct">Cancellation token.</param>
        /// <returns>A1 range occupied by the imported reader data.</returns>
        public string InsertDataReader(
            IDataReader reader,
            int startRow = 1,
            int startColumn = 1,
            bool includeHeaders = true,
            string? tableName = null,
            TableStyle style = TableStyle.TableStyleMedium2,
            bool includeAutoFilter = true,
            bool createTable = true,
            bool autoFit = false,
            CancellationToken ct = default) {
            if (reader == null) throw new ArgumentNullException(nameof(reader));
            if (startRow < 1) throw new ArgumentOutOfRangeException(nameof(startRow));
            if (startColumn < 1) throw new ArgumentOutOfRangeException(nameof(startColumn));
            if (reader.FieldCount < 1) throw new ArgumentException("Data reader must expose at least one field.", nameof(reader));

            string[] headers = BuildReaderHeaders(reader);
            Type[] fieldTypes = BuildReaderFieldTypes(reader);
            int row = startRow;
            if (includeHeaders) {
                for (int i = 0; i < headers.Length; i++) {
                    CellValue(row, startColumn + i, headers[i]);
                }

                row++;
            }

            int dataRows = 0;
            while (reader.Read()) {
                ct.ThrowIfCancellationRequested();
                for (int i = 0; i < headers.Length; i++) {
                    object? value = reader.IsDBNull(i) ? null : reader.GetValue(i);
                    int column = startColumn + i;
                    CellValue(row, column, value);

                    string? numberFormat = GetReaderNumberFormat(fieldTypes[i], value);
                    if (!string.IsNullOrWhiteSpace(numberFormat)) {
                        FormatCell(row, column, numberFormat!);
                    }
                }

                row++;
                dataRows++;
            }

            int occupiedRows = dataRows + (includeHeaders ? 1 : 0);
            if (occupiedRows == 0) {
                return string.Empty;
            }

            string range = A1.CellReference(startRow, startColumn) + ":" +
                A1.CellReference(startRow + occupiedRows - 1, startColumn + headers.Length - 1);

            if (createTable) {
                AddTable(range, includeHeaders, tableName ?? string.Empty, style, includeAutoFilter);
            }

            if (autoFit) {
                AutoFitColumnsFor(Enumerable.Range(startColumn, headers.Length));
            }

            return range;
        }

        private static string[] BuildReaderHeaders(IDataReader reader) {
            var headers = new List<string>(reader.FieldCount);
            for (int i = 0; i < reader.FieldCount; i++) {
                string name;
                try {
                    name = reader.GetName(i);
                } catch (Exception) {
                    name = string.Empty;
                }

                if (string.IsNullOrWhiteSpace(name)) {
                    name = "Column" + (i + 1).ToString(CultureInfo.InvariantCulture);
                }

                headers.Add(name);
            }

            EnsureUniqueReaderHeaders(headers);
            return headers.ToArray();
        }

        private static Type[] BuildReaderFieldTypes(IDataReader reader) {
            var types = new Type[reader.FieldCount];
            for (int i = 0; i < reader.FieldCount; i++) {
                try {
                    types[i] = reader.GetFieldType(i) ?? typeof(object);
                } catch (Exception) {
                    types[i] = typeof(object);
                }
            }

            return types;
        }

        private static void EnsureUniqueReaderHeaders(IList<string> headers) {
            var seen = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < headers.Count; i++) {
                string baseName = string.IsNullOrWhiteSpace(headers[i])
                    ? "Column" + (i + 1).ToString(CultureInfo.InvariantCulture)
                    : headers[i].Trim();
                if (!seen.TryGetValue(baseName, out int count)) {
                    seen[baseName] = 1;
                    headers[i] = baseName;
                    continue;
                }

                count++;
                seen[baseName] = count;
                headers[i] = baseName + " (" + count.ToString(CultureInfo.InvariantCulture) + ")";
            }
        }

        private static string? GetReaderNumberFormat(Type fieldType, object? value) {
            Type type = Nullable.GetUnderlyingType(fieldType) ?? fieldType;
            if (type == typeof(DateTime) || type == typeof(DateTimeOffset) || value is DateTime || value is DateTimeOffset) {
                return "yyyy-mm-dd hh:mm";
            }

            if (type == typeof(TimeSpan) || value is TimeSpan) {
                return "[h]:mm:ss";
            }

            return null;
        }
    }
}
