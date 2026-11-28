using OfficeIMO.Excel.OpenXml;
using OfficeIMO.OpenXml.Internal;
using System.Data;
using System.Diagnostics.CodeAnalysis;

namespace OfficeIMO.Excel {
    internal sealed partial class ExcelDocumentReader {
        /// <summary>
        /// Returns all Excel tables defined in the workbook.
        /// </summary>
        public IReadOnlyList<ExcelTableInfo> GetTables() {
            var sheetLookup = new Dictionary<string, (string Name, int Index)>(StringComparer.OrdinalIgnoreCase);
            int sheetIndex = 0;
            foreach (OfficeXmlNode sheet in ReadWorkbookNode.Child(Spreadsheet, "sheets").Children(Spreadsheet, "sheet")) {
                string? id = sheet.Attribute(Relationships, "id");
                string? name = sheet.Attribute("name");
                if (!string.IsNullOrWhiteSpace(id) && !string.IsNullOrWhiteSpace(name)) {
                    sheetLookup[id!] = (name!, sheetIndex);
                }
                sheetIndex++;
            }

            var tables = new List<ExcelTableInfo>();
            if (_liveWorkbook is ExcelPartRef workbook) {
                foreach ((string id, ExcelPartRef worksheet) in workbook.Parts) {
                    if (worksheet.Kind != ExcelPartKind.Worksheet) continue;
                    sheetLookup.TryGetValue(id, out var sheet);
                    foreach (ExcelPartRef table in worksheet.PartsOfKind(ExcelPartKind.Table)) {
                        AddTable(tables, table.Root, sheet);
                    }
                }
            } else {
                foreach ((string id, string worksheet) in _package!.GetWorksheetRelationships()) {
                    sheetLookup.TryGetValue(id, out var sheet);
                    foreach (string table in _package.GetWorksheetTablePartNames(worksheet, validate: true)) {
                        AddTable(tables, _package.ReadTableRoot(table, _opt), sheet);
                    }
                }
            }
            return tables;
        }

        private static void AddTable(List<ExcelTableInfo> tables, OfficeXmlNode table, (string? Name, int Index) sheet) {
            if (table.IsNone) return;
            string name = table.Attribute("name") ?? table.Attribute("displayName") ?? string.Empty;
            OfficeXmlNode style = table.Child(Spreadsheet, "tableStyleInfo");
            var columns = new List<ExcelTableColumnInfo>();
            OfficeXmlNode tableColumns = table.Child(Spreadsheet, "tableColumns");
            _ = ReadTableCount(tableColumns, "count"); // The editing model also rejects an invalid declared count.
            foreach (OfficeXmlNode column in tableColumns.Children(Spreadsheet, "tableColumn")) {
                columns.Add(new ExcelTableColumnInfo(columns.Count + 1, column.Attribute("name") ?? string.Empty,
                    GetTotalsRowFunction(column)));
            }
            tables.Add(new ExcelTableInfo(name, table.Attribute("displayName") ?? name, table.Attribute("ref") ?? string.Empty,
                sheet.Name ?? string.Empty, sheet.Name == null ? -1 : sheet.Index, style.Attribute("name"),
                (ReadTableCount(table, "headerRowCount") ?? 1U) > 0U, (ReadTableCount(table, "totalsRowCount") ?? 0U) > 0U,
                table.Child(Spreadsheet, "autoFilter").Exists, ReadTableFlag(style, "showFirstColumn"),
                ReadTableFlag(style, "showLastColumn"), ReadTableFlag(style, "showRowStripes"), ReadTableFlag(style, "showColumnStripes"), columns));
        }

        private static uint? ReadTableCount(OfficeXmlNode node, string name) =>
            node.Attribute(name) is string value ? System.Xml.XmlConvert.ToUInt32(value) : null;

        private static bool ReadTableFlag(OfficeXmlNode node, string name) =>
            node.Attribute(name) is string value && System.Xml.XmlConvert.ToBoolean(value);

        /// <summary>
        /// Gets metadata for a table by name or display name.
        /// </summary>
        public ExcelTableInfo GetTable(string tableName) {
            if (string.IsNullOrWhiteSpace(tableName)) throw new ArgumentNullException(nameof(tableName));

            foreach (var table in GetTables()) {
                if (string.Equals(table.Name, tableName, StringComparison.OrdinalIgnoreCase)
                    || string.Equals(table.DisplayName, tableName, StringComparison.OrdinalIgnoreCase)) {
                    return table;
                }
            }

            throw new KeyNotFoundException($"Table '{tableName}' was not found.");
        }

        /// <summary>
        /// Reads an Excel table by name into a dense matrix.
        /// </summary>
        public object?[,] ReadTable(string tableName) {
            var table = GetTable(tableName);
            return GetSheet(table.SheetName).ReadRange(table.Range);
        }

        /// <summary>
        /// Reads an Excel table by name into a DataTable.
        /// </summary>
        public DataTable ReadTableAsDataTable(string tableName, bool? headersInFirstRow = null, ExcelExecutionMode? mode = null, CancellationToken ct = default) {
            var table = GetTable(tableName);
            return GetSheet(table.SheetName).ReadRangeAsDataTable(table.Range, headersInFirstRow ?? table.HasHeaderRow, mode, ct);
        }

        /// <summary>
        /// Reads an Excel table by name as dictionaries using the table header row.
        /// </summary>
        public IEnumerable<Dictionary<string, object?>> ReadTableObjects(string tableName, ExcelExecutionMode? mode = null, CancellationToken ct = default) {
            var table = GetTable(tableName);
            EnsureTableHasHeaderRow(table, nameof(ReadTableObjects));
            return GetSheet(table.SheetName).ReadObjects(table.Range, mode, ct);
        }

        /// <summary>
        /// Reads an Excel table by name and maps rows to objects using the table header row.
        /// </summary>
        public IEnumerable<T> ReadTableObjects<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] T>(
            string tableName,
            ExcelExecutionMode? mode = null,
            CancellationToken ct = default) where T : new() {
            var table = GetTable(tableName);
            EnsureTableHasHeaderRow(table, nameof(ReadTableObjects));
            return GetSheet(table.SheetName).ReadObjects<T>(table.Range, mode, ct);
        }

        /// <summary>
        /// Streams an Excel table by name and maps rows to objects using the table header row.
        /// </summary>
        public IEnumerable<T> ReadTableObjectsStream<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] T>(
            string tableName,
            CancellationToken ct = default) where T : new() {
            var table = GetTable(tableName);
            EnsureTableHasHeaderRow(table, nameof(ReadTableObjectsStream));
            return GetSheet(table.SheetName).ReadObjectsStream<T>(table.Range, ct);
        }

        private static void EnsureTableHasHeaderRow(ExcelTableInfo table, string operationName) {
            if (!table.HasHeaderRow) {
                throw new InvalidOperationException($"{operationName} requires table '{table.Name}' to have a header row. Use ReadTableAsDataTable(..., headersInFirstRow: false) for headerless tables.");
            }
        }

        private static string? GetTotalsRowFunction(OfficeXmlNode column) {
            if (column.Attribute("totalsRowFunction") is string standard) {
                return string.IsNullOrWhiteSpace(standard) ? null : standard;
            }
            // Preserve the reader's interpretation of a lone nonstandard attribute.
            if (!column.HasAttributes) return null;
            foreach (OfficeXmlAttribute attribute in column.Attributes()) {
                if (string.Equals(attribute.LocalName, "totalsRowFunction", StringComparison.OrdinalIgnoreCase)) {
                    return string.IsNullOrWhiteSpace(attribute.Value) ? null : attribute.Value;
                }
            }
            return null;
        }
    }
}
