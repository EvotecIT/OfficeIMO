using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Threading;

namespace OfficeIMO.Excel {
    public sealed partial class ExcelDocumentReader {
        /// <summary>
        /// Returns all Excel tables defined in the workbook.
        /// </summary>
        public IReadOnlyList<ExcelTableInfo> GetTables() {
            var workbook = WorkbookRoot;
            var sheets = workbook.Sheets?.OfType<Sheet>().ToList() ?? new List<Sheet>();
            var sheetLookup = new Dictionary<string, (string Name, int Index)>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < sheets.Count; i++) {
                var sheet = sheets[i];
                string? id = sheet.Id?.Value;
                string? name = sheet.Name?.Value;
                if (!string.IsNullOrWhiteSpace(id) && !string.IsNullOrWhiteSpace(name)) {
                    sheetLookup[id!] = (name!, i);
                }
            }

            var tables = new List<ExcelTableInfo>();
            foreach (var worksheetPart in WorkbookPartRoot.WorksheetParts) {
                string relId = WorkbookPartRoot.GetIdOfPart(worksheetPart);
                sheetLookup.TryGetValue(relId, out var sheetInfo);
                string sheetName = sheetInfo.Name ?? string.Empty;
                int sheetIndex = string.IsNullOrEmpty(sheetInfo.Name) ? -1 : sheetInfo.Index;

                foreach (var tablePart in worksheetPart.TableDefinitionParts) {
                    var table = tablePart.Table;
                    if (table == null) {
                        continue;
                    }

                    string name = table.Name?.Value ?? table.DisplayName?.Value ?? string.Empty;
                    string displayName = table.DisplayName?.Value ?? name;
                    string range = table.Reference?.Value ?? string.Empty;
                    var columns = table.TableColumns?.Elements<TableColumn>()
                        .Select((column, index) => new ExcelTableColumnInfo(
                            index + 1,
                            column.Name?.Value ?? string.Empty,
                            GetOpenXmlAttributeValue(column, "totalsRowFunction")))
                        .ToList() ?? new List<ExcelTableColumnInfo>();

                    tables.Add(new ExcelTableInfo(
                        name,
                        displayName,
                        range,
                        sheetName,
                        sheetIndex,
                        table.TableStyleInfo?.Name?.Value,
                        (table.HeaderRowCount?.Value ?? 1U) > 0U,
                        (table.TotalsRowCount?.Value ?? 0U) > 0U,
                        table.GetFirstChild<AutoFilter>() != null,
                        columns));
                }
            }

            return tables;
        }

        /// <summary>
        /// Gets metadata for a table by name or display name.
        /// </summary>
        public ExcelTableInfo GetTable(string tableName) {
            if (string.IsNullOrWhiteSpace(tableName)) throw new ArgumentNullException(nameof(tableName));

            var table = GetTables()
                .FirstOrDefault(item =>
                    string.Equals(item.Name, tableName, StringComparison.OrdinalIgnoreCase)
                    || string.Equals(item.DisplayName, tableName, StringComparison.OrdinalIgnoreCase));
            if (table == null) {
                throw new KeyNotFoundException($"Table '{tableName}' was not found.");
            }

            return table;
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
        public DataTable ReadTableAsDataTable(string tableName, bool? headersInFirstRow = null, ExecutionMode? mode = null, CancellationToken ct = default) {
            var table = GetTable(tableName);
            return GetSheet(table.SheetName).ReadRangeAsDataTable(table.Range, headersInFirstRow ?? table.HasHeaderRow, mode, ct);
        }

        /// <summary>
        /// Reads an Excel table by name as dictionaries using the table header row.
        /// </summary>
        public IEnumerable<Dictionary<string, object?>> ReadTableObjects(string tableName, ExecutionMode? mode = null, CancellationToken ct = default) {
            var table = GetTable(tableName);
            EnsureTableHasHeaderRow(table, nameof(ReadTableObjects));
            return GetSheet(table.SheetName).ReadObjects(table.Range, mode, ct);
        }

        /// <summary>
        /// Reads an Excel table by name and maps rows to objects using the table header row.
        /// </summary>
        public IEnumerable<T> ReadTableObjects<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] T>(
            string tableName,
            ExecutionMode? mode = null,
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

        private static string? GetOpenXmlAttributeValue(OpenXmlElement element, string localName) {
            if (element == null) throw new ArgumentNullException(nameof(element));
            if (string.IsNullOrWhiteSpace(localName)) throw new ArgumentException("Attribute name is required.", nameof(localName));

            var attribute = element.GetAttributes().FirstOrDefault(a => string.Equals(a.LocalName, localName, StringComparison.OrdinalIgnoreCase));
            return string.IsNullOrWhiteSpace(attribute.Value) ? null : attribute.Value;
        }
    }
}
