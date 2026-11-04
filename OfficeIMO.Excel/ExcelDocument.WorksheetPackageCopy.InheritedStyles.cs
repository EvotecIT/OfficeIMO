using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        // Package copies may preserve malformed XML; invalid numeric style metadata must not abort copying valid cells.
        private static uint? ReadCopiedStyleIndex(UInt32Value? value) =>
            uint.TryParse(value?.InnerText, NumberStyles.Integer, CultureInfo.InvariantCulture, out uint index)
                ? index : null;

        private static HashSet<string> BuildInheritedStyleCellSet(Worksheet worksheet) {
            var styledRows = new HashSet<int>();
            foreach (Row row in worksheet.Descendants<Row>()) {
                if (ReadCopiedStyleIndex(row.StyleIndex).HasValue
                    && ReadCopiedStyleIndex(row.RowIndex) is uint index && index >= 1 && index <= 1048576) {
                    styledRows.Add((int)index);
                }
            }

            var styledColumns = new List<(int Min, int Max)>();
            foreach (Column column in worksheet.Elements<Columns>().SelectMany(columns => columns.Elements<Column>())) {
                if (!ReadCopiedStyleIndex(column.Style).HasValue) continue;
                uint? first = column.Min == null ? 1U : ReadCopiedStyleIndex(column.Min);
                uint? last = column.Max == null ? first : ReadCopiedStyleIndex(column.Max);
                if (first.HasValue && last.HasValue && first.Value >= 1 && first.Value <= last.Value && last.Value <= 16384) {
                    styledColumns.Add(((int)first.Value, (int)last.Value));
                }
            }

            var inherited = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (styledRows.Count == 0 && styledColumns.Count == 0) return inherited;
            foreach (Cell cell in worksheet.Descendants<Cell>()) {
                string? reference = cell.CellReference?.Value;
                if (ExcelA1.TryParseCellReferenceFast(reference, out int row, out int column)
                    && (styledRows.Contains(row) || styledColumns.Any(range => column >= range.Min && column <= range.Max))) {
                    inherited.Add(reference!);
                }
            }
            return inherited;
        }

        private static bool CellHasInheritedRowOrColumnStyle(Cell cell, HashSet<string> inheritedStyleCells) {
            string? reference = cell.CellReference?.Value;
            return !string.IsNullOrEmpty(reference) && inheritedStyleCells.Contains(reference!);
        }
    }
}
