using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        internal void CleanupAutoFilterArtifacts() {
            var ws = WorksheetRoot;

            var tableRanges = new List<(int RowStart, int ColumnStart, int RowEnd, int ColumnEnd)>();
            foreach (var tableDefinitionPart in _worksheetPart.TableDefinitionParts) {
                var table = tableDefinitionPart.Table;
                var tableRange = table?.Reference?.Value;
                if (string.IsNullOrWhiteSpace(tableRange)) {
                    continue;
                }

                var tableAutoFilter = table?.Elements<AutoFilter>().FirstOrDefault();
                if (!A1.TryParseRange(tableRange!, out _, out int c1, out _, out int c2)) {
                    tableAutoFilter?.Remove();
                    tableDefinitionPart.Table?.Save();
                    continue;
                }

                tableRanges.Add(A1.ParseRange(tableRange!));
                if (tableAutoFilter == null) {
                    continue;
                }

                if (!string.Equals(tableAutoFilter.Reference?.Value, tableRange, StringComparison.OrdinalIgnoreCase)) {
                    tableAutoFilter.Reference = tableRange;
                }

                NormalizeAutoFilterColumns(tableAutoFilter, c2 - c1 + 1);
                tableDefinitionPart.Table?.Save();
            }

            var worksheetAutoFilter = ws.Elements<AutoFilter>().FirstOrDefault();
            if (worksheetAutoFilter == null) {
                return;
            }

            string? worksheetRange = worksheetAutoFilter.Reference?.Value;
            if (string.IsNullOrWhiteSpace(worksheetRange) || !A1.TryParseRange(worksheetRange!, out int wsR1, out int wsC1, out int wsR2, out int wsC2)) {
                ws.RemoveChild(worksheetAutoFilter);
                return;
            }

            NormalizeAutoFilterColumns(worksheetAutoFilter, wsC2 - wsC1 + 1);
            if (tableRanges.Any(tableRange => RangesOverlap(tableRange, (wsR1, wsC1, wsR2, wsC2)))) {
                ws.RemoveChild(worksheetAutoFilter);
            }
        }

        private static void NormalizeAutoFilterColumns(AutoFilter autoFilter, int width) {
            var seen = new HashSet<uint>();
            foreach (var filterColumn in autoFilter.Elements<FilterColumn>().ToList()) {
                uint? columnId = filterColumn.ColumnId?.Value;
                bool hasPayload = filterColumn.ChildElements.Any();
                if (!columnId.HasValue || columnId.Value >= width || !seen.Add(columnId.Value) || !hasPayload) {
                    filterColumn.Remove();
                }
            }
        }

        private static bool RangesOverlap((int RowStart, int ColumnStart, int RowEnd, int ColumnEnd) first, (int RowStart, int ColumnStart, int RowEnd, int ColumnEnd) second) {
            return first.RowStart <= second.RowEnd &&
                   first.RowEnd >= second.RowStart &&
                   first.ColumnStart <= second.ColumnEnd &&
                   first.ColumnEnd >= second.ColumnStart;
        }
    }
}
