using OfficeIMO.Excel.Utilities;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        // Two-cell markers describe grid positions. Reusing their native pixel extents with an
        // export fallback grid can both omit occupied cells and invent extra empty columns/rows.
        private static int ResolveVisualMarkerLastColumn(ExcelWorksheetGeometryIndex geometry,
            int? toColumn, int offset, int start, int width, IReadOnlyList<ExcelColumnInfo> columns,
            ExcelImageExportOptions options) {
            if (toColumn.HasValue && ExcelImageExportLimits.IsValidMarkerSpan(start - 1, toColumn.Value - 1, 16384)) {
                if (offset > 0) return ResolveLastVisualColumn(geometry, toColumn.Value, offset, columns, options);
                if (toColumn.Value > start) return toColumn.Value - 1;
            }
            return ResolveLastVisualColumn(geometry, start, width, columns, options);
        }

        private static int ResolveVisualMarkerLastRow(int? toRow, int offset, int start, int height,
            IReadOnlyDictionary<int, ExcelRowInfo> rows, bool defaultRowsHidden, ExcelImageExportOptions options) {
            if (toRow.HasValue && ExcelImageExportLimits.IsValidMarkerSpan(start - 1, toRow.Value - 1, 1048576)) {
                if (offset > 0) return ResolveLastVisualRow(toRow.Value, offset, rows, defaultRowsHidden, options);
                if (toRow.Value > start) return toRow.Value - 1;
            }
            return ResolveLastVisualRow(start, height, rows, defaultRowsHidden, options);
        }
    }
}
