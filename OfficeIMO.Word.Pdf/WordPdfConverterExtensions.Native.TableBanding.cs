using System.Collections.Generic;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        // Use the source table's body range and band sizes for every projected property.
        // PDF presentation stripes have a fixed phase and cannot represent Word band groups.
        private static NativeTableConditionalStyleDefaults GetNativeHorizontalBandStyle(WordTable table,
            NativeTableStyleDefaults defaults, int row, int headerRows, int footerStart) {
            if (table.ConditionalFormattingNoHorizontalBand == true || defaults.RowBandSize <= 0 ||
                row < headerRows || row >= footerStart) return NativeTableConditionalStyleDefaults.Empty;
            return (row - headerRows) / defaults.RowBandSize % 2 == 0
                ? defaults.Band1HorizontalStyle : defaults.Band2HorizontalStyle;
        }

        private static NativeTableConditionalStyleDefaults GetNativeVerticalBandStyle(WordTable table,
            NativeTableStyleDefaults defaults, int row, int column, int headerRows, int footerStart) {
            int bandColumn = column - (table.ConditionalFormattingFirstColumn == true ? 1 : 0);
            if (table.ConditionalFormattingNoVerticalBand == true || defaults.ColumnBandSize <= 0 ||
                bandColumn < 0 || row < headerRows || row >= footerStart) return NativeTableConditionalStyleDefaults.Empty;
            return bandColumn / defaults.ColumnBandSize % 2 == 0
                ? defaults.Band1VerticalStyle : defaults.Band2VerticalStyle;
        }

        private static void ApplyNativeTableConditionalFills(WordTable table, TableLayout layout,
            PdfCore.PdfTableStyle style, NativeTableStyleDefaults defaults) {
            int columns = GetNativeTableColumnCount(layout);
            int footerStart = table.ConditionalFormattingLastRow == true && layout.Rows.Count > style.HeaderRowCount
                ? layout.Rows.Count - 1 : layout.Rows.Count;
            var fills = style.CellFills == null ? new Dictionary<(int Row, int Column), global::OfficeIMO.OfficeColor>()
                : new Dictionary<(int Row, int Column), global::OfficeIMO.OfficeColor>(style.CellFills);
            for (int row = 0; row < layout.Rows.Count; row++) {
                int column = GetNativeTableRowStartColumn(layout, row);
                foreach (WordTableCell cell in layout.Rows[row]) {
                    if (IsNativeHorizontalMergeContinuation(cell)) continue;
                    int span = GetNativeCellColumnSpan(cell);
                    if (!IsNativeVerticalMergeContinuation(cell)) {
                        var resolved = GetNativeTableCellStyleDefaults(table, defaults, row, column, span,
                            columns, style.HeaderRowCount, footerStart);
                        if (resolved.CellFill.HasValue) fills[(row, column)] = resolved.CellFill.Value;
                    }
                    column += span;
                }
            }
            if (fills.Count > 0) style.CellFills = fills;
        }
    }
}
