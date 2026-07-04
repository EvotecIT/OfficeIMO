using OfficeIMO.Excel.LegacyXls.Model;

namespace OfficeIMO.Excel.LegacyXls.Projection {
    internal static partial class LegacyXlsWorkbookProjector {
        private static void ApplyTableBlockLevelFormatting(
            LegacyXlsWorkbook workbook,
            ExcelSheet sheet,
            LegacyXlsTableDefinition tableDefinition) {
            LegacyXlsTableBlockLevelFormatting? formatting = tableDefinition.BlockLevelFormatting;
            if (formatting == null || !A1.TryParseStrictRange(tableDefinition.Range, out int firstRow, out int firstColumn, out int lastRow, out int lastColumn)) {
                return;
            }

            if (tableDefinition.HasHeaderRow
                && TryResolveTableRegionStyleIndex(workbook, sheet, formatting.HeaderStyleRecordIndex, out uint headerStyleIndex)) {
                ApplyCellStyleRange(sheet, firstRow, firstColumn, firstRow, lastColumn, headerStyleIndex);
            }

            int totalRows = checked((int)tableDefinition.TotalRowCount);
            int dataFirstRow = tableDefinition.HasHeaderRow ? firstRow + 1 : firstRow;
            int dataLastRow = Math.Max(dataFirstRow - 1, lastRow - totalRows);
            if (dataFirstRow <= dataLastRow
                && TryResolveTableRegionStyleIndex(workbook, sheet, formatting.DataStyleRecordIndex, out uint dataStyleIndex)) {
                ApplyCellStyleRange(sheet, dataFirstRow, firstColumn, dataLastRow, lastColumn, dataStyleIndex);
            }

            int totalFirstRow = lastRow - totalRows + 1;
            if (totalRows > 0
                && totalFirstRow <= lastRow
                && TryResolveTableRegionStyleIndex(workbook, sheet, formatting.TotalStyleRecordIndex, out uint totalStyleIndex)) {
                ApplyCellStyleRange(sheet, totalFirstRow, firstColumn, lastRow, lastColumn, totalStyleIndex);
            }
        }

        private static bool TryResolveTableRegionStyleIndex(
            LegacyXlsWorkbook workbook,
            ExcelSheet sheet,
            int? styleRecordIndex,
            out uint styleIndex) {
            styleIndex = 0U;
            if (!styleRecordIndex.HasValue
                || styleRecordIndex.Value < 0
                || styleRecordIndex.Value >= workbook.CellStyles.Count) {
                return false;
            }

            LegacyXlsCellStyle style = workbook.CellStyles[styleRecordIndex.Value];
            LegacyXlsCellFormat? format = workbook.GetCellFormat(style.StyleFormatIndex);
            if (format == null) {
                return false;
            }

            styleIndex = sheet.GetOrCreateLegacyCellFormatStyleIndex(workbook, format);
            return styleIndex != 0U;
        }

        private static void ApplyCellStyleRange(
            ExcelSheet sheet,
            int firstRow,
            int firstColumn,
            int lastRow,
            int lastColumn,
            uint styleIndex) {
            for (int row = firstRow; row <= lastRow; row++) {
                for (int column = firstColumn; column <= lastColumn; column++) {
                    sheet.SetCellStyleIndex(row, column, styleIndex, save: false);
                }
            }
        }
    }
}
