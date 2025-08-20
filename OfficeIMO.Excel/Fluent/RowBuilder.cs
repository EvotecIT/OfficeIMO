using OfficeIMO.Excel;
using System;

namespace OfficeIMO.Excel.Fluent {
    public class RowBuilder {
        private readonly ExcelSheet _sheet;
        private readonly int _rowIndex;

        internal RowBuilder(SheetBuilder sheetBuilder, ExcelSheet sheet, int rowIndex) {
            _sheet = sheet;
            _rowIndex = rowIndex;
        }

        public RowBuilder Cell(int column, object value) {
            _sheet.SetCellValue(_rowIndex, column, value);
            return this;
        }

        public RowBuilder Values(params object?[] values) {
            for (int i = 0; i < values.Length; i++) {
                _sheet.SetCellValue(_rowIndex, i + 1, values[i]);
            }
            return this;
        }
    }
}
