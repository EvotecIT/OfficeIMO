using System;
using OfficeIMO.Excel;
namespace OfficeIMO.Excel.Fluent {
    public class ColumnBuilder {
        private readonly ExcelSheet _sheet;
        private readonly int _columnIndex;

        internal ColumnBuilder(ExcelSheet sheet, int columnIndex) {
            if (columnIndex < 1) throw new ArgumentOutOfRangeException(nameof(columnIndex));
            _sheet = sheet;
            _columnIndex = columnIndex;
        }

        public ColumnBuilder AutoFit() {
            _sheet.AutoFitColumn(_columnIndex);
            return this;
        }

        public ColumnBuilder Width(double width) {
            _sheet.SetColumnWidth(_columnIndex, width);
            return this;
        }

        public ColumnBuilder Hidden(bool hidden) {
            _sheet.SetColumnHidden(_columnIndex, hidden);
            return this;
        }

        public ColumnBuilder Cell(int row, object? value = null, string? formula = null, string? numberFormat = null) {
            if (row < 1) throw new ArgumentOutOfRangeException(nameof(row));
            _sheet.Cell(row, _columnIndex, value, formula, numberFormat);
            return this;
        }
    }
}
