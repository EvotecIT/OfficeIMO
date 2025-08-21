using System;
using OfficeIMO.Excel;
namespace OfficeIMO.Excel.Fluent {
    public class ColumnCollectionBuilder {
        private readonly ExcelSheet _sheet;

        internal ColumnCollectionBuilder(ExcelSheet sheet) {
            _sheet = sheet;
        }

        public ColumnCollectionBuilder Col(int index, Action<ColumnBuilder> action) {
            var builder = new ColumnBuilder(_sheet, index);
            action(builder);
            return this;
        }
    }
}
