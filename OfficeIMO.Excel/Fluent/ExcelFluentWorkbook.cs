using System;
using OfficeIMO.Excel;
namespace OfficeIMO.Excel.Fluent {
    public class ExcelFluentWorkbook {
        internal ExcelDocument Workbook { get; }

        public ExcelFluentWorkbook(ExcelDocument workbook) {
            Workbook = workbook;
        }

        public ExcelFluentWorkbook Sheet(string name, Action<SheetBuilder> action) {
            var builder = new SheetBuilder(this);
            builder.AddSheet(name);
            action(builder);
            return this;
        }

        public ExcelFluentWorkbook Sheet(Action<SheetBuilder> action) {
            var builder = new SheetBuilder(this);
            action(builder);
            return this;
        }

        /// <summary>
        /// Ends fluent configuration and returns the underlying <see cref="ExcelDocument"/>.
        /// </summary>
        /// <returns>The wrapped <see cref="ExcelDocument"/> for further processing.</returns>
        public ExcelDocument End() {
            return Workbook;
        }
    }
}
