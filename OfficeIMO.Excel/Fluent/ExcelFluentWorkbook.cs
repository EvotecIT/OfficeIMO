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
            
            // If in Sequential mode, run all operations in NoLock scope
            if (Workbook.Execution.Mode == ExecutionMode.Sequential) {
                using (Locking.EnterNoLockScope()) {
                    action(builder);
                }
            } else {
                action(builder);
            }
            return this;
        }

        public ExcelFluentWorkbook Sheet(Action<SheetBuilder> action) {
            var builder = new SheetBuilder(this);
            
            // If in Sequential mode, run all operations in NoLock scope
            if (Workbook.Execution.Mode == ExecutionMode.Sequential) {
                using (Locking.EnterNoLockScope()) {
                    action(builder);
                }
            } else {
                action(builder);
            }
            return this;
        }

        public ExcelDocument End() {
            return Workbook;
        }
    }
}
