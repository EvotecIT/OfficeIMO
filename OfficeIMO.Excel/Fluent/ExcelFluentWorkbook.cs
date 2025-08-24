using System;
using OfficeIMO.Excel;
namespace OfficeIMO.Excel.Fluent {
    public class ExcelFluentWorkbook {
        internal ExcelDocument Workbook { get; }

        public ExcelFluentWorkbook(ExcelDocument workbook) {
            Workbook = workbook;
            // Favor stability for fluent scenarios: default to Sequential mode
            // Users can override (Workbook.Execution.Mode) if desired
            Workbook.Execution.Mode = ExecutionMode.Sequential;
        }

        public ExcelFluentWorkbook Sheet(string name, Action<SheetBuilder> action) {
            var builder = new SheetBuilder(this);
            builder.AddSheet(name);
            // Run all fluent operations in NoLock scope for stability
            using (Locking.EnterNoLockScope()) {
                action(builder);
            }
            return this;
        }

        public ExcelFluentWorkbook Sheet(Action<SheetBuilder> action) {
            var builder = new SheetBuilder(this);
            // Run all fluent operations in NoLock scope for stability
            using (Locking.EnterNoLockScope()) {
                action(builder);
            }
            return this;
        }

        public ExcelDocument End() {
            return Workbook;
        }
    }
}
