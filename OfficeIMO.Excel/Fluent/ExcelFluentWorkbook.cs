using System;
using OfficeIMO.Excel;
namespace OfficeIMO.Excel.Fluent {
    /// <summary>
    /// Fluent API wrapper over <see cref="ExcelDocument"/> for quick workbook generation.
    /// </summary>
    public class ExcelFluentWorkbook {
        internal ExcelDocument Workbook { get; }

        /// <summary>Creates a fluent workbook wrapper. Defaults execution policy to Sequential for predictability.</summary>
        public ExcelFluentWorkbook(ExcelDocument workbook) {
            Workbook = workbook;
            // Favor stability for fluent scenarios: default to Sequential mode
            // Users can override (Workbook.Execution.Mode) if desired
            Workbook.Execution.Mode = ExecutionMode.Sequential;
        }

        /// <summary>
        /// Adds a new sheet and executes the builder action to populate it.
        /// </summary>
        public ExcelFluentWorkbook Sheet(string name, Action<SheetBuilder> action) {
            var builder = new SheetBuilder(this);
            builder.AddSheet(name);
            // Run all fluent operations in NoLock scope for stability
            using (Locking.EnterNoLockScope()) {
                action(builder);
            }
            return this;
        }

        /// <summary>
        /// Adds a new sheet with default name and executes the builder action.
        /// </summary>
        public ExcelFluentWorkbook Sheet(Action<SheetBuilder> action) {
            var builder = new SheetBuilder(this);
            // Run all fluent operations in NoLock scope for stability
            using (Locking.EnterNoLockScope()) {
                action(builder);
            }
            return this;
        }

        /// <summary>Finishes the fluent pipeline and returns the underlying <see cref="ExcelDocument"/>.</summary>
        public ExcelDocument End() {
            return Workbook;
        }
    }
}
