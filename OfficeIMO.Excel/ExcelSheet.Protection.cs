using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Gets whether the worksheet is protected.
        /// </summary>
        public bool IsProtected {
            get {
                var ws = _worksheetPart.Worksheet;
                return ws.Elements<SheetProtection>().Any();
            }
        }

        /// <summary>
        /// Applies worksheet protection using the provided options.
        /// </summary>
        /// <param name="options">Protection options (defaults allow selection of locked/unlocked cells).</param>
        public void Protect(ExcelSheetProtectionOptions? options = null) {
            var opts = options ?? new ExcelSheetProtectionOptions();

            WriteLock(() => {
                var ws = _worksheetPart.Worksheet;
                var protection = ws.Elements<SheetProtection>().FirstOrDefault();
                if (protection == null) {
                    protection = new SheetProtection();
                    ws.Append(protection);
                }

                protection.Sheet = true;
                protection.SelectLockedCells = opts.AllowSelectLockedCells;
                protection.SelectUnlockedCells = opts.AllowSelectUnlockedCells;
                protection.FormatCells = opts.AllowFormatCells;
                protection.FormatColumns = opts.AllowFormatColumns;
                protection.FormatRows = opts.AllowFormatRows;
                protection.InsertColumns = opts.AllowInsertColumns;
                protection.InsertRows = opts.AllowInsertRows;
                protection.InsertHyperlinks = opts.AllowInsertHyperlinks;
                protection.DeleteColumns = opts.AllowDeleteColumns;
                protection.DeleteRows = opts.AllowDeleteRows;
                protection.Sort = opts.AllowSort;
                protection.AutoFilter = opts.AllowAutoFilter;
                protection.PivotTables = opts.AllowPivotTables;

                EnsureWorksheetElementOrder();
                ws.Save();
            });
        }

        /// <summary>
        /// Removes worksheet protection.
        /// </summary>
        public void Unprotect() {
            WriteLock(() => {
                var ws = _worksheetPart.Worksheet;
                var protection = ws.Elements<SheetProtection>().FirstOrDefault();
                if (protection != null) {
                    ws.RemoveChild(protection);
                }

                var ranges = ws.Elements<ProtectedRanges>().FirstOrDefault();
                if (ranges != null) {
                    ws.RemoveChild(ranges);
                }

                ws.Save();
            });
        }
    }
}
