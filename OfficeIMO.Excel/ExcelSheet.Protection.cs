using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Gets whether the worksheet is protected.
        /// </summary>
        public bool IsProtected {
            get {
                var ws = WorksheetRoot;
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
                var ws = WorksheetRoot;
                var protection = ws.Elements<SheetProtection>().FirstOrDefault();
                if (protection == null) {
                    protection = new SheetProtection();
                    ws.Append(protection);
                }

                protection.Sheet = true;
                protection.SelectLockedCells = !opts.AllowSelectLockedCells;
                protection.SelectUnlockedCells = !opts.AllowSelectUnlockedCells;
                protection.FormatCells = !opts.AllowFormatCells;
                protection.FormatColumns = !opts.AllowFormatColumns;
                protection.FormatRows = !opts.AllowFormatRows;
                protection.InsertColumns = !opts.AllowInsertColumns;
                protection.InsertRows = !opts.AllowInsertRows;
                protection.InsertHyperlinks = !opts.AllowInsertHyperlinks;
                protection.DeleteColumns = !opts.AllowDeleteColumns;
                protection.DeleteRows = !opts.AllowDeleteRows;
                protection.Sort = !opts.AllowSort;
                protection.AutoFilter = !opts.AllowAutoFilter;
                protection.PivotTables = !opts.AllowPivotTables;
                SetOptionalProtectionFlag(protection, "objects", opts.ProtectObjects, value => protection.Objects = value);
                SetOptionalProtectionFlag(protection, "scenarios", opts.ProtectScenarios, value => protection.Scenarios = value);
                string? hash = ExcelProtectionHash.ResolveLegacyHash(opts.Password, opts.LegacyPasswordHash);
                if (hash != null) {
                    protection.Password = hash;
                } else {
                    protection.Password = null;
                    protection.RemoveAttribute("password", string.Empty);
                }

                EnsureWorksheetElementOrder();
                ws.Save();
            });
        }

        /// <summary>
        /// Applies worksheet protection with permissions for common Excel table editing workflows.
        /// </summary>
        public void ProtectTableEditing(string? password = null) {
            Protect(ExcelSheetProtectionOptions.TableEditing(password));
        }

        private static void SetOptionalProtectionFlag(SheetProtection protection, string attributeName, bool? value, Action<bool> assign) {
            if (value.HasValue) {
                assign(value.Value);
            } else {
                protection.RemoveAttribute(attributeName, string.Empty);
            }
        }

        /// <summary>
        /// Removes worksheet protection.
        /// </summary>
        public void Unprotect() {
            WriteLock(() => {
                var ws = WorksheetRoot;
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
