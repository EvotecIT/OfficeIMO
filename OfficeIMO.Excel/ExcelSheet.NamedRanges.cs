namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Creates or updates a named range in the workbook, scoped to this sheet.
        /// </summary>
        /// <param name="name">Defined name to create or update.</param>
        /// <param name="range">A1 range (e.g. "A1:B5").</param>
        /// <param name="save">When true, saves the workbook after the change.</param>
        public void SetNamedRange(string name, string range, bool save = true) {
            _excelDocument.SetNamedRange(name, range, this, save);
        }

        /// <summary>
        /// Gets the A1 range for the given defined name, scoped to this sheet when applicable.
        /// </summary>
        /// <param name="name">Defined name to resolve.</param>
        /// <returns>A1 range (e.g. "A1:B5") or null if not found.</returns>
        public string? GetNamedRange(string name) {
            return _excelDocument.GetNamedRange(name, this);
        }

        /// <summary>
        /// Returns all defined names visible to this sheet with their A1 ranges.
        /// </summary>
        public System.Collections.Generic.IReadOnlyDictionary<string, string> GetAllNamedRanges() {
            return _excelDocument.GetAllNamedRanges(this);
        }

        /// <summary>
        /// Removes a defined name scoped to this sheet.
        /// </summary>
        /// <param name="name">Defined name to remove.</param>
        /// <param name="save">When true, saves the workbook after removal.</param>
        /// <returns>True if removed; false if not found.</returns>
        public bool RemoveNamedRange(string name, bool save = true) {
            return _excelDocument.RemoveNamedRange(name, this, save);
        }
    }
}

