namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>Creates or replaces a formula definition scoped to this worksheet.</summary>
        /// <param name="name">The defined name used by worksheet formulas.</param>
        /// <param name="formula">An Excel expression, with or without a leading equals sign.</param>
        /// <param name="save">Whether to persist the workbook change immediately.</param>
        /// <param name="hidden">Whether to hide the name in Excel's normal name list.</param>
        /// <param name="validationMode">How the defined name is validated or sanitized.</param>
        public void SetNamedFormula(string name, string formula, bool save = true, bool hidden = false,
            ExcelDefinedNameValidationMode validationMode = ExcelDefinedNameValidationMode.Sanitize) =>
            _excelDocument.SetNamedFormula(name, formula, this, save, hidden, validationMode);

        /// <summary>
        /// Creates or updates a named range in the workbook, scoped to this sheet.
        /// </summary>
        /// <param name="name">Defined name to create or update.</param>
        /// <param name="range">A1 range (e.g. "A1:B5").</param>
        /// <param name="save">When true, saves the workbook after the change.</param>
        /// <param name="hidden">When true, marks the defined name as hidden.</param>
        public void SetNamedRange(string name, string range, bool save = true, bool hidden = false) {
            _excelDocument.SetNamedRange(name, range, this, save, hidden);
        }

        /// <summary>
        /// Creates or updates a named range in the workbook, scoped to this sheet, with validation control.
        /// </summary>
        /// <param name="name">Defined name to create or update.</param>
        /// <param name="range">A1 range (e.g. "A1:B5").</param>
        /// <param name="save">When true, saves the workbook after the change.</param>
        /// <param name="hidden">When true, marks the defined name as hidden.</param>
        /// <param name="validationMode">Controls how the name and range are validated: Sanitize (default) clamps/adjusts; Strict throws on invalid input.</param>
        public void SetNamedRange(string name, string range, bool save, bool hidden, ExcelDefinedNameValidationMode validationMode) {
            _excelDocument.SetNamedRange(name, range, this, save, hidden, validationMode);
        }

        /// <summary>
        /// Gets the range a defined name refers to, preferring a name local to this sheet over a workbook name.
        /// </summary>
        /// <param name="name">Defined name to resolve.</param>
        /// <returns>The range, or null when the name is missing or is not one cell range on an existing worksheet.</returns>
        public ExcelRange? GetNamedRange(string name) {
            return _excelDocument.GetNamedRange(name, this);
        }

        internal string? GetNamedRangeText(string name) => _excelDocument.GetNamedRangeText(name, this);

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

