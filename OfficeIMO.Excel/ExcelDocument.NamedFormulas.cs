namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        /// <summary>Creates or replaces a workbook or worksheet-scoped formula definition.</summary>
        /// <param name="name">The defined name used by worksheet formulas.</param>
        /// <param name="formula">An Excel expression, with or without a leading equals sign.</param>
        /// <param name="scope">The worksheet scope, or null for a workbook name.</param>
        /// <param name="save">Whether to persist the workbook change immediately.</param>
        /// <param name="hidden">Whether Excel hides the name from its normal name list.</param>
        /// <param name="validationMode">How the defined name is validated or sanitized.</param>
        /// <remarks>Storage preserves expressions outside the calculation subset. Calculation reports those expressions as unsupported.</remarks>
        public void SetNamedFormula(string name, string formula, ExcelSheet? scope = null, bool save = true,
            bool hidden = false, ExcelDefinedNameValidationMode validationMode = ExcelDefinedNameValidationMode.Sanitize) {
            if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("A defined name is required.", nameof(name));
            if (string.IsNullOrWhiteSpace(formula)) throw new ArgumentException("A formula is required.", nameof(formula));
            string expression = Utilities.ExcelSanitizer.SanitizeFormula(formula, qualifyArrayFunctions: true);
            if (string.IsNullOrWhiteSpace(expression)) throw new ArgumentException("A formula is required.", nameof(formula));
            StoreDefinedName(NormalizeDefinedName(name, validationMode), expression, scope, save, hidden);
        }
    }
}
