using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Fluent data validation builder for an <see cref="ExcelRange"/>.
    /// </summary>
    public sealed class ExcelRangeDataValidationBuilder {
        private readonly ExcelRange _range;

        internal ExcelRangeDataValidationBuilder(ExcelRange range) {
            _range = range ?? throw new ArgumentNullException(nameof(range));
        }

        /// <summary>
        /// Applies list validation from inline items.
        /// </summary>
        public ExcelRange List(params string[] items) {
            return List((IEnumerable<string>)items);
        }

        /// <summary>
        /// Applies list validation from inline items.
        /// </summary>
        public ExcelRange List(IEnumerable<string> items, bool allowBlank = true) {
            _range.Sheet.ValidationList(_range.Address, items, allowBlank);
            return _range;
        }

        /// <summary>
        /// Applies list validation using a workbook or sheet-local named range.
        /// </summary>
        public ExcelRange ListNamedRange(string namedRange, bool allowBlank = true) {
            _range.Sheet.ValidationListNamedRange(_range.Address, namedRange, allowBlank);
            return _range;
        }

        /// <summary>
        /// Applies list validation using a worksheet range.
        /// </summary>
        public ExcelRange ListRange(string sourceA1Range, string? sourceSheetName = null, bool allowBlank = true) {
            _range.Sheet.ValidationListRange(_range.Address, sourceA1Range, sourceSheetName, allowBlank);
            return _range;
        }

        /// <summary>
        /// Applies whole-number validation.
        /// </summary>
        public ExcelRange WholeNumber(DataValidationOperatorValues @operator, int formula1, int? formula2 = null, bool allowBlank = true, string? errorTitle = null, string? errorMessage = null) {
            _range.Sheet.ValidationWholeNumber(_range.Address, @operator, formula1, formula2, allowBlank, errorTitle, errorMessage);
            return _range;
        }

        /// <summary>
        /// Applies whole-number validation between two values.
        /// </summary>
        public ExcelRange WholeNumberBetween(int minimum, int maximum, bool allowBlank = true, string? errorTitle = null, string? errorMessage = null) {
            return WholeNumber(DataValidationOperatorValues.Between, minimum, maximum, allowBlank, errorTitle, errorMessage);
        }

        /// <summary>
        /// Applies decimal-number validation.
        /// </summary>
        public ExcelRange Decimal(DataValidationOperatorValues @operator, double formula1, double? formula2 = null, bool allowBlank = true, string? errorTitle = null, string? errorMessage = null) {
            _range.Sheet.ValidationDecimal(_range.Address, @operator, formula1, formula2, allowBlank, errorTitle, errorMessage);
            return _range;
        }

        /// <summary>
        /// Applies decimal-number validation between two values.
        /// </summary>
        public ExcelRange DecimalBetween(double minimum, double maximum, bool allowBlank = true, string? errorTitle = null, string? errorMessage = null) {
            return Decimal(DataValidationOperatorValues.Between, minimum, maximum, allowBlank, errorTitle, errorMessage);
        }

        /// <summary>
        /// Applies date validation.
        /// </summary>
        public ExcelRange Date(DataValidationOperatorValues @operator, DateTime formula1, DateTime? formula2 = null, bool allowBlank = true, string? errorTitle = null, string? errorMessage = null) {
            _range.Sheet.ValidationDate(_range.Address, @operator, formula1, formula2, allowBlank, errorTitle, errorMessage);
            return _range;
        }

        /// <summary>
        /// Applies date validation between two dates.
        /// </summary>
        public ExcelRange DateBetween(DateTime minimum, DateTime maximum, bool allowBlank = true, string? errorTitle = null, string? errorMessage = null) {
            return Date(DataValidationOperatorValues.Between, minimum, maximum, allowBlank, errorTitle, errorMessage);
        }

        /// <summary>
        /// Applies time validation.
        /// </summary>
        public ExcelRange Time(DataValidationOperatorValues @operator, TimeSpan formula1, TimeSpan? formula2 = null, bool allowBlank = true, string? errorTitle = null, string? errorMessage = null) {
            _range.Sheet.ValidationTime(_range.Address, @operator, formula1, formula2, allowBlank, errorTitle, errorMessage);
            return _range;
        }

        /// <summary>
        /// Applies text-length validation.
        /// </summary>
        public ExcelRange TextLength(DataValidationOperatorValues @operator, int formula1, int? formula2 = null, bool allowBlank = true, string? errorTitle = null, string? errorMessage = null) {
            _range.Sheet.ValidationTextLength(_range.Address, @operator, formula1, formula2, allowBlank, errorTitle, errorMessage);
            return _range;
        }

        /// <summary>
        /// Applies custom formula validation.
        /// </summary>
        public ExcelRange CustomFormula(string formula, bool allowBlank = true, string? errorTitle = null, string? errorMessage = null) {
            _range.Sheet.ValidationCustomFormula(_range.Address, formula, allowBlank, errorTitle, errorMessage);
            return _range;
        }

        /// <summary>
        /// Applies prompt and error message metadata to validations that overlap the range.
        /// </summary>
        public ExcelRange Messages(ExcelDataValidationMessageOptions options) {
            _range.Sheet.SetDataValidationMessages(_range.Address, options);
            return _range;
        }

        /// <summary>
        /// Removes data validations that overlap the range.
        /// </summary>
        public ExcelRange Clear() {
            _range.Sheet.RemoveDataValidations(_range.Address);
            return _range;
        }
    }
}
