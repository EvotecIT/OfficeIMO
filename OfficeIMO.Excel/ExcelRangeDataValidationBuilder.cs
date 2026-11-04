namespace OfficeIMO.Excel {
    /// <summary>
    /// Adds data validation rules to an <see cref="ExcelRange"/>.
    /// </summary>
    /// <remarks>
    /// Existing rules are listed, edited and removed through the worksheet
    /// (<see cref="ExcelSheet.GetDataValidations"/>, <see cref="ExcelSheet.SetDataValidationMessages"/>,
    /// <see cref="ExcelSheet.UpdateDataValidationMessages"/>, <see cref="ExcelSheet.RemoveDataValidations"/>).
    /// </remarks>
    public sealed class ExcelRangeDataValidationBuilder {
        private readonly ExcelRange _range;

        internal ExcelRangeDataValidationBuilder(ExcelRange range) {
            _range = range ?? throw new ArgumentNullException(nameof(range));
        }

        /// <summary>Restricts input to a drop-down list of inline items.</summary>
        public ExcelRange List(params string[] items) => List((IEnumerable<string>)items);

        /// <summary>Restricts input to a drop-down list of inline items.</summary>
        public ExcelRange List(IEnumerable<string> items, bool allowBlank = true) {
            _range.Sheet.CreateDataValidationList(_range.Reference, items, allowBlank);
            return _range;
        }

        /// <summary>Restricts input to a drop-down list read from a workbook or sheet-local named range.</summary>
        public ExcelRange ListNamedRange(string namedRange, bool allowBlank = true) {
            _range.Sheet.CreateDataValidationListNamedRange(_range.Reference, namedRange, allowBlank);
            return _range;
        }

        /// <summary>Restricts input to a drop-down list read from a worksheet range; the current sheet when no sheet is named.</summary>
        public ExcelRange ListRange(string sourceA1Range, string? sourceSheetName = null, bool allowBlank = true) {
            _range.Sheet.CreateDataValidationListRange(_range.Reference, sourceA1Range, sourceSheetName, allowBlank);
            return _range;
        }

        /// <summary>Restricts input to whole numbers; Between and NotBetween need <paramref name="formula2"/>.</summary>
        public ExcelRange WholeNumber(ExcelDataValidationOperator @operator, int formula1, int? formula2 = null, bool allowBlank = true, string? errorTitle = null, string? errorMessage = null) {
            _range.Sheet.CreateDataValidationWholeNumber(_range.Reference, @operator, formula1, formula2, allowBlank, errorTitle, errorMessage);
            return _range;
        }

        /// <summary>Restricts input to decimal numbers; Between and NotBetween need <paramref name="formula2"/>.</summary>
        public ExcelRange Decimal(ExcelDataValidationOperator @operator, double formula1, double? formula2 = null, bool allowBlank = true, string? errorTitle = null, string? errorMessage = null) {
            _range.Sheet.CreateDataValidationDecimal(_range.Reference, @operator, formula1, formula2, allowBlank, errorTitle, errorMessage);
            return _range;
        }

        /// <summary>Restricts input to dates; Between and NotBetween need <paramref name="formula2"/>.</summary>
        public ExcelRange Date(ExcelDataValidationOperator @operator, DateTime formula1, DateTime? formula2 = null, bool allowBlank = true, string? errorTitle = null, string? errorMessage = null) {
            _range.Sheet.CreateDataValidationDate(_range.Reference, @operator, formula1, formula2, allowBlank, errorTitle, errorMessage);
            return _range;
        }

        /// <summary>Restricts input to times of day; Between and NotBetween need <paramref name="formula2"/>.</summary>
        public ExcelRange Time(ExcelDataValidationOperator @operator, TimeSpan formula1, TimeSpan? formula2 = null, bool allowBlank = true, string? errorTitle = null, string? errorMessage = null) {
            _range.Sheet.CreateDataValidationTime(_range.Reference, @operator, formula1, formula2, allowBlank, errorTitle, errorMessage);
            return _range;
        }

        /// <summary>Restricts the length of text input; Between and NotBetween need <paramref name="formula2"/>.</summary>
        public ExcelRange TextLength(ExcelDataValidationOperator @operator, int formula1, int? formula2 = null, bool allowBlank = true, string? errorTitle = null, string? errorMessage = null) {
            _range.Sheet.CreateDataValidationTextLength(_range.Reference, @operator, formula1, formula2, allowBlank, errorTitle, errorMessage);
            return _range;
        }

        /// <summary>Accepts input when a custom formula evaluates to true.</summary>
        public ExcelRange CustomFormula(string formula, bool allowBlank = true, string? errorTitle = null, string? errorMessage = null) {
            _range.Sheet.CreateDataValidationCustomFormula(_range.Reference, formula, allowBlank, errorTitle, errorMessage);
            return _range;
        }
    }
}
