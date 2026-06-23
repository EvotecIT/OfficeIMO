namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Applies reusable header-based column number formats to the worksheet.
        /// </summary>
        /// <param name="plan">Column format plan to apply.</param>
        /// <param name="includeHeader">True to include header cells in every rule unless a rule already includes them.</param>
        /// <param name="autoFit">True to auto-fit every resolved formatted column.</param>
        /// <param name="options">Read options used while resolving headers.</param>
        /// <returns>One result per configured rule.</returns>
        public IReadOnlyList<ExcelColumnFormatResult> ApplyColumnFormatPlan(
            ExcelColumnFormatPlan plan,
            bool includeHeader = false,
            bool autoFit = false,
            ExcelReadOptions? options = null) {
            if (plan == null) throw new ArgumentNullException(nameof(plan));

            var results = new List<ExcelColumnFormatResult>(plan.Rules.Count);
            foreach (ExcelColumnFormatRule rule in plan.Rules) {
                string numberFormat = rule.ResolveNumberFormat();
                if (!TryGetColumnIndexByHeader(rule.Header, out int columnIndex, options)) {
                    results.Add(new ExcelColumnFormatResult(
                        rule.Header,
                        columnIndex: null,
                        applied: false,
                        numberFormat,
                        warning: $"Header '{rule.Header}' was not found on worksheet '{Name}'."));
                    continue;
                }

                ColumnStyleByHeader(rule.Header, includeHeader || rule.IncludeHeader, options).NumberFormat(numberFormat);
                if (autoFit || rule.AutoFit) {
                    AutoFitColumn(columnIndex);
                }

                results.Add(new ExcelColumnFormatResult(
                    rule.Header,
                    columnIndex,
                    applied: true,
                    numberFormat,
                    warning: null));
            }

            return results;
        }
    }
}
