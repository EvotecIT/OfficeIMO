namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Applies reusable header-based column number formats to the worksheet.
        /// </summary>
        /// <param name="plan">Column format plan to apply.</param>
        /// <param name="includeHeader">True to include header cells in every rule unless a rule already includes them.</param>
        /// <param name="autoFit">True to auto-fit every resolved formatted column.</param>
        /// <param name="headerRow">Optional one-based row containing the headers used to resolve column names.</param>
        /// <param name="options">Read options used while resolving headers.</param>
        /// <returns>One result per configured rule.</returns>
        public IReadOnlyList<ExcelColumnFormatResult> ApplyColumnFormatPlan(
            ExcelColumnFormatPlan plan,
            bool includeHeader = false,
            bool autoFit = false,
            int? headerRow = null,
            ExcelReadOptions? options = null) {
            if (plan == null) throw new ArgumentNullException(nameof(plan));
            if (headerRow.HasValue && headerRow.Value <= 0) throw new ArgumentOutOfRangeException(nameof(headerRow), "Header row must be 1 or greater.");

            var results = new List<ExcelColumnFormatResult>(plan.Rules.Count);
            foreach (ExcelColumnFormatRule rule in plan.Rules) {
                string numberFormat = rule.ResolveNumberFormat();
                bool ruleIncludeHeader = includeHeader || rule.IncludeHeader;
                if (!TryResolveColumnFormatRule(rule, ruleIncludeHeader, headerRow, options, out int columnIndex, out bool directCandidateMatched, out bool directCandidateAvailable)) {
                    results.Add(new ExcelColumnFormatResult(
                        rule.Header,
                        columnIndex: null,
                        applied: false,
                        numberFormat,
                        warning: $"Header '{rule.Header}' was not found on worksheet '{Name}'."));
                    continue;
                }

                if (directCandidateMatched && !ruleIncludeHeader) {
                    bool updateMaterializedWorksheet = Document.HasMaterializedDirectTabularFastSaveWorksheet(this);
                    if (Document.TrySetDirectTabularSaveCandidateColumnNumberFormat(this, columnIndex, numberFormat)) {
                        if (updateMaterializedWorksheet) {
                            ApplyColumnNumberFormat(columnIndex, headerRow, ruleIncludeHeader, numberFormat);
                        }
                    } else {
                        ApplyColumnNumberFormat(columnIndex, headerRow, ruleIncludeHeader, numberFormat);
                    }
                } else if (directCandidateMatched) {
                    ColumnStyleByHeader(rule.Header, ruleIncludeHeader, options).NumberFormat(numberFormat);
                } else {
                    ApplyColumnNumberFormat(columnIndex, headerRow, ruleIncludeHeader, numberFormat);
                }

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

        private bool TryResolveColumnFormatRule(
            ExcelColumnFormatRule rule,
            bool includeHeader,
            int? headerRow,
            ExcelReadOptions? options,
            out int columnIndex,
            out bool directCandidateMatched,
            out bool directCandidateAvailable) {
            columnIndex = 0;
            directCandidateMatched = false;
            directCandidateAvailable = false;

            bool allowDirectCandidateLookup = !headerRow.HasValue || headerRow.Value == 1;
            if (allowDirectCandidateLookup
                && Document.TryGetDirectTabularSaveCandidateColumnByHeader(this, rule.Header, includeHeader, options, out columnIndex, out _, out _, out directCandidateAvailable)) {
                directCandidateMatched = true;
                return true;
            }

            if (allowDirectCandidateLookup && directCandidateAvailable && columnIndex > 0) {
                directCandidateMatched = true;
                return true;
            }

            return headerRow.HasValue
                ? TryGetColumnIndexByHeaderRow(rule.Header, headerRow.Value, out columnIndex, options)
                : TryGetColumnIndexByHeader(rule.Header, out columnIndex, options);
        }

        private bool TryGetColumnIndexByHeaderRow(string header, int headerRow, out int columnIndex, ExcelReadOptions? options) {
            columnIndex = 0;
            if (string.IsNullOrWhiteSpace(header)) {
                return false;
            }

            string usedRange = GetUsedRangeA1();
            var (_, firstColumn, _, lastColumn) = A1.ParseRange(usedRange);
            int columnCount = lastColumn - firstColumn + 1;
            if (columnCount <= 0) {
                return false;
            }

            var headerValues = new string?[columnCount];
            bool hasExplicitHeader = false;
            for (int offset = 0; offset < columnCount; offset++) {
                headerValues[offset] = TryGetCellText(headerRow, firstColumn + offset, out string text) ? text : null;
                if (!string.IsNullOrWhiteSpace(headerValues[offset])) {
                    hasExplicitHeader = true;
                }
            }

            if (!hasExplicitHeader) {
                return false;
            }

            bool normalizeHeaders = options?.NormalizeHeaders ?? true;
            var headers = ExcelHeaderNameHelper.BuildUniqueHeaders(columnCount, c => headerValues[c], normalizeHeaders);
            string normalizedHeader = ExcelHeaderNameHelper.NormalizeHeader(header, normalizeHeaders);
            for (int offset = 0; offset < headers.Length; offset++) {
                if (string.IsNullOrWhiteSpace(headerValues[offset])) {
                    continue;
                }

                if (string.Equals(headers[offset], normalizedHeader, StringComparison.OrdinalIgnoreCase)) {
                    columnIndex = firstColumn + offset;
                    return true;
                }
            }

            return false;
        }

        private void ApplyColumnNumberFormat(int columnIndex, int? headerRow, bool includeHeader, string numberFormat) {
            string usedRange = GetUsedRangeA1();
            var (usedStartRow, _, usedEndRow, _) = A1.ParseRange(usedRange);
            int startRow = headerRow.HasValue
                ? (includeHeader ? headerRow.Value : headerRow.Value + 1)
                : (includeHeader ? usedStartRow : usedStartRow + 1);
            if (startRow > usedEndRow) {
                return;
            }

            string column = A1.ColumnIndexToLetters(columnIndex);
            FormatRange(
                column + startRow.ToString(System.Globalization.CultureInfo.InvariantCulture) + ":" +
                column + usedEndRow.ToString(System.Globalization.CultureInfo.InvariantCulture),
                numberFormat);
        }
    }
}
