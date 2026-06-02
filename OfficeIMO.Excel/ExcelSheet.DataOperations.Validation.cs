using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        // -------- Validation --------

        /// <summary>
        /// Applies a list validation to the specified A1 range using explicit items.
        /// </summary>
        public void ValidationList(string a1Range, IEnumerable<string> items, bool allowBlank = true) {
            if (string.IsNullOrWhiteSpace(a1Range)) throw new ArgumentNullException(nameof(a1Range));
            if (items == null) throw new ArgumentNullException(nameof(items));

            var joined = string.Join(",", items.Select(i => i?.Replace("\"", "\"\"") ?? string.Empty));
            var formula = "\"" + joined + "\""; // e.g., "New,Processed,Hold"

            var dv = new DataValidation {
                Type = DataValidationValues.List,
                AllowBlank = allowBlank,
                SequenceOfReferences = new ListValue<StringValue> { InnerText = a1Range }
            };
            dv.Append(new Formula1(formula));
            AppendDataValidation(dv);
        }

        /// <summary>
        /// Applies a list validation to the specified A1 range using a workbook or sheet-local named range.
        /// </summary>
        public void ValidationListNamedRange(string a1Range, string namedRange, bool allowBlank = true) {
            if (string.IsNullOrWhiteSpace(a1Range)) throw new ArgumentNullException(nameof(a1Range));
            if (string.IsNullOrWhiteSpace(namedRange)) throw new ArgumentNullException(nameof(namedRange));

            var normalizedNamedRange = namedRange.Trim();
            if (!normalizedNamedRange.StartsWith("=", StringComparison.Ordinal)) {
                normalizedNamedRange = "=" + normalizedNamedRange;
            }

            var dv = new DataValidation {
                Type = DataValidationValues.List,
                AllowBlank = allowBlank,
                SequenceOfReferences = new ListValue<StringValue> { InnerText = a1Range }
            };
            dv.Append(new Formula1(normalizedNamedRange));
            AppendDataValidation(dv);
        }

        /// <summary>
        /// Applies a list validation to the specified A1 range using a referenced worksheet range.
        /// When <paramref name="sourceSheetName"/> is omitted, the current worksheet is used.
        /// </summary>
        public void ValidationListRange(string a1Range, string sourceA1Range, string? sourceSheetName = null, bool allowBlank = true) {
            if (string.IsNullOrWhiteSpace(a1Range)) throw new ArgumentNullException(nameof(a1Range));
            if (string.IsNullOrWhiteSpace(sourceA1Range)) throw new ArgumentNullException(nameof(sourceA1Range));

            var normalizedSourceRange = sourceA1Range.Trim();
            if (normalizedSourceRange.StartsWith("=", StringComparison.Ordinal)) {
                normalizedSourceRange = normalizedSourceRange.Substring(1).Trim();
            }

            string formulaRange;
            if (string.IsNullOrWhiteSpace(sourceSheetName)) {
                formulaRange = normalizedSourceRange;
            } else {
                var effectiveSheetName = sourceSheetName!.Trim();
                formulaRange = ExcelChartUtils.EnsureSheetQualifiedRange(effectiveSheetName, normalizedSourceRange);
            }
            var formula = "=" + formulaRange;

            var dv = new DataValidation {
                Type = DataValidationValues.List,
                AllowBlank = allowBlank,
                SequenceOfReferences = new ListValue<StringValue> { InnerText = a1Range }
            };
            dv.Append(new Formula1(formula));
            AppendDataValidation(dv);
        }

        /// <summary>
        /// Applies a whole number validation to the specified A1 range.
        /// </summary>
        public void ValidationWholeNumber(string a1Range, DataValidationOperatorValues @operator, int formula1, int? formula2 = null, bool allowBlank = true, string? errorTitle = null, string? errorMessage = null) {
            string f1 = formula1.ToString(System.Globalization.CultureInfo.InvariantCulture);
            string? f2 = formula2?.ToString(System.Globalization.CultureInfo.InvariantCulture);
            ValidationAdd(a1Range, DataValidationValues.Whole, @operator, f1, f2, allowBlank, errorTitle, errorMessage);
        }

        /// <summary>
        /// Applies a decimal number validation to the specified A1 range.
        /// </summary>
        public void ValidationDecimal(string a1Range, DataValidationOperatorValues @operator, double formula1, double? formula2 = null, bool allowBlank = true, string? errorTitle = null, string? errorMessage = null) {
            string f1 = formula1.ToString(System.Globalization.CultureInfo.InvariantCulture);
            string? f2 = formula2?.ToString(System.Globalization.CultureInfo.InvariantCulture);
            ValidationAdd(a1Range, DataValidationValues.Decimal, @operator, f1, f2, allowBlank, errorTitle, errorMessage);
        }

        /// <summary>
        /// Applies a date validation to the specified A1 range.
        /// </summary>
        public void ValidationDate(string a1Range, DataValidationOperatorValues @operator, DateTime formula1, DateTime? formula2 = null, bool allowBlank = true, string? errorTitle = null, string? errorMessage = null) {
            string f1 = formula1.ToOADate().ToString(System.Globalization.CultureInfo.InvariantCulture);
            string? f2 = formula2?.ToOADate().ToString(System.Globalization.CultureInfo.InvariantCulture);
            ValidationAdd(a1Range, DataValidationValues.Date, @operator, f1, f2, allowBlank, errorTitle, errorMessage);
        }

        /// <summary>
        /// Applies a time validation to the specified A1 range.
        /// </summary>
        public void ValidationTime(string a1Range, DataValidationOperatorValues @operator, TimeSpan formula1, TimeSpan? formula2 = null, bool allowBlank = true, string? errorTitle = null, string? errorMessage = null) {
            string f1 = formula1.TotalDays.ToString(System.Globalization.CultureInfo.InvariantCulture);
            string? f2 = formula2?.TotalDays.ToString(System.Globalization.CultureInfo.InvariantCulture);
            ValidationAdd(a1Range, DataValidationValues.Time, @operator, f1, f2, allowBlank, errorTitle, errorMessage);
        }

        /// <summary>
        /// Applies a text length validation to the specified A1 range.
        /// </summary>
        public void ValidationTextLength(string a1Range, DataValidationOperatorValues @operator, int formula1, int? formula2 = null, bool allowBlank = true, string? errorTitle = null, string? errorMessage = null) {
            string f1 = formula1.ToString(System.Globalization.CultureInfo.InvariantCulture);
            string? f2 = formula2?.ToString(System.Globalization.CultureInfo.InvariantCulture);
            ValidationAdd(a1Range, DataValidationValues.TextLength, @operator, f1, f2, allowBlank, errorTitle, errorMessage);
        }

        /// <summary>
        /// Applies a custom formula validation to the specified A1 range.
        /// </summary>
        public void ValidationCustomFormula(string a1Range, string formula, bool allowBlank = true, string? errorTitle = null, string? errorMessage = null) {
            if (string.IsNullOrWhiteSpace(formula)) throw new ArgumentNullException(nameof(formula));
            ValidationAdd(a1Range, DataValidationValues.Custom, null, formula, null, allowBlank, errorTitle, errorMessage);
        }

        private void ValidationAdd(string a1Range, DataValidationValues type, DataValidationOperatorValues? @operator, string formula1, string? formula2, bool allowBlank, string? errorTitle, string? errorMessage) {
            if (string.IsNullOrWhiteSpace(a1Range)) throw new ArgumentNullException(nameof(a1Range));
            if (string.IsNullOrWhiteSpace(formula1)) throw new ArgumentNullException(nameof(formula1));

            bool requiresTwo = @operator == DataValidationOperatorValues.Between || @operator == DataValidationOperatorValues.NotBetween;
            if (requiresTwo && formula2 == null) throw new ArgumentNullException(nameof(formula2));

            DataValidation dv = new DataValidation {
                Type = type,
                AllowBlank = allowBlank,
                Operator = @operator,
                SequenceOfReferences = new ListValue<StringValue> { InnerText = a1Range }
            };

            if (!string.IsNullOrEmpty(errorTitle) || !string.IsNullOrEmpty(errorMessage)) {
                dv.ShowErrorMessage = true;
                dv.ErrorTitle = errorTitle;
                dv.Error = errorMessage;
            }

            dv.Append(new Formula1(formula1));
            if (formula2 != null) {
                dv.Append(new Formula2(formula2));
            }

            AppendDataValidation(dv);
        }

        private void AppendDataValidation(DataValidation dataValidation) {
            using var preserveDirectDataSet = _excelDocument.PreserveDirectDataSetSaveCandidateDuringDirtyMarks();
            WriteLockWorksheetPreparationOnly(() => {
                Worksheet ws = WorksheetRoot;
                DataValidations? dvs = ws.GetFirstChild<DataValidations>();
                if (dvs == null) {
                    dvs = new DataValidations();
                    InsertDataValidations(ws, dvs);
                }
                uint existingCount = dvs.Count?.Value ?? (uint)dvs.Elements<DataValidation>().Count();
                dvs.Append(dataValidation);
                dvs.Count = existingCount + 1U;
            });
        }

        private static void InsertDataValidations(Worksheet worksheet, DataValidations dataValidations) {
            var tableParts = worksheet.GetFirstChild<TableParts>();
            if (tableParts != null) {
                worksheet.InsertBefore(dataValidations, tableParts);
                return;
            }

            ConditionalFormatting? conditionalFormatting = null;
            foreach (var candidate in worksheet.Elements<ConditionalFormatting>()) {
                conditionalFormatting = candidate;
            }

            if (conditionalFormatting != null) {
                worksheet.InsertAfter(dataValidations, conditionalFormatting);
                return;
            }

            var autoFilter = worksheet.GetFirstChild<AutoFilter>();
            if (autoFilter != null) {
                worksheet.InsertAfter(dataValidations, autoFilter);
                return;
            }

            var sheetData = worksheet.GetFirstChild<SheetData>();
            if (sheetData != null) {
                worksheet.InsertAfter(dataValidations, sheetData);
            } else {
                worksheet.Append(dataValidations);
            }
        }
    }
}
