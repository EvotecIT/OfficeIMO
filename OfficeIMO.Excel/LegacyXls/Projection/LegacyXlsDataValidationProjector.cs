using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel.LegacyXls.Model;

namespace OfficeIMO.Excel.LegacyXls.Projection {
    internal static class LegacyXlsDataValidationProjector {
        internal static void Project(LegacyXlsWorkbook workbook, ExcelSheet sheet, LegacyXlsDataValidation validation) {
            switch (validation.Type) {
                case LegacyXlsDataValidationType.None:
                    ProjectAnyValue(sheet, validation);
                    break;
                case LegacyXlsDataValidationType.WholeNumber:
                    ProjectWholeNumber(sheet, validation);
                    break;
                case LegacyXlsDataValidationType.Decimal:
                    ProjectDecimal(sheet, validation);
                    break;
                case LegacyXlsDataValidationType.List:
                    ProjectList(sheet, validation);
                    break;
                case LegacyXlsDataValidationType.Date:
                    ProjectDate(sheet, validation);
                    break;
                case LegacyXlsDataValidationType.Time:
                    ProjectTime(sheet, validation);
                    break;
                case LegacyXlsDataValidationType.TextLength:
                    ProjectTextLength(sheet, validation);
                    break;
                case LegacyXlsDataValidationType.Custom:
                    ProjectCustom(sheet, validation);
                    break;
            }
        }

        private static void ProjectWholeNumber(ExcelSheet sheet, LegacyXlsDataValidation validation) {
            ProjectFormulaBackedValidation(sheet, validation, DataValidationValues.Whole);
        }

        private static void ProjectDecimal(ExcelSheet sheet, LegacyXlsDataValidation validation) {
            ProjectFormulaBackedValidation(sheet, validation, DataValidationValues.Decimal);
        }

        private static void ProjectList(ExcelSheet sheet, LegacyXlsDataValidation validation) {
            if (validation.ListItems.Count == 0
                && string.IsNullOrWhiteSpace(validation.ListSourceRange)
                && string.IsNullOrWhiteSpace(validation.ListSourceName)) {
                return;
            }

            string range = ToRange(validation);
            string formula;
            if (validation.ListItems.Count > 0) {
                formula = "\"" + string.Join(",", validation.ListItems.Select(item => item.Replace("\"", "\"\""))) + "\"";
            } else if (!string.IsNullOrWhiteSpace(validation.ListSourceRange)) {
                string sourceRange = validation.ListSourceRange!.Trim();
                if (sourceRange.StartsWith("=", StringComparison.Ordinal)) {
                    sourceRange = sourceRange.Substring(1).Trim();
                }

                if (!string.IsNullOrWhiteSpace(validation.ListSourceSheetName)) {
                    sourceRange = ExcelChartUtils.EnsureSheetQualifiedRange(validation.ListSourceSheetName!.Trim(), sourceRange);
                }

                formula = "=" + sourceRange;
            } else {
                string sourceName = validation.ListSourceName!.Trim();
                formula = sourceName.StartsWith("=", StringComparison.Ordinal) ? sourceName : "=" + sourceName;
            }

            DataValidation openXmlValidation = CreateOpenXmlValidation(validation, range, DataValidationValues.List, null);
            openXmlValidation.Append(new Formula1(formula));
            sheet.AppendLegacyDataValidation(openXmlValidation);
        }

        private static void ProjectDate(ExcelSheet sheet, LegacyXlsDataValidation validation) {
            ProjectFormulaBackedValidation(sheet, validation, DataValidationValues.Date);
        }

        private static void ProjectTime(ExcelSheet sheet, LegacyXlsDataValidation validation) {
            ProjectFormulaBackedValidation(sheet, validation, DataValidationValues.Time);
        }

        private static void ProjectTextLength(ExcelSheet sheet, LegacyXlsDataValidation validation) {
            ProjectFormulaBackedValidation(sheet, validation, DataValidationValues.TextLength);
        }

        private static void ProjectCustom(ExcelSheet sheet, LegacyXlsDataValidation validation) {
            if (string.IsNullOrWhiteSpace(validation.Formula1)) {
                return;
            }

            string range = ToRange(validation);
            DataValidation openXmlValidation = CreateOpenXmlValidation(validation, range, DataValidationValues.Custom, null);
            openXmlValidation.Append(new Formula1(validation.Formula1));
            sheet.AppendLegacyDataValidation(openXmlValidation);
        }

        private static void ProjectAnyValue(ExcelSheet sheet, LegacyXlsDataValidation validation) {
            string range = ToRange(validation);
            DataValidation openXmlValidation = CreateOpenXmlValidation(validation, range, DataValidationValues.None, null);
            sheet.AppendLegacyDataValidation(openXmlValidation);
        }

        private static void ProjectFormulaBackedValidation(ExcelSheet sheet, LegacyXlsDataValidation validation, DataValidationValues type) {
            if (string.IsNullOrWhiteSpace(validation.Formula1)) {
                return;
            }

            string range = ToRange(validation);
            DataValidation openXmlValidation = CreateOpenXmlValidation(validation, range, type, ToDataValidationOperator(validation.Operator));
            openXmlValidation.Append(new Formula1(validation.Formula1));
            string? formula2 = validation.Formula2;
            if (!string.IsNullOrWhiteSpace(formula2)) {
                openXmlValidation.Append(new Formula2(formula2!));
            }

            sheet.AppendLegacyDataValidation(openXmlValidation);
        }

        private static DataValidation CreateOpenXmlValidation(
            LegacyXlsDataValidation validation,
            string range,
            DataValidationValues type,
            DataValidationOperatorValues? @operator) {
            var openXmlValidation = new DataValidation {
                Type = type,
                AllowBlank = validation.AllowBlank,
                Operator = @operator,
                SequenceOfReferences = new DocumentFormat.OpenXml.ListValue<DocumentFormat.OpenXml.StringValue> { InnerText = range }
            };

            if (validation.ShowInputMessage
                && (!string.IsNullOrWhiteSpace(validation.PromptTitle) || !string.IsNullOrWhiteSpace(validation.Prompt))) {
                openXmlValidation.ShowInputMessage = true;
                openXmlValidation.PromptTitle = validation.PromptTitle;
                openXmlValidation.Prompt = validation.Prompt;
            }

            if (validation.ShowErrorMessage
                && (!string.IsNullOrWhiteSpace(validation.ErrorTitle) || !string.IsNullOrWhiteSpace(validation.Error))) {
                openXmlValidation.ShowErrorMessage = true;
                openXmlValidation.ErrorTitle = validation.ErrorTitle;
                openXmlValidation.Error = validation.Error;
            }

            openXmlValidation.ShowInputMessage = validation.ShowInputMessage;
            openXmlValidation.ShowErrorMessage = validation.ShowErrorMessage;
            openXmlValidation.PromptTitle = validation.PromptTitle;
            openXmlValidation.Prompt = validation.Prompt;
            openXmlValidation.ErrorTitle = validation.ErrorTitle;
            openXmlValidation.Error = validation.Error;
            if (validation.ErrorStyle != LegacyXlsDataValidationErrorStyle.Stop) {
                openXmlValidation.ErrorStyle = ToDataValidationErrorStyle(validation.ErrorStyle);
            }

            if (validation.SuppressDropDown) {
                openXmlValidation.ShowDropDown = true;
            }

            return openXmlValidation;
        }

        private static string ToRange(LegacyXlsDataValidation validation) {
            return string.Join(" ", validation.Ranges);
        }

        private static DataValidationOperatorValues ToDataValidationOperator(LegacyXlsDataValidationOperator @operator) {
            return @operator switch {
                LegacyXlsDataValidationOperator.Between => DataValidationOperatorValues.Between,
                LegacyXlsDataValidationOperator.NotBetween => DataValidationOperatorValues.NotBetween,
                LegacyXlsDataValidationOperator.Equal => DataValidationOperatorValues.Equal,
                LegacyXlsDataValidationOperator.NotEqual => DataValidationOperatorValues.NotEqual,
                LegacyXlsDataValidationOperator.GreaterThan => DataValidationOperatorValues.GreaterThan,
                LegacyXlsDataValidationOperator.LessThan => DataValidationOperatorValues.LessThan,
                LegacyXlsDataValidationOperator.GreaterThanOrEqual => DataValidationOperatorValues.GreaterThanOrEqual,
                _ => DataValidationOperatorValues.LessThanOrEqual
            };
        }

        private static DataValidationErrorStyleValues ToDataValidationErrorStyle(LegacyXlsDataValidationErrorStyle errorStyle) {
            return errorStyle switch {
                LegacyXlsDataValidationErrorStyle.Warning => DataValidationErrorStyleValues.Warning,
                LegacyXlsDataValidationErrorStyle.Information => DataValidationErrorStyleValues.Information,
                _ => DataValidationErrorStyleValues.Stop
            };
        }
    }
}
