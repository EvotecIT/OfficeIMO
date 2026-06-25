using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel.LegacyXls.Model;
using System.Globalization;

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
                    ProjectDate(workbook, sheet, validation);
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
            if (!int.TryParse(validation.Formula1, NumberStyles.Integer, CultureInfo.InvariantCulture, out int formula1)) {
                ProjectFormulaBackedValidation(sheet, validation, DataValidationValues.Whole);
                return;
            }

            int? formula2 = null;
            if (validation.Formula2 != null) {
                if (!int.TryParse(validation.Formula2, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsedFormula2)) {
                    ProjectFormulaBackedValidation(sheet, validation, DataValidationValues.Whole);
                    return;
                }

                formula2 = parsedFormula2;
            }

            string range = ToRange(validation);
            sheet.ValidationWholeNumber(
                range,
                ToDataValidationOperator(validation.Operator),
                formula1,
                formula2,
                validation.AllowBlank,
                validation.ShowErrorMessage ? validation.ErrorTitle : null,
                validation.ShowErrorMessage ? validation.Error : null);
            ProjectMessages(sheet, validation, range);
        }

        private static void ProjectDecimal(ExcelSheet sheet, LegacyXlsDataValidation validation) {
            if (!double.TryParse(validation.Formula1, NumberStyles.Float, CultureInfo.InvariantCulture, out double formula1)) {
                ProjectFormulaBackedValidation(sheet, validation, DataValidationValues.Decimal);
                return;
            }

            double? formula2 = null;
            if (validation.Formula2 != null) {
                if (!double.TryParse(validation.Formula2, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsedFormula2)) {
                    ProjectFormulaBackedValidation(sheet, validation, DataValidationValues.Decimal);
                    return;
                }

                formula2 = parsedFormula2;
            }

            string range = ToRange(validation);
            sheet.ValidationDecimal(
                range,
                ToDataValidationOperator(validation.Operator),
                formula1,
                formula2,
                validation.AllowBlank,
                validation.ShowErrorMessage ? validation.ErrorTitle : null,
                validation.ShowErrorMessage ? validation.Error : null);
            ProjectMessages(sheet, validation, range);
        }

        private static void ProjectList(ExcelSheet sheet, LegacyXlsDataValidation validation) {
            if (validation.ListItems.Count == 0
                && string.IsNullOrWhiteSpace(validation.ListSourceRange)
                && string.IsNullOrWhiteSpace(validation.ListSourceName)) {
                return;
            }

            string range = ToRange(validation);
            if (validation.ListItems.Count > 0) {
                sheet.ValidationList(range, validation.ListItems, validation.AllowBlank);
            } else if (!string.IsNullOrWhiteSpace(validation.ListSourceRange)) {
                sheet.ValidationListRange(range, validation.ListSourceRange!, validation.ListSourceSheetName, allowBlank: validation.AllowBlank);
            } else {
                sheet.ValidationListNamedRange(range, validation.ListSourceName!, validation.AllowBlank);
            }

            ProjectMessages(sheet, validation, range);
        }

        private static void ProjectDate(LegacyXlsWorkbook workbook, ExcelSheet sheet, LegacyXlsDataValidation validation) {
            if (!TryParseDate(workbook, validation.Formula1, out DateTime formula1)) {
                ProjectFormulaBackedValidation(sheet, validation, DataValidationValues.Date);
                return;
            }

            DateTime? formula2 = null;
            if (validation.Formula2 != null) {
                if (!TryParseDate(workbook, validation.Formula2, out DateTime parsedFormula2)) {
                    ProjectFormulaBackedValidation(sheet, validation, DataValidationValues.Date);
                    return;
                }

                formula2 = parsedFormula2;
            }

            string range = ToRange(validation);
            sheet.ValidationDate(
                range,
                ToDataValidationOperator(validation.Operator),
                formula1,
                formula2,
                validation.AllowBlank,
                validation.ShowErrorMessage ? validation.ErrorTitle : null,
                validation.ShowErrorMessage ? validation.Error : null);
            ProjectMessages(sheet, validation, range);
        }

        private static void ProjectTime(ExcelSheet sheet, LegacyXlsDataValidation validation) {
            if (!TryParseTime(validation.Formula1, out TimeSpan formula1)) {
                ProjectFormulaBackedValidation(sheet, validation, DataValidationValues.Time);
                return;
            }

            TimeSpan? formula2 = null;
            if (validation.Formula2 != null) {
                if (!TryParseTime(validation.Formula2, out TimeSpan parsedFormula2)) {
                    ProjectFormulaBackedValidation(sheet, validation, DataValidationValues.Time);
                    return;
                }

                formula2 = parsedFormula2;
            }

            string range = ToRange(validation);
            sheet.ValidationTime(
                range,
                ToDataValidationOperator(validation.Operator),
                formula1,
                formula2,
                validation.AllowBlank,
                validation.ShowErrorMessage ? validation.ErrorTitle : null,
                validation.ShowErrorMessage ? validation.Error : null);
            ProjectMessages(sheet, validation, range);
        }

        private static void ProjectTextLength(ExcelSheet sheet, LegacyXlsDataValidation validation) {
            if (!int.TryParse(validation.Formula1, NumberStyles.Integer, CultureInfo.InvariantCulture, out int formula1)) {
                ProjectFormulaBackedValidation(sheet, validation, DataValidationValues.TextLength);
                return;
            }

            int? formula2 = null;
            if (validation.Formula2 != null) {
                if (!int.TryParse(validation.Formula2, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsedFormula2)) {
                    ProjectFormulaBackedValidation(sheet, validation, DataValidationValues.TextLength);
                    return;
                }

                formula2 = parsedFormula2;
            }

            string range = ToRange(validation);
            sheet.ValidationTextLength(
                range,
                ToDataValidationOperator(validation.Operator),
                formula1,
                formula2,
                validation.AllowBlank,
                validation.ShowErrorMessage ? validation.ErrorTitle : null,
                validation.ShowErrorMessage ? validation.Error : null);
            ProjectMessages(sheet, validation, range);
        }

        private static void ProjectCustom(ExcelSheet sheet, LegacyXlsDataValidation validation) {
            if (string.IsNullOrWhiteSpace(validation.Formula1)) {
                return;
            }

            string range = ToRange(validation);
            sheet.ValidationCustomFormula(
                range,
                validation.Formula1,
                validation.AllowBlank,
                validation.ShowErrorMessage ? validation.ErrorTitle : null,
                validation.ShowErrorMessage ? validation.Error : null);
            ProjectMessages(sheet, validation, range);
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

            return openXmlValidation;
        }

        private static bool TryParseDate(LegacyXlsWorkbook workbook, string formula, out DateTime value) {
            value = default;
            return double.TryParse(formula, NumberStyles.Float, CultureInfo.InvariantCulture, out double serial)
                && LegacyXlsDateSerialConverter.TryConvert(serial, workbook.Uses1904DateSystem, out value);
        }

        private static bool TryParseTime(string formula, out TimeSpan value) {
            value = default;
            if (!double.TryParse(formula, NumberStyles.Float, CultureInfo.InvariantCulture, out double days)
                || double.IsNaN(days)
                || double.IsInfinity(days)) {
                return false;
            }

            try {
                value = TimeSpan.FromDays(days);
                return true;
            } catch (OverflowException) {
                return false;
            }
        }

        private static void ProjectMessages(ExcelSheet sheet, LegacyXlsDataValidation validation, string range) {
            if (validation.ShowInputMessage
                || validation.ShowErrorMessage
                || validation.PromptTitle != null
                || validation.Prompt != null
                || validation.ErrorTitle != null
                || validation.Error != null
                || validation.ErrorStyle != LegacyXlsDataValidationErrorStyle.Stop
                || validation.SuppressDropDown) {
                sheet.SetDataValidationMessages(range, new ExcelDataValidationMessageOptions {
                    PromptTitle = validation.PromptTitle,
                    Prompt = validation.Prompt,
                    ErrorTitle = validation.ErrorTitle,
                    Error = validation.Error,
                    ShowInputMessage = validation.ShowInputMessage,
                    ShowErrorMessage = validation.ShowErrorMessage,
                    PreserveShowMessageFlags = true,
                    ErrorStyle = ToDataValidationErrorStyle(validation.ErrorStyle),
                    SuppressDropDown = validation.SuppressDropDown ? true : null
                });
            }
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
