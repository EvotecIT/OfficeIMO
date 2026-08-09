using System.Globalization;
using System.Text;
using OfficeIMO.Excel;
using OfficeIMO.OpenDocument;

namespace OfficeIMO.Excel.OpenDocument;

public static partial class ExcelOpenDocumentConversionExtensions {
    private static bool TryCreateOdsValidationCondition(
        ExcelDataValidationSnapshot validation,
        out OdsValidationConditionSyntax? condition) {
        condition = null;
        string type = validation.Type?.Trim() ?? string.Empty;
        if (string.Equals(type, "list", StringComparison.OrdinalIgnoreCase)) {
            if (!TryParseExcelValidationList(validation.Formula1, out IReadOnlyList<string>? values)) return false;
            condition = OdsValidationConditionSyntax.CreateList(values!);
            return true;
        }

        OdsValidationValueKind valueKind;
        if (string.Equals(type, "whole", StringComparison.OrdinalIgnoreCase)) valueKind = OdsValidationValueKind.WholeNumber;
        else if (string.Equals(type, "decimal", StringComparison.OrdinalIgnoreCase)) valueKind = OdsValidationValueKind.DecimalNumber;
        else if (string.Equals(type, "textLength", StringComparison.OrdinalIgnoreCase)) valueKind = OdsValidationValueKind.TextLength;
        else return false;

        if (!TryMapValidationOperator(validation.Operator, out OdsValidationComparison comparison)
            || !IsInvariantValidationNumber(validation.Formula1, valueKind)
            || ((comparison == OdsValidationComparison.Between || comparison == OdsValidationComparison.NotBetween)
                && !IsInvariantValidationNumber(validation.Formula2, valueKind))) return false;
        condition = OdsValidationConditionSyntax.Create(valueKind, comparison, validation.Formula1!, validation.Formula2);
        return true;
    }

    private static bool TryMapValidationOperator(string? value, out OdsValidationComparison comparison) {
        switch (value?.Trim()) {
            case null:
            case "":
            case "between": comparison = OdsValidationComparison.Between; return true;
            case "notBetween": comparison = OdsValidationComparison.NotBetween; return true;
            case "equal": comparison = OdsValidationComparison.Equal; return true;
            case "notEqual": comparison = OdsValidationComparison.NotEqual; return true;
            case "lessThan": comparison = OdsValidationComparison.LessThan; return true;
            case "lessThanOrEqual": comparison = OdsValidationComparison.LessThanOrEqual; return true;
            case "greaterThan": comparison = OdsValidationComparison.GreaterThan; return true;
            case "greaterThanOrEqual": comparison = OdsValidationComparison.GreaterThanOrEqual; return true;
            default: comparison = default; return false;
        }
    }

    private static OdsValidationMessageType ParseOdsValidationMessageType(string? value) {
        if (string.Equals(value, "warning", StringComparison.OrdinalIgnoreCase)) return OdsValidationMessageType.Warning;
        if (string.Equals(value, "information", StringComparison.OrdinalIgnoreCase)) return OdsValidationMessageType.Information;
        return OdsValidationMessageType.Stop;
    }

    private static bool IsInvariantValidationNumber(string? value, OdsValidationValueKind valueKind) {
        if (string.IsNullOrWhiteSpace(value)) return false;
        if (valueKind == OdsValidationValueKind.WholeNumber || valueKind == OdsValidationValueKind.TextLength) {
            return int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out _);
        }
        return double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double number)
            && !double.IsNaN(number) && !double.IsInfinity(number);
    }

    private static bool TryParseExcelValidationList(string? formula, out IReadOnlyList<string>? values) {
        values = null;
        string text = formula?.Trim() ?? string.Empty;
        if (text.Length < 2 || text[0] != '"' || text[text.Length - 1] != '"') return false;
        string body = text.Substring(1, text.Length - 2);
        var result = new List<string>();
        var current = new StringBuilder();
        for (int index = 0; index < body.Length; index++) {
            if (body[index] == '"') {
                if (index + 1 >= body.Length || body[index + 1] != '"') return false;
                current.Append('"');
                index++;
            } else if (body[index] == ',') {
                result.Add(current.ToString());
                current.Clear();
            } else {
                current.Append(body[index]);
            }
        }
        result.Add(current.ToString());
        values = result;
        return true;
    }

    private static bool TryApplyOdsValidation(ExcelSheet sheet, string references, OdsValidation validation) {
        OdsValidationConditionSyntax? condition = validation.ParsedCondition;
        if (condition == null) return false;
        if (condition.ValueKind == OdsValidationValueKind.List) {
            sheet.ValidationList(references, condition.ListValues, validation.AllowEmptyCell);
            return true;
        }
        if (!TryMapValidationComparison(condition.Comparison, out ExcelDataValidationOperator comparison)) return false;

        switch (condition.ValueKind) {
            case OdsValidationValueKind.WholeNumber:
                if (!int.TryParse(condition.FirstOperand, NumberStyles.Integer, CultureInfo.InvariantCulture, out int wholeFirst)
                    || !TryParseOptionalInteger(condition.SecondOperand, out int? wholeSecond)) return false;
                sheet.ValidationWholeNumber(references, comparison, wholeFirst, wholeSecond, validation.AllowEmptyCell);
                return true;
            case OdsValidationValueKind.DecimalNumber:
                if (!double.TryParse(condition.FirstOperand, NumberStyles.Float, CultureInfo.InvariantCulture, out double decimalFirst)
                    || !TryParseOptionalDouble(condition.SecondOperand, out double? decimalSecond)
                    || double.IsNaN(decimalFirst) || double.IsInfinity(decimalFirst)) return false;
                sheet.ValidationDecimal(references, comparison, decimalFirst, decimalSecond, validation.AllowEmptyCell);
                return true;
            case OdsValidationValueKind.TextLength:
                if (!int.TryParse(condition.FirstOperand, NumberStyles.Integer, CultureInfo.InvariantCulture, out int lengthFirst)
                    || !TryParseOptionalInteger(condition.SecondOperand, out int? lengthSecond)) return false;
                sheet.ValidationTextLength(references, comparison, lengthFirst, lengthSecond, validation.AllowEmptyCell);
                return true;
            default:
                return false;
        }
    }

    private static void ApplyOdsValidationMessages(ExcelSheet sheet, string address, OdsValidation validation) {
        if (!validation.HasHelpMessage && !validation.HasErrorMessage) return;
        sheet.SetDataValidationMessages(address, new ExcelDataValidationMessageOptions {
            PromptTitle = validation.HelpTitle,
            Prompt = validation.HelpText,
            ShowInputMessage = validation.ShowHelpMessage,
            ErrorTitle = validation.ErrorTitle,
            Error = validation.ErrorText,
            ShowErrorMessage = validation.ShowErrorMessage,
            ErrorStyle = validation.ErrorMessageType switch {
                OdsValidationMessageType.Warning => ExcelDataValidationErrorStyle.Warning,
                OdsValidationMessageType.Information => ExcelDataValidationErrorStyle.Information,
                _ => ExcelDataValidationErrorStyle.Stop
            },
            PreserveShowMessageFlags = true
        });
    }

    private static bool TryMapValidationComparison(
        OdsValidationComparison? comparison,
        out ExcelDataValidationOperator value) {
        switch (comparison) {
            case OdsValidationComparison.Between: value = ExcelDataValidationOperator.Between; return true;
            case OdsValidationComparison.NotBetween: value = ExcelDataValidationOperator.NotBetween; return true;
            case OdsValidationComparison.Equal: value = ExcelDataValidationOperator.Equal; return true;
            case OdsValidationComparison.NotEqual: value = ExcelDataValidationOperator.NotEqual; return true;
            case OdsValidationComparison.LessThan: value = ExcelDataValidationOperator.LessThan; return true;
            case OdsValidationComparison.LessThanOrEqual: value = ExcelDataValidationOperator.LessThanOrEqual; return true;
            case OdsValidationComparison.GreaterThan: value = ExcelDataValidationOperator.GreaterThan; return true;
            case OdsValidationComparison.GreaterThanOrEqual: value = ExcelDataValidationOperator.GreaterThanOrEqual; return true;
            default: value = default; return false;
        }
    }

    private static bool TryParseOptionalInteger(string? text, out int? value) {
        value = null;
        if (text == null) return true;
        if (!int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsed)) return false;
        value = parsed;
        return true;
    }

    private static bool TryParseOptionalDouble(string? text, out double? value) {
        value = null;
        if (text == null) return true;
        if (!double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed)
            || double.IsNaN(parsed) || double.IsInfinity(parsed)) return false;
        value = parsed;
        return true;
    }
}
