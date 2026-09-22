using OfficeIMO.Excel;
using OfficeIMO.OpenDocument;
using System.Globalization;

namespace OfficeIMO.Excel.OpenDocument;

public static partial class ExcelOpenDocumentConversionExtensions {
    private static bool TryFormatTemporalOperand(string? formula, OdsValidationValueKind kind,
        ExcelDateSystem dateSystem, out string? operand) {
        operand = null;
        if (!double.TryParse(formula, NumberStyles.Float, CultureInfo.InvariantCulture, out double serial)
            || double.IsNaN(serial) || double.IsInfinity(serial)) return false;

        if (kind == OdsValidationValueKind.Date) {
            DateTime date;
            try {
                date = ExcelDateSystemConverter.FromSerial(serial, dateSystem);
            } catch (ArgumentException) {
                return false;
            }
            if (date.TimeOfDay != TimeSpan.Zero) return false;
            operand = string.Format(CultureInfo.InvariantCulture, "DATE({0};{1};{2})",
                date.Year, date.Month, date.Day);
            return true;
        }

        if (kind != OdsValidationValueKind.Time || serial < 0 || serial >= 1) return false;
        double seconds = serial * 86400d;
        double rounded = Math.Round(seconds);
        // OpenFormula TIME permits evaluators to truncate fractional seconds.
        if (rounded >= 86400d || Math.Abs(seconds - rounded) > 0.0000001d) return false;
        var time = TimeSpan.FromSeconds(rounded);
        operand = string.Format(CultureInfo.InvariantCulture, "TIME({0};{1};{2})",
            time.Hours, time.Minutes, time.Seconds);
        return true;
    }

    private static bool TryParseDateOperand(string? operand, out DateTime value) {
        value = default;
        if (!TrySplitTemporalFunction(operand, "DATE", out string[]? arguments)
            || !int.TryParse(arguments![0], NumberStyles.Integer, CultureInfo.InvariantCulture, out int year)
            || !int.TryParse(arguments[1], NumberStyles.Integer, CultureInfo.InvariantCulture, out int month)
            || !int.TryParse(arguments[2], NumberStyles.Integer, CultureInfo.InvariantCulture, out int day)) return false;
        try {
            value = new DateTime(year, month, day);
            return true;
        } catch (ArgumentOutOfRangeException) {
            return false;
        }
    }

    private static bool TryParseTimeOperand(string? operand, out TimeSpan value) {
        value = default;
        if (!TrySplitTemporalFunction(operand, "TIME", out string[]? arguments)
            || !int.TryParse(arguments![0], NumberStyles.Integer, CultureInfo.InvariantCulture, out int hours)
            || !int.TryParse(arguments[1], NumberStyles.Integer, CultureInfo.InvariantCulture, out int minutes)
            || !decimal.TryParse(arguments[2], NumberStyles.AllowDecimalPoint,
                CultureInfo.InvariantCulture, out decimal secondValue)
            || secondValue != decimal.Truncate(secondValue)
            || hours < 0 || hours > 23 || minutes < 0 || minutes > 59
            || secondValue < 0 || secondValue > 59) return false;
        int seconds = (int)secondValue;
        value = new TimeSpan(hours, minutes, seconds);
        return true;
    }

    private static bool TryParseOptionalDateOperand(string? operand, out DateTime? value) {
        value = null;
        if (operand == null) return true;
        if (!TryParseDateOperand(operand, out DateTime parsed)) return false;
        value = parsed;
        return true;
    }

    private static bool TryParseOptionalTimeOperand(string? operand, out TimeSpan? value) {
        value = null;
        if (operand == null) return true;
        if (!TryParseTimeOperand(operand, out TimeSpan parsed)) return false;
        value = parsed;
        return true;
    }

    private static bool TrySplitTemporalFunction(string? operand, string name, out string[]? arguments) {
        arguments = null;
        if (operand == null || operand.Length > 64
            || !operand.StartsWith(name + "(", StringComparison.OrdinalIgnoreCase)
            || !operand.EndsWith(")", StringComparison.Ordinal)) return false;
        string body = operand.Substring(name.Length + 1, operand.Length - name.Length - 2);
        string[] parts = body.Split(';');
        if (parts.Length != 3) return false;
        for (int index = 0; index < parts.Length; index++) {
            parts[index] = parts[index].Trim();
            if (parts[index].Length == 0) return false;
        }
        arguments = parts;
        return true;
    }
}
