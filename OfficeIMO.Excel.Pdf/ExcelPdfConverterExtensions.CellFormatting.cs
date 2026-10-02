namespace OfficeIMO.Excel.Pdf {
    public static partial class ExcelPdfConverterExtensions {
        private static string FormatCellValue(object? value, ExcelCellStyleSnapshot? style, string emptyCellText, ExcelDateSystem dateSystem) {
            if (value == null) {
                return emptyCellText;
            }

            string? formatCode = style?.NumberFormatCode;
            if (!string.IsNullOrWhiteSpace(formatCode)) {
                string? formatted = TryFormatCellValue(value, style!, formatCode!, dateSystem);
                if (formatted != null) {
                    return formatted;
                }
            }

            if (value is IFormattable formattable) {
                return formattable.ToString(null, CultureInfo.InvariantCulture) ?? emptyCellText;
            }

            return value.ToString() ?? emptyCellText;
        }

        private static string? TryFormatCellValue(object value, ExcelCellStyleSnapshot style, string formatCode, ExcelDateSystem dateSystem) {
            if (string.Equals(formatCode, "General", StringComparison.OrdinalIgnoreCase) || formatCode == "@") return null;
            double number;
            if (value is DateTime date) {
                // Styled reads keep numeric serials. A DateTime here is an explicit
                // ISO date cell, so retain its ticks instead of round-tripping through OA.
                number = ExcelDateSystemConverter.ToSerialForDisplay(date, dateSystem);
                string section = OfficeIMO.Spreadsheet.SpreadsheetNumberFormatDisplay.SelectNumberFormatSection(
                    formatCode, number < 0 ? 1 : number == 0 ? 2 : 0, number, out _);
                if (ExcelNumberFormatDisplay.IsDateNumberFormat(style.NumberFormatId, section)
                    && !OfficeIMO.Spreadsheet.SpreadsheetNumberFormatDisplay.HasElapsedTimeToken(section)) {
                    return OfficeIMO.Spreadsheet.SpreadsheetNumberFormatDisplay.FormatDateTimeValue(date, formatCode, number);
                }
            } else if (!TryGetDouble(value, out number)) {
                return null;
            }
            return ExcelNumberFormatDisplay.FormatNumericText(number, style.NumberFormatId,
                formatCode, number.ToString(CultureInfo.InvariantCulture), dateSystem);
        }

        private static bool TryGetDouble(object value, out double number) {
            switch (value) {
                case double doubleValue:
                    number = doubleValue;
                    return true;
                case float floatValue:
                    number = floatValue;
                    return true;
                case decimal decimalValue:
                    number = (double)decimalValue;
                    return true;
                case int intValue:
                    number = intValue;
                    return true;
                case long longValue:
                    number = longValue;
                    return true;
                case short shortValue:
                    number = shortValue;
                    return true;
                case byte byteValue:
                    number = byteValue;
                    return true;
                case sbyte sbyteValue:
                    number = sbyteValue;
                    return true;
                case ushort ushortValue:
                    number = ushortValue;
                    return true;
                case uint uintValue:
                    number = uintValue;
                    return true;
                case ulong ulongValue:
                    number = ulongValue;
                    return true;
                default:
                    number = default;
                    return false;
            }
        }

    }
}
