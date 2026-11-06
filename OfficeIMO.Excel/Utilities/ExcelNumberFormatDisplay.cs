using System.Globalization;
using OfficeIMO.Spreadsheet;
using static OfficeIMO.Spreadsheet.SpreadsheetNumberFormatDisplay;

namespace OfficeIMO.Excel {
    internal static partial class ExcelNumberFormatDisplay {

        internal static string FormatNumericText(
            double value,
            uint numberFormatId,
            string? formatCode,
            string fallback,
            ExcelDateSystem dateSystem = ExcelDateSystem.NineteenHundred,
            IFormatProvider? provider = null) {
            if (numberFormatId == 0U) {
                return fallback;
            }

            string? resolvedFormatCode = ResolveFormatCode(numberFormatId, formatCode);
            if (string.IsNullOrWhiteSpace(resolvedFormatCode) || string.Equals(resolvedFormatCode, "General", StringComparison.OrdinalIgnoreCase)) {
                return fallback;
            }

            string nonEmptyFormatCode = resolvedFormatCode!;
            string section = SelectNumberFormatSection(nonEmptyFormatCode, value < 0 ? 1 : value == 0 ? 2 : 0, value, out _);
            if (IsDateNumberFormat(numberFormatId, section)) {
                return FormatDateValue(value, numberFormatId, nonEmptyFormatCode, dateSystem, provider);
            }

            return numberFormatId == 49U ? value.ToString(provider ?? CultureInfo.InvariantCulture)
                : FormatNumericValue(value, nonEmptyFormatCode, provider) ?? fallback;
        }

        internal static bool IsDateNumberFormat(uint numberFormatId, string? formatCode)
            => ExcelBuiltInNumberFormats.IsDate(numberFormatId)
            || ExcelNumberFormatClassifier.LooksLikeDateFormat(formatCode);

        private static string? ResolveFormatCode(uint numberFormatId, string? formatCode) =>
            string.IsNullOrWhiteSpace(formatCode)
                ? ExcelBuiltInNumberFormats.GetCode(numberFormatId)
                : formatCode;

        internal static bool TryGetDateSample(uint numberFormatId, string? formatCode, out string sample, double? value = null) {
            sample = string.Empty;
            string? resolved = ResolveFormatCode(numberFormatId, formatCode);
            if (resolved == null || HasElapsedTimeToken(resolved)) {
                return false;
            }

            switch (numberFormatId) {
                case 14:
                    sample = "12/31/9999";
                    return true;
                case 15:
                    sample = "30-Sep-99";
                    return true;
                case 16:
                    sample = "30-Sep";
                    return true;
                case 17:
                    sample = "Sep-99";
                    return true;
                case 18:
                    sample = "12:00 PM";
                    return true;
                case 19:
                    sample = "12:00:00 PM";
                    return true;
                case 20:
                    sample = "23:59";
                    return true;
                case 21:
                    sample = "23:59:59";
                    return true;
                case 22:
                    sample = "12/31/9999 23:59";
                    return true;
                case 45:
                    sample = "59:59";
                    return true;
                case 47:
                    sample = "59:59.0";
                    return true;
            }

            var sampleDate = new DateTime(2099, 12, 31, 23, 59, 59);
            string? formatted = FormatDateTimeSample(sampleDate, resolved, value ?? sampleDate.ToOADate());
            sample = formatted ?? string.Empty;
            return formatted != null;
        }

        private static string FormatDateValue(double value, uint numberFormatId, string formatCode, ExcelDateSystem dateSystem, IFormatProvider? provider = null) {
            if (TryFormatElapsedValue(value, formatCode, out string durationText)) return durationText;
            // An unsupported or out-of-range elapsed value must not become a calendar date.
            string section = SelectNumberFormatSection(formatCode, value < 0 ? 1 : value == 0 ? 2 : 0, value, out _);
            if (HasElapsedTimeToken(section)) return value.ToString(CultureInfo.InvariantCulture);
            DateTime date;
            try {
                date = HasFractionalSecondToken(section)
                    ? ExcelDateSystemConverter.FromSerialForDisplay(value, dateSystem)
                    : ExcelDateSystemConverter.FromSerial(value, dateSystem);
            } catch (ArgumentException) {
                return value.ToString(CultureInfo.InvariantCulture);
            }
            return FormatDateTimeText(date, numberFormatId, formatCode, value, provider)
                ?? value.ToString(provider ?? CultureInfo.InvariantCulture);
        }

        /// <summary>Shares ISO-cell and numeric-date section selection and calendar-safe provider formatting.</summary>
        internal static string? FormatDateTimeText(DateTime date, uint numberFormatId, string formatCode, double sectionValue, IFormatProvider? provider = null) {
            if (numberFormatId == 14U && provider != null) {
                var calendar = DateTimeFormatInfo.GetInstance(provider).Calendar;
                if (date < calendar.MinSupportedDateTime || date > calendar.MaxSupportedDateTime) return null;
                return date.ToString("d", provider);
            }
            return FormatDateTimeValue(date, formatCode, sectionValue, provider);
        }
    }
}
