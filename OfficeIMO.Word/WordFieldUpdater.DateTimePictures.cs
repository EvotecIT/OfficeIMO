using System.Globalization;
using System.Text;

namespace OfficeIMO.Word {
    internal static partial class WordFieldUpdater {
        private const string DefaultDateTimeFormat = "yyyy-MM-dd HH:mm:ss";

        private enum DateTimePictureState {
            Absent,
            Empty,
            Present
        }

        internal static bool TryFormatDateTime(
            DateTime source,
            WordFieldInventory.ParsedFieldInstruction parsed,
            out string value,
            out string message) {
            if (!TryResolveDateTimeFormat(parsed.Switches, out string format, out string? customFormat, out message)) {
                value = string.Empty;
                return false;
            }
            try {
                value = source.ToString(format, CultureInfo.InvariantCulture);
                message = string.Empty;
                return true;
            } catch (FormatException) {
                value = string.Empty;
                message = "Date/time format switch " + customFormat +
                    " is not supported for deterministic field refresh.";
                return false;
            }
        }

        internal static bool TryFormatDateTime(
            DateTimeOffset source,
            WordFieldInventory.ParsedFieldInstruction parsed,
            out string value,
            out string message) {
            if (!TryResolveDateTimeFormat(parsed.Switches, out string format, out string? customFormat, out message)) {
                value = string.Empty;
                return false;
            }
            try {
                value = source.ToString(format, CultureInfo.InvariantCulture);
                message = string.Empty;
                return true;
            } catch (FormatException) {
                value = string.Empty;
                message = "Date/time format switch " + customFormat +
                    " is not supported for deterministic field refresh.";
                return false;
            }
        }

        private static bool TryResolveDateTimeFormat(
            IReadOnlyList<string> switches,
            out string format,
            out string? customFormat,
            out string message) {
            DateTimePictureState state = GetDateTimePicture(switches, out customFormat);
            if (state == DateTimePictureState.Empty) {
                format = string.Empty;
                message = "Date/time format switch has an empty picture and is not supported for deterministic field refresh.";
                return false;
            }
            format = state == DateTimePictureState.Absent
                ? DefaultDateTimeFormat
                : NormalizeDateTimeFormat(customFormat!);
            message = string.Empty;
            return true;
        }

        private static DateTimePictureState GetDateTimePicture(
            IReadOnlyList<string> switches,
            out string? format) {
            for (int index = switches.Count - 1; index >= 0; index--) {
                string fieldSwitch = switches[index].Trim();
                if (!fieldSwitch.StartsWith(@"\@", StringComparison.Ordinal)) continue;
                string rawFormat = fieldSwitch.Substring(2).Trim();
                format = TrimQuotes(rawFormat);
                return string.IsNullOrWhiteSpace(format)
                    ? DateTimePictureState.Empty
                    : DateTimePictureState.Present;
            }
            format = null;
            return DateTimePictureState.Absent;
        }

        private static string NormalizeDateTimeFormat(string format) {
            var normalized = new StringBuilder(format.Length);
            char quote = '\0';
            for (int index = 0; index < format.Length;) {
                char current = format[index];
                if (current == '\\' && index + 1 < format.Length) {
                    normalized.Append(current);
                    normalized.Append(format[index + 1]);
                    index += 2;
                    continue;
                }
                if (current == '\'' || current == '"') {
                    quote = quote == '\0' ? current : quote == current ? '\0' : quote;
                    normalized.Append(current);
                    index++;
                    continue;
                }
                if (quote == '\0' && MatchesDateTimeToken(format, index, "am/pm")) {
                    normalized.Append("tt");
                    index += "am/pm".Length;
                    continue;
                }
                if (quote == '\0' && MatchesDateTimeToken(format, index, "a/p")) {
                    normalized.Append('t');
                    index += "a/p".Length;
                    continue;
                }
                normalized.Append(current);
                index++;
            }
            return normalized.ToString();
        }

        private static bool MatchesDateTimeToken(string format, int index, string token) =>
            index + token.Length <= format.Length &&
            string.Compare(format, index, token, 0, token.Length, StringComparison.OrdinalIgnoreCase) == 0;
    }
}
