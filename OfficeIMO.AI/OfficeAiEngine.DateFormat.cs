using System.Globalization;

namespace OfficeIMO.AI;

public sealed partial class OfficeAiEngine {
    // A Date field represents a full calendar date. Parsing must not invent a current year or first day.
    private static bool HasCompleteDateFormat(string format, CultureInfo culture) {
        if (format.Length == 1) {
            format = format[0] switch {
                'd' => culture.DateTimeFormat.ShortDatePattern,
                'D' => culture.DateTimeFormat.LongDatePattern,
                'o' or 'O' => "yyyy-MM-dd",
                'r' or 'R' => "ddd, dd MMM yyyy",
                _ => string.Empty
            };
        }
        bool year = false, month = false, day = false;
        for (int index = 0; index < format.Length; index++) {
            char token = format[index];
            if (token == '\\') { index++; continue; }
            if (token is '\'' or '"') {
                char quote = token;
                while (++index < format.Length && format[index] != quote)
                    if (format[index] == '\\') index++;
                continue;
            }
            if (token == '%') {
                if (++index >= format.Length) return false;
                token = format[index];
                if (token == 'd') day = true;
                if (token == 'M') month = true;
                if (token == 'y') year = true;
                continue;
            }
            int count = 1;
            while (index + 1 < format.Length && format[index + 1] == token) { count++; index++; }
            if (token == 'y') year = true;
            if (token == 'M') month = true;
            // ddd/dddd name a weekday rather than a day of the month.
            if (token == 'd' && count <= 2) day = true;
        }
        return year && month && day;
    }
}
