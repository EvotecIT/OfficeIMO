using System.Globalization;
using System.Text;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private string FormatFormulaDateText(double serial, DateTime date, string format) {
            if (_excelDocument.DateSystem != ExcelDateSystem.NineteenHundred || serial < 0d || serial >= 61d)
                return date.ToString(format.Length == 1 ? "%" + format : format, CultureInfo.InvariantCulture);

            // Render the early Excel calendar without forcing its fictitious date into DateTime.
            int year = Math.Floor(serial) == 0d ? 1900 : date.Year;
            int month = Math.Floor(serial) == 0d ? 1 : date.Month;
            int day = Math.Floor(serial) == 0d ? 0 : GetFormulaCalendarDay(serial, date);
            var outputFormat = new StringBuilder(format.Length);
            char quote = '\0';
            for (int index = 0; index < format.Length; index++) {
                char token = format[index];
                if (token == '\\' && index + 1 < format.Length) {
                    outputFormat.Append(token).Append(format[++index]);
                    continue;
                }
                if (token == '\'' || token == '"') {
                    if (quote == '\0') quote = token;
                    else if (quote == token) quote = '\0';
                    outputFormat.Append(token);
                    continue;
                }
                if (quote != '\0' || (token != 'y' && token != 'M' && token != 'd')) {
                    outputFormat.Append(token);
                    continue;
                }
                int count = 1;
                while (index + 1 < format.Length && format[index + 1] == token) { count++; index++; }
                string literal;
                if (token == 'y') literal = (count <= 2 ? year % 100 : year).ToString(new string('0', count), CultureInfo.InvariantCulture);
                else if (token == 'M') literal = count <= 2 ? month.ToString(new string('0', count), CultureInfo.InvariantCulture)
                    : count == 3 ? CultureInfo.InvariantCulture.DateTimeFormat.GetAbbreviatedMonthName(month)
                    : CultureInfo.InvariantCulture.DateTimeFormat.GetMonthName(month);
                else literal = count <= 2 ? day.ToString(new string('0', count), CultureInfo.InvariantCulture)
                    : count == 3 ? CultureInfo.InvariantCulture.DateTimeFormat.GetAbbreviatedDayName(GetFormulaWeekday(serial))
                    : CultureInfo.InvariantCulture.DateTimeFormat.GetDayName(GetFormulaWeekday(serial));
                if (outputFormat.Length > 0 && outputFormat[outputFormat.Length - 1] == '%') outputFormat.Length--;
                outputFormat.Append('\'').Append(literal).Append('\'');
            }
            return date.ToString(outputFormat.ToString(), CultureInfo.InvariantCulture);
        }
    }
}
