using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private bool TryEvaluateDateTimeFunction(string function, string args, out double result) {
            result = 0;
            var tokens = SplitFormulaArguments(args);

            if (function == "TODAY") {
                if (tokens.Count != 0) {
                    return false;
                }

                result = DateTime.Today.ToOADate();
                return true;
            }

            if (function == "NOW") {
                if (tokens.Count != 0) {
                    return false;
                }

                result = DateTime.Now.ToOADate();
                return true;
            }

            if (function == "NETWORKDAYS") {
                return TryEvaluateNetworkDays(tokens, out result);
            }

            if (function == "WORKDAY" || function == "WORKDAY.INTL") {
                return TryEvaluateWorkday(function, tokens, out result);
            }

            if (function == "DATEVALUE" || function == "TIMEVALUE") {
                return TryEvaluateDateTimeTextValue(function, tokens, out result);
            }

            if (function == "DATEDIF") {
                return TryEvaluateDateDif(tokens, out result);
            }

            if (function == "YEARFRAC") {
                return TryEvaluateYearFrac(tokens, out result);
            }

            if (!TryResolveFormulaOrNumericArguments(tokens, out var numbers)) {
                return false;
            }

            if (function == "DATE") {
                if (numbers.Count != 3
                    || !TryGetWholeNumber(numbers[0], out int year)
                    || !TryGetWholeNumber(numbers[1], out int month)
                    || !TryGetWholeNumber(numbers[2], out int day)) {
                    return false;
                }

                if (year >= 0 && year <= 1899) {
                    year += 1900;
                }

                if (year < 1 || year > 9999) {
                    return false;
                }

                try {
                    result = new DateTime(year, 1, 1).AddMonths(month - 1).AddDays(day - 1).ToOADate();
                } catch (ArgumentOutOfRangeException) {
                    return false;
                }

                return true;
            }

            if (function == "TIME") {
                if (numbers.Count != 3) {
                    return false;
                }

                double seconds = numbers[0] * 3600d + numbers[1] * 60d + numbers[2];
                seconds %= 86400d;
                if (seconds < 0) {
                    seconds += 86400d;
                }

                result = seconds / 86400d;
                return true;
            }

            if (function == "EDATE" || function == "EOMONTH") {
                if (numbers.Count != 2 || !TryGetWholeNumber(numbers[1], out int months)) {
                    return false;
                }

                if (!TryGetDateFromSerial(numbers[0], out DateTime startDate)) {
                    return false;
                }

                try {
                    DateTime shifted = startDate.AddMonths(months);
                    result = function == "EOMONTH"
                        ? new DateTime(shifted.Year, shifted.Month, DateTime.DaysInMonth(shifted.Year, shifted.Month)).ToOADate()
                        : shifted.ToOADate();
                } catch (ArgumentOutOfRangeException) {
                    return false;
                }

                return true;
            }

            if (function == "DAYS") {
                if (numbers.Count != 2
                    || !TryGetDateFromSerial(numbers[0], out DateTime endDate)
                    || !TryGetDateFromSerial(numbers[1], out DateTime startDate)) {
                    return false;
                }

                result = (endDate - startDate).TotalDays;
                return true;
            }

            if (function == "DAYS360") {
                if (tokens.Count < 2
                    || tokens.Count > 3
                    || !TryEvaluateFormulaOrNumeric(tokens[0], out double startSerial)
                    || !TryEvaluateFormulaOrNumeric(tokens[1], out double endSerial)
                    || !TryGetDateFromSerial(startSerial, out DateTime startDate)
                    || !TryGetDateFromSerial(endSerial, out DateTime endDate)) {
                    return false;
                }

                bool europeanMethod = false;
                if (tokens.Count == 3 && !TryResolveBooleanArgument(tokens[2], out europeanMethod)) {
                    return false;
                }

                result = europeanMethod ? Days360European(startDate, endDate) : Days360Us(startDate, endDate);
                return true;
            }

            if (function == "WEEKDAY") {
                if (numbers.Count < 1 || numbers.Count > 2
                    || !TryGetDateFromSerial(numbers[0], out DateTime date)) {
                    return false;
                }

                int returnType = 1;
                if (numbers.Count == 2 && !TryGetWholeNumber(numbers[1], out returnType)) {
                    return false;
                }

                int day = (int)date.DayOfWeek;
                if (returnType == 1) {
                    result = day + 1;
                    return true;
                }

                if (returnType == 2) {
                    result = day == 0 ? 7 : day;
                    return true;
                }

                if (returnType == 3) {
                    result = day == 0 ? 6 : day - 1;
                    return true;
                }

                return false;
            }

            if (function == "WEEKNUM" || function == "ISOWEEKNUM") {
                if (numbers.Count < 1
                    || numbers.Count > 2
                    || (function == "ISOWEEKNUM" && numbers.Count != 1)
                    || !TryGetDateFromSerial(numbers[0], out DateTime date)) {
                    return false;
                }

                if (function == "ISOWEEKNUM") {
                    result = GetIsoWeekNumber(date);
                    return true;
                }

                int returnType = 1;
                if (numbers.Count == 2 && !TryGetWholeNumber(numbers[1], out returnType)) {
                    return false;
                }

                return TryGetWeekStartDay(returnType, out DayOfWeek weekStart)
                    && TryGetWeekNumber(date, weekStart, returnType == 21, out result);
            }

            if (numbers.Count != 1) {
                return false;
            }

            DateTime dateTime;
            try {
                dateTime = DateTime.FromOADate(numbers[0]);
            } catch (ArgumentException) {
                return false;
            }

            switch (function) {
                case "YEAR":
                    result = dateTime.Year;
                    return true;
                case "MONTH":
                    result = dateTime.Month;
                    return true;
                case "DAY":
                    result = dateTime.Day;
                    return true;
                case "HOUR":
                    result = dateTime.Hour;
                    return true;
                case "MINUTE":
                    result = dateTime.Minute;
                    return true;
                case "SECOND":
                    result = dateTime.Second;
                    return true;
                default:
                    return false;
            }
        }

        private bool TryResolveFormulaOrNumericArguments(IReadOnlyList<string> tokens, out List<double> numbers) {
            numbers = new List<double>();
            foreach (string token in tokens) {
                if (!TryEvaluateFormulaOrNumeric(token, out double value)) {
                    return false;
                }

                numbers.Add(value);
            }

            return true;
        }

        private bool TryEvaluateNetworkDays(IReadOnlyList<string> tokens, out double result) {
            result = 0;
            if (tokens.Count < 2 || tokens.Count > 3
                || !TryEvaluateFormulaOrNumeric(tokens[0], out double startSerial)
                || !TryEvaluateFormulaOrNumeric(tokens[1], out double endSerial)
                || !TryGetDateFromSerial(startSerial, out DateTime startDate)
                || !TryGetDateFromSerial(endSerial, out DateTime endDate)) {
                return false;
            }

            var holidays = new HashSet<DateTime>();
            if (tokens.Count == 3 && !TryResolveHolidayDates(tokens[2], holidays)) {
                return false;
            }

            int direction = startDate <= endDate ? 1 : -1;
            DateTime current = direction == 1 ? startDate : endDate;
            DateTime last = direction == 1 ? endDate : startDate;
            int days = 0;
            while (current <= last) {
                if (current.DayOfWeek != DayOfWeek.Saturday
                    && current.DayOfWeek != DayOfWeek.Sunday
                    && !holidays.Contains(current.Date)) {
                    days++;
                }

                current = current.AddDays(1);
            }

            result = days * direction;
            return true;
        }

        private bool TryResolveHolidayDates(string token, HashSet<DateTime> holidays) {
            List<FormulaArgumentValue> values;
            if (token.IndexOf(':') >= 0) {
                if (!TryResolveFormulaRange(token, out values)) {
                    return false;
                }
            } else if (!TryResolveFormulaArguments(token, out values)) {
                return false;
            }

            foreach (var value in values) {
                if (value.Number.HasValue && TryGetDateFromSerial(value.Number.Value, out DateTime date)) {
                    holidays.Add(date);
                }
            }

            return true;
        }

        private bool TryEvaluateYearFrac(IReadOnlyList<string> tokens, out double result) {
            result = 0;
            if (tokens.Count < 2
                || tokens.Count > 3
                || !TryEvaluateFormulaOrNumeric(tokens[0], out double startSerial)
                || !TryEvaluateFormulaOrNumeric(tokens[1], out double endSerial)
                || !TryGetDateFromSerial(startSerial, out DateTime startDate)
                || !TryGetDateFromSerial(endSerial, out DateTime endDate)
                || endDate < startDate) {
                return false;
            }

            int basis = 0;
            if (tokens.Count == 3 && !TryGetWholeNumberArgument(tokens[2], out basis)) {
                return false;
            }

            switch (basis) {
                case 0:
                    result = Days360Us(startDate, endDate) / 360d;
                    return true;
                case 1:
                    result = ActualActualYearFraction(startDate, endDate);
                    return true;
                case 2:
                    result = (endDate - startDate).TotalDays / 360d;
                    return true;
                case 3:
                    result = (endDate - startDate).TotalDays / 365d;
                    return true;
                case 4:
                    result = Days360European(startDate, endDate) / 360d;
                    return true;
                default:
                    return false;
            }
        }

        private static int Days360Us(DateTime startDate, DateTime endDate) {
            int startDay = startDate.Day;
            int endDay = endDate.Day;

            if (startDay == 31 || IsLastDayOfFebruary(startDate)) {
                startDay = 30;
            }

            if (endDay == 31 && startDay >= 30) {
                endDay = 30;
            }

            return ((endDate.Year - startDate.Year) * 360)
                + ((endDate.Month - startDate.Month) * 30)
                + endDay - startDay;
        }

        private static int Days360European(DateTime startDate, DateTime endDate) {
            int startDay = Math.Min(startDate.Day, 30);
            int endDay = Math.Min(endDate.Day, 30);
            return ((endDate.Year - startDate.Year) * 360)
                + ((endDate.Month - startDate.Month) * 30)
                + endDay - startDay;
        }

        private static bool TryGetWeekStartDay(int returnType, out DayOfWeek weekStart) {
            switch (returnType) {
                case 1:
                case 17:
                    weekStart = DayOfWeek.Sunday;
                    return true;
                case 2:
                case 11:
                case 21:
                    weekStart = DayOfWeek.Monday;
                    return true;
                case 12:
                    weekStart = DayOfWeek.Tuesday;
                    return true;
                case 13:
                    weekStart = DayOfWeek.Wednesday;
                    return true;
                case 14:
                    weekStart = DayOfWeek.Thursday;
                    return true;
                case 15:
                    weekStart = DayOfWeek.Friday;
                    return true;
                case 16:
                    weekStart = DayOfWeek.Saturday;
                    return true;
                default:
                    weekStart = DayOfWeek.Sunday;
                    return false;
            }
        }

        private static bool TryGetWeekNumber(DateTime date, DayOfWeek weekStart, bool isoSystem, out double result) {
            if (isoSystem) {
                result = GetIsoWeekNumber(date);
                return true;
            }

            DateTime firstDay = new DateTime(date.Year, 1, 1);
            DateTime firstWeekStart = firstDay.AddDays(-GetDayOffset(firstDay.DayOfWeek, weekStart));
            result = Math.Floor((date.Date - firstWeekStart).TotalDays / 7d) + 1d;
            return result >= 1d && result <= 54d;
        }

        private static int GetIsoWeekNumber(DateTime date) {
            DayOfWeek day = CultureInfo.InvariantCulture.Calendar.GetDayOfWeek(date);
            if (day >= DayOfWeek.Monday && day <= DayOfWeek.Wednesday) {
                date = date.AddDays(3);
            }

            return CultureInfo.InvariantCulture.Calendar.GetWeekOfYear(
                date,
                CalendarWeekRule.FirstFourDayWeek,
                DayOfWeek.Monday);
        }

        private static int GetDayOffset(DayOfWeek day, DayOfWeek weekStart) {
            int offset = (int)day - (int)weekStart;
            return offset < 0 ? offset + 7 : offset;
        }

        private static double ActualActualYearFraction(DateTime startDate, DateTime endDate) {
            if (startDate == endDate) {
                return 0d;
            }

            if (startDate.Year == endDate.Year) {
                return (endDate - startDate).TotalDays / DaysInYear(startDate.Year);
            }

            double fraction = (new DateTime(startDate.Year + 1, 1, 1) - startDate).TotalDays / DaysInYear(startDate.Year);
            for (int year = startDate.Year + 1; year < endDate.Year; year++) {
                fraction += 1d;
            }

            fraction += (endDate - new DateTime(endDate.Year, 1, 1)).TotalDays / DaysInYear(endDate.Year);
            return fraction;
        }

        private static int DaysInYear(int year) {
            return DateTime.IsLeapYear(year) ? 366 : 365;
        }

        private static bool IsLastDayOfFebruary(DateTime date) {
            return date.Month == 2 && date.Day == DateTime.DaysInMonth(date.Year, 2);
        }

        private bool TryEvaluateDateDif(IReadOnlyList<string> tokens, out double result) {
            result = 0;
            if (tokens.Count != 3
                || !TryEvaluateFormulaOrNumeric(tokens[0], out double startSerial)
                || !TryEvaluateFormulaOrNumeric(tokens[1], out double endSerial)
                || !TryGetDateFromSerial(startSerial, out DateTime startDate)
                || !TryGetDateFromSerial(endSerial, out DateTime endDate)
                || endDate < startDate
                || !TryResolveTextArgument(tokens[2], out string unit)) {
                return false;
            }

            switch (unit.ToUpperInvariant()) {
                case "D":
                    result = (endDate - startDate).TotalDays;
                    return true;
                case "M":
                    result = GetCompletedMonths(startDate, endDate);
                    return true;
                case "Y":
                    result = GetCompletedYears(startDate, endDate);
                    return true;
                case "YM":
                    result = GetRemainingCompletedMonthsAfterYears(startDate, endDate);
                    return true;
                case "YD":
                    result = GetDaysAfterLastAnniversary(startDate, endDate);
                    return true;
                case "MD":
                    result = GetRemainingDaysAfterMonths(startDate, endDate);
                    return true;
                default:
                    return false;
            }
        }

        private static int GetCompletedYears(DateTime startDate, DateTime endDate) {
            int years = endDate.Year - startDate.Year;
            if (endDate < AddYearsClamped(startDate, years)) {
                years--;
            }

            return years;
        }

        private static int GetCompletedMonths(DateTime startDate, DateTime endDate) {
            int months = (endDate.Year - startDate.Year) * 12 + endDate.Month - startDate.Month;
            if (endDate.Day < startDate.Day) {
                months--;
            }

            return months;
        }

        private static int GetRemainingCompletedMonthsAfterYears(DateTime startDate, DateTime endDate) {
            int years = GetCompletedYears(startDate, endDate);
            DateTime anniversary = AddYearsClamped(startDate, years);
            int months = endDate.Month - anniversary.Month;
            if (months < 0) {
                months += 12;
            }

            if (endDate.Day < anniversary.Day) {
                months--;
                if (months < 0) {
                    months += 12;
                }
            }

            return months;
        }

        private static int GetDaysAfterLastAnniversary(DateTime startDate, DateTime endDate) {
            DateTime anniversary = CreateClampedDate(endDate.Year, startDate.Month, startDate.Day);
            if (anniversary > endDate) {
                anniversary = CreateClampedDate(endDate.Year - 1, startDate.Month, startDate.Day);
            }

            return (int)(endDate - anniversary).TotalDays;
        }

        private static int GetRemainingDaysAfterMonths(DateTime startDate, DateTime endDate) {
            if (endDate.Day >= startDate.Day) {
                return endDate.Day - startDate.Day;
            }

            DateTime previousMonth = endDate.AddMonths(-1);
            int daysInPreviousMonth = DateTime.DaysInMonth(previousMonth.Year, previousMonth.Month);
            return endDate.Day + daysInPreviousMonth - startDate.Day;
        }

        private static DateTime AddYearsClamped(DateTime date, int years) {
            return CreateClampedDate(date.Year + years, date.Month, date.Day);
        }

        private static DateTime CreateClampedDate(int year, int month, int day) {
            int clampedDay = Math.Min(day, DateTime.DaysInMonth(year, month));
            return new DateTime(year, month, clampedDay);
        }

        private bool TryEvaluateWorkday(string function, IReadOnlyList<string> tokens, out double result) {
            result = 0;
            int maxTokens = function == "WORKDAY.INTL" ? 4 : 3;
            if (tokens.Count < 2 || tokens.Count > maxTokens
                || !TryEvaluateFormulaOrNumeric(tokens[0], out double startSerial)
                || !TryGetWholeNumberArgument(tokens[1], out int days)
                || !TryGetDateFromSerial(startSerial, out DateTime current)) {
                return false;
            }

            bool[] weekendMask = DefaultWeekendMask();
            int holidayIndex = 2;
            if (function == "WORKDAY.INTL") {
                holidayIndex = 3;
                if (tokens.Count >= 3 && !TryResolveWeekendMask(tokens[2], weekendMask)) {
                    return false;
                }
            }

            var holidays = new HashSet<DateTime>();
            if (tokens.Count > holidayIndex && !TryResolveHolidayDates(tokens[holidayIndex], holidays)) {
                return false;
            }

            if (days == 0) {
                result = current.ToOADate();
                return true;
            }

            int direction = days > 0 ? 1 : -1;
            int remaining = Math.Abs(days);
            while (remaining > 0) {
                current = current.AddDays(direction);
                if (!IsMaskedWeekend(current.DayOfWeek, weekendMask) && !holidays.Contains(current.Date)) {
                    remaining--;
                }
            }

            result = current.ToOADate();
            return true;
        }

        private bool TryEvaluateDateTimeTextValue(string function, IReadOnlyList<string> tokens, out double result) {
            result = 0;
            if (tokens.Count != 1 || !TryResolveTextArgument(tokens[0], out string text)) {
                return false;
            }

            text = text.Trim();
            if (string.IsNullOrEmpty(text)) {
                return false;
            }

            if (function == "DATEVALUE") {
                if (!TryParseFormulaDateText(text, out DateTime date)) {
                    return false;
                }

                result = date.Date.ToOADate();
                return true;
            }

            if (!TryParseFormulaTimeText(text, out TimeSpan time)) {
                return false;
            }

            result = time.TotalDays;
            return true;
        }

        private static bool TryParseFormulaDateText(string text, out DateTime date) {
            string[] exactFormats = {
                "yyyy-MM-dd",
                "yyyy-M-d",
                "yyyy/MM/dd",
                "yyyy/M/d",
                "MM/dd/yyyy",
                "M/d/yyyy",
                "dd-MMM-yyyy",
                "d-MMM-yyyy",
                "MMM d yyyy",
                "MMMM d yyyy"
            };

            return DateTime.TryParseExact(text, exactFormats, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces, out date)
                || DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces, out date);
        }

        private static bool TryParseFormulaTimeText(string text, out TimeSpan time) {
            string[] exactFormats = {
                "H:mm",
                "HH:mm",
                "H:mm:ss",
                "HH:mm:ss",
                "h:mm tt",
                "hh:mm tt",
                "h:mm:ss tt",
                "hh:mm:ss tt"
            };

            if (DateTime.TryParseExact(text, exactFormats, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces, out DateTime exactTime)) {
                time = exactTime.TimeOfDay;
                return true;
            }

            if (DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces, out DateTime parsedTime)) {
                time = parsedTime.TimeOfDay;
                return true;
            }

            time = default;
            return false;
        }

        private bool TryResolveWeekendMask(string token, bool[] weekendMask) {
            string trimmed = token.Trim();
            if (TryResolveTextArgument(trimmed, out string maskText) && TryParseWeekendTextMask(maskText, weekendMask)) {
                return true;
            }

            if (!TryGetWholeNumberArgument(trimmed, out int weekendCode)) {
                return false;
            }

            return TryApplyWeekendCode(weekendCode, weekendMask);
        }

        private static bool[] DefaultWeekendMask() {
            var weekendMask = new bool[7];
            weekendMask[(int)DayOfWeek.Saturday] = true;
            weekendMask[(int)DayOfWeek.Sunday] = true;
            return weekendMask;
        }

        private static bool TryParseWeekendTextMask(string text, bool[] weekendMask) {
            if (text.Length != 7 || text.Any(ch => ch != '0' && ch != '1') || text.All(ch => ch == '1')) {
                return false;
            }

            Array.Clear(weekendMask, 0, weekendMask.Length);
            for (int index = 0; index < text.Length; index++) {
                DayOfWeek day = index == 6 ? DayOfWeek.Sunday : (DayOfWeek)(index + 1);
                weekendMask[(int)day] = text[index] == '1';
            }

            return true;
        }

        private static bool TryApplyWeekendCode(int weekendCode, bool[] weekendMask) {
            Array.Clear(weekendMask, 0, weekendMask.Length);
            switch (weekendCode) {
                case 1:
                    weekendMask[(int)DayOfWeek.Saturday] = true;
                    weekendMask[(int)DayOfWeek.Sunday] = true;
                    return true;
                case 2:
                    weekendMask[(int)DayOfWeek.Sunday] = true;
                    weekendMask[(int)DayOfWeek.Monday] = true;
                    return true;
                case 3:
                    weekendMask[(int)DayOfWeek.Monday] = true;
                    weekendMask[(int)DayOfWeek.Tuesday] = true;
                    return true;
                case 4:
                    weekendMask[(int)DayOfWeek.Tuesday] = true;
                    weekendMask[(int)DayOfWeek.Wednesday] = true;
                    return true;
                case 5:
                    weekendMask[(int)DayOfWeek.Wednesday] = true;
                    weekendMask[(int)DayOfWeek.Thursday] = true;
                    return true;
                case 6:
                    weekendMask[(int)DayOfWeek.Thursday] = true;
                    weekendMask[(int)DayOfWeek.Friday] = true;
                    return true;
                case 7:
                    weekendMask[(int)DayOfWeek.Friday] = true;
                    weekendMask[(int)DayOfWeek.Saturday] = true;
                    return true;
                default:
                    if (weekendCode >= 11 && weekendCode <= 17) {
                        DayOfWeek singleWeekendDay = weekendCode == 11 ? DayOfWeek.Sunday : (DayOfWeek)(weekendCode - 11);
                        weekendMask[(int)singleWeekendDay] = true;
                        return true;
                    }

                    return false;
            }
        }

        private static bool IsMaskedWeekend(DayOfWeek day, bool[] weekendMask) {
            return weekendMask[(int)day];
        }

    }
}
