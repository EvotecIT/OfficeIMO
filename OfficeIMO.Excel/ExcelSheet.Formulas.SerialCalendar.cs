namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private static double GetFormulaTimePart(double serial, double secondsPerPart, double modulus) {
            // Excel extracts time parts after rounding to the nearest whole second.
            double seconds = Math.Round((serial - Math.Floor(serial)) * 86400d, 0, MidpointRounding.AwayFromZero);
            return Math.Floor(seconds / secondsPerPart) % modulus;
        }

        private double GetFormulaYearFracEuropean(DateTime startDate, DateTime endDate) {
            // Excel's producer corpus qualifies February-end normalization for this
            // basis in 1904 workbooks; the 1900 system retains the calendar day.
            bool normalizeFebruary = _excelDocument.DateSystem == ExcelDateSystem.NineteenFour;
            int startDay = normalizeFebruary && IsLastDayOfFebruary(startDate) ? 30 : Math.Min(startDate.Day, 30);
            int endDay = normalizeFebruary && IsLastDayOfFebruary(endDate) ? 30 : Math.Min(endDate.Day, 30);
            return (endDate.Year - startDate.Year) * 360 + (endDate.Month - startDate.Month) * 30 + endDay - startDay;
        }

        private bool IsFictitiousLeapDay(double serial) =>
            _excelDocument.DateSystem == ExcelDateSystem.NineteenHundred && Math.Floor(serial) == 60d;

        private int GetFormulaCalendarDay(double serial, DateTime date) => IsFictitiousLeapDay(serial) ? 29 : date.Day;

        private int GetFormulaMonthLength(int year, int month) =>
            _excelDocument.DateSystem == ExcelDateSystem.NineteenHundred && year == 1900 && month == 2
                ? 29 : DateTime.DaysInMonth(year, month);

        private DayOfWeek GetFormulaWeekday(double serial) {
            if (_excelDocument.DateSystem == ExcelDateSystem.NineteenHundred && serial >= 0d && serial < 61d)
                return (DayOfWeek)((Math.Floor(serial) + 6d) % 7d);
            return ExcelDateSystemConverter.FromSerial(serial, _excelDocument.DateSystem).DayOfWeek;
        }

        private double GetFormulaMonthShift(double serial, DateTime date, int months, bool endOfMonth) {
            DateTime monthStart = new DateTime(date.Year, date.Month, 1).AddMonths(months);
            int day = endOfMonth ? DateTime.DaysInMonth(monthStart.Year, monthStart.Month)
                : Math.Min(GetFormulaCalendarDay(serial, date), DateTime.DaysInMonth(monthStart.Year, monthStart.Month));
            return ToExcelDateSerial(monthStart) + day - 1;
        }

        private double GetFormulaDays360(double startSerial, DateTime startDate, double endSerial, DateTime endDate, bool european) {
            int startDay = GetFormulaCalendarDay(startSerial, startDate);
            int endDay = GetFormulaCalendarDay(endSerial, endDate);
            if (european) {
                startDay = Math.Min(startDay, 30);
                endDay = Math.Min(endDay, 30);
            } else {
                if (startDay == 31 || (startDate.Month == 2 && startDay == GetFormulaMonthLength(startDate.Year, 2))) startDay = 30;
                if (endDay == 31 && startDay >= 30) endDay = 30;
            }
            return (endDate.Year - startDate.Year) * 360 + (endDate.Month - startDate.Month) * 30 + endDay - startDay;
        }

        private double GetFormulaRemainingDays(double startSerial, DateTime startDate, double endSerial, DateTime endDate) {
            int startDay = GetFormulaCalendarDay(startSerial, startDate);
            int endDay = GetFormulaCalendarDay(endSerial, endDate);
            if (endDay >= startDay) return endDay - startDay;
            DateTime previousMonth = endDate.AddMonths(-1);
            return endDay + GetFormulaMonthLength(previousMonth.Year, previousMonth.Month) - startDay;
        }

        private int GetFormulaCompletedYears(double startSerial, DateTime startDate, double endSerial, DateTime endDate) {
            int years = endDate.Year - startDate.Year;
            if (endDate.Month < startDate.Month || (endDate.Month == startDate.Month
                && GetFormulaCalendarDay(endSerial, endDate) < GetFormulaCalendarDay(startSerial, startDate))) years--;
            return years;
        }

        private int GetFormulaCompletedMonths(double startSerial, DateTime startDate, double endSerial, DateTime endDate) {
            int months = (endDate.Year - startDate.Year) * 12 + endDate.Month - startDate.Month;
            if (GetFormulaCalendarDay(endSerial, endDate) < GetFormulaCalendarDay(startSerial, startDate)) months--;
            return months;
        }

        private double GetFormulaAnniversaryDays(double startSerial, DateTime startDate, double endSerial, DateTime endDate) {
            double anniversarySerial = GetFormulaAnniversarySerial(startSerial, startDate, endDate.Year);
            if (endDate.Month < startDate.Month || (endDate.Month == startDate.Month
                && GetFormulaCalendarDay(endSerial, endDate) < GetFormulaCalendarDay(startSerial, startDate))) {
                anniversarySerial = GetFormulaAnniversarySerial(startSerial, startDate, endDate.Year - 1);
            }
            return Math.Floor(endSerial) - anniversarySerial;
        }

        private double GetFormulaAnniversarySerial(double startSerial, DateTime startDate, int year) {
            int day = Math.Min(GetFormulaCalendarDay(startSerial, startDate), GetFormulaMonthLength(year, startDate.Month));
            return ToExcelDateSerial(new DateTime(year, startDate.Month, 1)) + day - 1;
        }

        private bool TryGetFormulaWeekNumber(double serial, DateTime date, DayOfWeek weekStart, bool isoSystem, out double result) {
            if (isoSystem) {
                result = GetIsoWeekNumber(date);
                return true;
            }
            double firstDay = ToExcelDateSerial(new DateTime(date.Year, 1, 1));
            double firstWeekStart = firstDay - GetDayOffset(GetFormulaWeekday(firstDay), weekStart);
            result = Math.Floor((Math.Floor(serial) - firstWeekStart) / 7d) + 1;
            return result >= 1d && result <= 54d;
        }
    }
}
