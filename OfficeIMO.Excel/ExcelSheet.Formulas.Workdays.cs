namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private bool TryEvaluateNetworkDays(IReadOnlyList<string> tokens, out double result) {
            result = 0;
            if (tokens.Count < 2 || tokens.Count > 3
                || !TryEvaluateFormulaOrNumeric(tokens[0], out double startSerial)
                || !TryEvaluateFormulaOrNumeric(tokens[1], out double endSerial)
                || !TryGetDateFromSerial(startSerial, out DateTime startDate)
                || !TryGetDateFromSerial(endSerial, out DateTime endDate)) {
                return false;
            }

            var holidays = new HashSet<double>();
            if (tokens.Count == 3 && !TryResolveHolidayDates(tokens[2], holidays)) {
                return false;
            }

            int direction = Math.Floor(startSerial) <= Math.Floor(endSerial) ? 1 : -1;
            double current = Math.Floor(Math.Min(startSerial, endSerial));
            double last = Math.Floor(Math.Max(startSerial, endSerial));
            int days = 0;
            while (current <= last) {
                DayOfWeek weekday = GetFormulaWeekday(current);
                if (weekday != DayOfWeek.Saturday && weekday != DayOfWeek.Sunday
                    && !holidays.Contains(current)) {
                    days++;
                }

                current++;
            }

            result = days * direction;
            return true;
        }

        private bool TryResolveHolidayDates(string token, HashSet<double> holidays) {
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
                    holidays.Add(Math.Floor(value.Number.Value));
                }
            }

            return true;
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

            var holidays = new HashSet<double>();
            if (tokens.Count > holidayIndex && !TryResolveHolidayDates(tokens[holidayIndex], holidays)) {
                return false;
            }

            double currentSerial = Math.Floor(startSerial);
            if (days == 0) {
                result = currentSerial;
                return true;
            }

            int direction = days > 0 ? 1 : -1;
            int remaining = Math.Abs(days);
            while (remaining > 0) {
                currentSerial += direction;
                if (!TryGetDateFromSerial(currentSerial, out _)) return false;
                if (!IsMaskedWeekend(GetFormulaWeekday(currentSerial), weekendMask) && !holidays.Contains(currentSerial)) {
                    remaining--;
                }
            }

            result = currentSerial;
            return true;
        }

    }
}
