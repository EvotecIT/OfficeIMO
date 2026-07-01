using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Excel {
    internal static partial class ExcelConditionalVisualEvaluator {
        private static bool IsTimePeriodRule(ExcelConditionalFormattingInfo rule) =>
            string.Equals(rule.Type, "TimePeriod", StringComparison.OrdinalIgnoreCase);

        private static bool CanEvaluateTimePeriodRule(ExcelSheet sheet, IReadOnlyList<ExcelVisualCell> cells, ExcelConditionalFormattingInfo rule, DateTime referenceDate) =>
            !string.IsNullOrWhiteSpace(rule.TimePeriod) &&
            TryGetTimePeriodBounds(rule.TimePeriod!, referenceDate, out _, out _) &&
            GetRuleCells(cells, rule.Range).Any(cell => TryGetCellDate(sheet, cell, out _));

        private static void ApplyTimePeriodFormat(
            ExcelSheet sheet,
            IReadOnlyList<ExcelVisualCell> cells,
            ExcelConditionalFormattingInfo rule,
            DateTime referenceDate,
            Dictionary<string, ExcelConditionalCellFormat> formats,
            HashSet<string> stoppedCells) {
            if (string.IsNullOrWhiteSpace(rule.TimePeriod) ||
                !TryGetTimePeriodBounds(rule.TimePeriod!, referenceDate, out DateTime start, out DateTime endExclusive)) {
                return;
            }

            foreach (ExcelVisualCell cell in GetRuleCells(cells, rule.Range)) {
                string key = Key(cell.Row, cell.Column);
                if (stoppedCells.Contains(key) ||
                    !TryGetCellDate(sheet, cell, out DateTime cellDate) ||
                    cellDate < start ||
                    cellDate >= endExclusive) {
                    continue;
                }

                ApplyDifferentialFormat(rule, key, formats);

                if (rule.StopIfTrue) {
                    stoppedCells.Add(key);
                }
            }
        }

        private static bool TryGetCellDate(ExcelSheet sheet, ExcelVisualCell cell, out DateTime date) {
            date = default;
            if (!TryGetCellNumericValue(sheet, cell, out double serial)) {
                return false;
            }

            try {
                date = ExcelDateSystemConverter.FromSerial(Math.Floor(serial), sheet.Document.DateSystem).Date;
                return true;
            } catch (ArgumentException) {
                return false;
            }
        }

        private static bool TryGetTimePeriodBounds(string timePeriod, DateTime referenceDate, out DateTime start, out DateTime endExclusive) {
            DateTime today = referenceDate.Date;
            if (string.Equals(timePeriod, "Yesterday", StringComparison.OrdinalIgnoreCase)) {
                start = today.AddDays(-1);
                endExclusive = today;
                return true;
            }

            if (string.Equals(timePeriod, "Today", StringComparison.OrdinalIgnoreCase)) {
                start = today;
                endExclusive = today.AddDays(1);
                return true;
            }

            if (string.Equals(timePeriod, "Tomorrow", StringComparison.OrdinalIgnoreCase)) {
                start = today.AddDays(1);
                endExclusive = today.AddDays(2);
                return true;
            }

            if (string.Equals(timePeriod, "Last7Days", StringComparison.OrdinalIgnoreCase)) {
                start = today.AddDays(-6);
                endExclusive = today.AddDays(1);
                return true;
            }

            DateTime thisWeekStart = today.AddDays(-(((int)today.DayOfWeek + 6) % 7));
            if (string.Equals(timePeriod, "LastWeek", StringComparison.OrdinalIgnoreCase)) {
                start = thisWeekStart.AddDays(-7);
                endExclusive = thisWeekStart;
                return true;
            }

            if (string.Equals(timePeriod, "ThisWeek", StringComparison.OrdinalIgnoreCase)) {
                start = thisWeekStart;
                endExclusive = thisWeekStart.AddDays(7);
                return true;
            }

            if (string.Equals(timePeriod, "NextWeek", StringComparison.OrdinalIgnoreCase)) {
                start = thisWeekStart.AddDays(7);
                endExclusive = thisWeekStart.AddDays(14);
                return true;
            }

            DateTime thisMonthStart = new DateTime(today.Year, today.Month, 1);
            if (string.Equals(timePeriod, "LastMonth", StringComparison.OrdinalIgnoreCase)) {
                start = thisMonthStart.AddMonths(-1);
                endExclusive = thisMonthStart;
                return true;
            }

            if (string.Equals(timePeriod, "ThisMonth", StringComparison.OrdinalIgnoreCase)) {
                start = thisMonthStart;
                endExclusive = thisMonthStart.AddMonths(1);
                return true;
            }

            if (string.Equals(timePeriod, "NextMonth", StringComparison.OrdinalIgnoreCase)) {
                start = thisMonthStart.AddMonths(1);
                endExclusive = thisMonthStart.AddMonths(2);
                return true;
            }

            start = default;
            endExclusive = default;
            return false;
        }
    }
}
