using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private sealed class PivotNumericGrouping {
            internal double Start, End;
            internal decimal DecimalStart, DecimalInterval;
            internal OpenXmlElement[] SavedItems = Array.Empty<OpenXmlElement>();
            internal PivotFieldValues Labels = null!;
            internal ExcelPivotGrouping SourceGrouping = null!;

            internal PivotFieldValue Group(PivotFieldValue raw) {
                if (raw.Kind != PivotFieldValueKind.Number || raw.Number == null)
                    throw new NotSupportedException("Numeric pivot grouping requires numeric source keys.");
                double value = raw.Number.Value;
                if (double.IsNaN(value) || double.IsInfinity(value))
                    throw new NotSupportedException("Numeric pivot grouping requires finite source keys.");
                int index = value < Start ? 0 : value > End ? Labels.Items.Count - 1
                    : Math.Min((int)decimal.Floor(((decimal)value - DecimalStart) / DecimalInterval) + 1,
                        Labels.Items.Count - 2);
                return Labels.Items[index];
            }

            internal bool TryResolveBoundary(object? criterion, out string? label) {
                label = null;
                if (criterion is not IConvertible || criterion is string || criterion is bool || criterion is DateTime)
                    return false;
                double number;
                try { number = Convert.ToDouble(criterion, CultureInfo.InvariantCulture); }
                catch (Exception exception) when (exception is FormatException || exception is InvalidCastException || exception is OverflowException) {
                    return false;
                }
                if (!IsSupportedGroupNumber(number) || number < Start || number >= End) return false;
                decimal offset = ((decimal)number - DecimalStart) / DecimalInterval;
                if (offset != decimal.Truncate(offset) || offset < 0 || offset >= Labels.Items.Count - 2)
                    return false;
                label = Labels.Items[(int)offset + 1].Text;
                return true;
            }
        }

        private static PivotNumericGrouping? ReadPivotNumericGrouping(CacheField field, int index) {
            FieldGroup? group = field.FieldGroup;
            if (group == null) return null;
            RangeProperties? range = group.GetFirstChild<RangeProperties>();
            if (group.Base?.Value != index || group.ParentId != null || range == null
                || range.GroupBy?.Value is GroupByValues groupBy && groupBy != GroupByValues.Range
                || range.StartNumber == null || range.EndNum == null || range.GroupInterval == null)
                throw new NotSupportedException("This grouped pivot profile is not qualified for numeric materialization or lookup.");
            double start = range.StartNumber.Value, end = range.EndNum.Value, interval = range.GroupInterval.Value;
            if (!TryGetNumericGroupSpan(start, end, interval, out int span))
                throw new NotSupportedException("The numeric pivot grouping bounds are invalid.");
            OpenXmlElement[] items = group.GetFirstChild<GroupItems>()?.ChildElements.ToArray()
                ?? Array.Empty<OpenXmlElement>();
            if (items.Length != span + 2 || items.Any(item => item is not StringItem))
                throw new NotSupportedException("The numeric pivot group labels do not match its range boundaries.");
            var labels = items.Cast<StringItem>().Select(item => PivotFieldValue.FromText(item.Val?.Value ?? string.Empty)).ToList();
            if (labels.Distinct().Count() != labels.Count)
                throw new NotSupportedException("The numeric pivot group labels are not distinct.");
            return new PivotNumericGrouping {
                Start = start, End = end,
                DecimalStart = (decimal)start, DecimalInterval = (decimal)interval,
                SavedItems = items, Labels = new PivotFieldValues(labels),
                SourceGrouping = ExcelPivotGrouping.Number(field.Name?.Value ?? string.Empty, interval, start, end)
            };
        }

        private static PivotFieldValues? BuildAuthorNumericGroupLabels(ExcelPivotGrouping grouping) {
            if (grouping.GroupBy != ExcelPivotGroupBy.Range || grouping.StartNumber is not double start
                || grouping.EndNumber is not double end || grouping.Interval is not double interval
                || !TryGetNumericGroupSpan(start, end, interval, out int span))
                return null;
            decimal decimalStart = (decimal)start, decimalEnd = (decimal)end, decimalInterval = (decimal)interval;
            bool integerRange = start == Math.Truncate(start) && end == Math.Truncate(end)
                && interval == Math.Truncate(interval);
            var labels = new List<PivotFieldValue>(span + 2) {
                PivotFieldValue.FromText("<" + start.ToString(CultureInfo.InvariantCulture))
            };
            for (int index = 0; index < span; index++) {
                decimal lower = decimalStart + index * decimalInterval;
                decimal upper = index == span - 1 ? decimalEnd
                    : lower + decimalInterval - (integerRange ? 1 : 0);
                labels.Add(PivotFieldValue.FromText(lower.ToString(CultureInfo.InvariantCulture)
                    + "-" + upper.ToString(CultureInfo.InvariantCulture)));
            }
            labels.Add(PivotFieldValue.FromText(">" + end.ToString(CultureInfo.InvariantCulture)));
            return new PivotFieldValues(labels);
        }

        private static bool IsSupportedGroupNumber(double number) => !double.IsNaN(number)
            && !double.IsInfinity(number) && Math.Abs(number) <= 9_000_000_000_000_000d;

        private static bool TryGetNumericGroupSpan(double start, double end, double interval, out int span) {
            span = 0;
            if (!IsSupportedGroupNumber(start) || !IsSupportedGroupNumber(end)
                || !IsSupportedGroupNumber(interval) || interval <= 0 || end <= start)
                return false;
            double approximateSpan = (end - start) / interval;
            if (double.IsInfinity(approximateSpan) || approximateSpan > 100_000) return false;
            decimal decimalInterval = (decimal)interval;
            if (decimalInterval <= 0) return false;
            decimal count = decimal.Ceiling(((decimal)end - (decimal)start) / decimalInterval);
            if (count < 1 || count > 99_998) return false;
            span = (int)count;
            return true;
        }

        private static PivotFieldValue MaterializedPivotAxisKey(ExcelSheet source, int row, int column, int field,
            IReadOnlyDictionary<int, PivotNumericGrouping> groupings,
            IReadOnlyDictionary<int, PivotDateGrouping> dateGroupings) {
            if (dateGroupings.TryGetValue(field, out var dateGrouping))
                return dateGrouping.Group(source.GetPivotFieldValue(row, column, null));
            if (groupings.TryGetValue(field, out var grouping))
                return grouping.Group(source.GetPivotFieldValue(row, column, grouping.SourceGrouping));
            return source.GetPivotFieldValue(row, column, null);
        }
    }
}
