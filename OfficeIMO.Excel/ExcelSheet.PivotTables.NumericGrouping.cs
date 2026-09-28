using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private sealed class PivotNumericGrouping {
            internal int Field;
            internal double Start, End, Interval;
            internal OpenXmlElement[] SavedItems = Array.Empty<OpenXmlElement>();
            internal PivotFieldValues Labels = null!;
            internal ExcelPivotGrouping SourceGrouping = null!;

            internal PivotFieldValue Group(PivotFieldValue raw) {
                if (raw.Kind != PivotFieldValueKind.Number || raw.Number == null)
                    throw new NotSupportedException("Numeric pivot grouping requires numeric source keys.");
                double value = raw.Number.Value;
                int index = value < Start ? 0 : value > End ? Labels.Items.Count - 1
                    : Math.Min((int)Math.Floor((value - Start) / Interval) + 1, Labels.Items.Count - 2);
                return Labels.Items[index];
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
            if (!IsSupportedIntegerGroupNumber(start) || !IsSupportedIntegerGroupNumber(end)
                || !IsSupportedIntegerGroupNumber(interval) || interval < 1 || end <= start)
                throw new NotSupportedException("The numeric pivot grouping bounds are invalid.");
            double span = Math.Ceiling((end - start) / interval);
            if (span < 1 || span > 99_998)
                throw new NotSupportedException("The numeric pivot grouping exceeds the supported item budget.");
            OpenXmlElement[] items = group.GetFirstChild<GroupItems>()?.ChildElements.ToArray()
                ?? Array.Empty<OpenXmlElement>();
            if (items.Length != (int)span + 2 || items.Any(item => item is not StringItem))
                throw new NotSupportedException("The numeric pivot group labels do not match its range boundaries.");
            var labels = items.Cast<StringItem>().Select(item => PivotFieldValue.FromText(item.Val?.Value ?? string.Empty)).ToList();
            if (labels.Distinct().Count() != labels.Count)
                throw new NotSupportedException("The numeric pivot group labels are not distinct.");
            return new PivotNumericGrouping {
                Field = index, Start = start, End = end, Interval = interval,
                SavedItems = items, Labels = new PivotFieldValues(labels),
                SourceGrouping = ExcelPivotGrouping.Number(field.Name?.Value ?? string.Empty, interval, start, end)
            };
        }

        private static PivotFieldValues? BuildAuthorNumericGroupLabels(ExcelPivotGrouping grouping) {
            if (grouping.GroupBy != ExcelPivotGroupBy.Range || grouping.StartNumber is not double start
                || grouping.EndNumber is not double end || grouping.Interval is not double interval
                || !IsSupportedIntegerGroupNumber(start) || !IsSupportedIntegerGroupNumber(end)
                || !IsSupportedIntegerGroupNumber(interval) || interval < 1 || end <= start)
                return null;
            double span = Math.Ceiling((end - start) / interval);
            if (span < 1 || span > 99_998) return null;
            var labels = new List<PivotFieldValue>((int)span + 2) {
                PivotFieldValue.FromText("<" + start.ToString(CultureInfo.InvariantCulture))
            };
            for (int index = 0; index < (int)span; index++) {
                double lower = start + index * interval;
                double upper = index == (int)span - 1 ? end : lower + interval - 1;
                labels.Add(PivotFieldValue.FromText(lower.ToString(CultureInfo.InvariantCulture)
                    + "-" + upper.ToString(CultureInfo.InvariantCulture)));
            }
            labels.Add(PivotFieldValue.FromText(">" + end.ToString(CultureInfo.InvariantCulture)));
            return new PivotFieldValues(labels);
        }

        private static bool IsSupportedIntegerGroupNumber(double number) => !double.IsNaN(number)
            && !double.IsInfinity(number) && Math.Abs(number) <= 9_000_000_000_000_000d
            && number == Math.Truncate(number);

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
