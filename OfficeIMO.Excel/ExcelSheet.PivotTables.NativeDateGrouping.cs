using System.Globalization;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private static void PrepareNativeDateGroupingFields(
            IReadOnlyList<GeneratedPivotGroupingField> generatedFields,
            IReadOnlyList<PivotFieldValues> sourceValues) {
            foreach (var fields in generatedFields.GroupBy(field => field.SourceIndex)) {
                // The native range and group-item contract is currently qualified for these levels.
                if (fields.Any(field => field.GroupBy != ExcelPivotGroupBy.Years
                    && field.GroupBy != ExcelPivotGroupBy.Quarters
                    && field.GroupBy != ExcelPivotGroupBy.Months))
                    continue;

                var dates = sourceValues[fields.Key].Items;
                if (dates.Count == 0) continue;
                if (dates.Any(value => value.Kind != PivotFieldValueKind.Date || value.Date == null))
                    throw new NotSupportedException("A native date pivot hierarchy requires date values in every source row.");

                DateTime minimum = dates.Min(value => value.Date!.Value);
                DateTime maximum = dates.Max(value => value.Date!.Value);
                var first = fields.First();
                DateTime start = first.Grouping.StartDate ?? minimum.Date;
                DateTime end = first.Grouping.EndDate ?? (maximum.Date < DateTime.MaxValue.Date
                    ? maximum.Date.AddDays(1) : maximum);
                if (end < start || end.Year - start.Year > 99_997)
                    throw new NotSupportedException("The native date pivot hierarchy has invalid or excessive date bounds.");
                foreach (var field in fields)
                    field.NativeGrouping = new NativePivotDateGrouping(field.FieldName, field.GroupBy, start, end);
            }
        }

        private sealed class NativePivotDateGrouping {
            public NativePivotDateGrouping(string fieldName, ExcelPivotGroupBy groupBy, DateTime start, DateTime end) {
                GroupBy = groupBy;
                Start = start;
                End = end;
                Grouping = ExcelPivotGrouping.Date(fieldName, groupBy, start, end);
                var labels = new List<PivotFieldValue> { PivotFieldValue.FromText("<" + start.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)) };
                if (groupBy == ExcelPivotGroupBy.Years) {
                    for (int year = start.Year; year <= end.Year; year++)
                        labels.Add(PivotFieldValue.FromText(year.ToString(CultureInfo.InvariantCulture)));
                } else if (groupBy == ExcelPivotGroupBy.Quarters) {
                    for (int quarter = 1; quarter <= 4; quarter++)
                        labels.Add(PivotFieldValue.FromText("Q" + quarter.ToString(CultureInfo.InvariantCulture)));
                } else {
                    for (int month = 1; month <= 12; month++)
                        labels.Add(PivotFieldValue.FromText(CultureInfo.InvariantCulture.DateTimeFormat.GetMonthName(month)));
                }
                labels.Add(PivotFieldValue.FromText(">" + end.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)));
                Labels = new PivotFieldValues(labels);
            }

            public ExcelPivotGroupBy GroupBy { get; }
            public DateTime Start { get; }
            public DateTime End { get; }
            public ExcelPivotGrouping Grouping { get; }
            public PivotFieldValues Labels { get; }

            public PivotFieldValue Group(PivotFieldValue source) {
                if (source.Kind != PivotFieldValueKind.Date || source.Date == null)
                    throw new NotSupportedException("A native date pivot hierarchy requires date source keys.");
                DateTime date = source.Date.Value;
                int position = date < Start ? 0 : date > End ? Labels.Items.Count - 1
                    : GroupBy == ExcelPivotGroupBy.Years ? date.Year - Start.Year + 1
                    : GroupBy == ExcelPivotGroupBy.Quarters ? (date.Month - 1) / 3 + 1
                    : date.Month;
                return Labels.Items[position];
            }
        }
    }
}
