using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private sealed class PivotDateGrouping {
            internal int SourceField;
            internal ExcelPivotGroupBy GroupBy;
            internal DateTime? Start, End;
            internal OpenXmlElement[] SavedItems = Array.Empty<OpenXmlElement>();
            internal PivotFieldValues Labels = null!;
            internal Dictionary<string, PivotFieldValue> LabelsByCaption = null!;

            internal PivotFieldValue Group(PivotFieldValue source) {
                if (source.Kind != PivotFieldValueKind.Date || source.Date == null)
                    throw new NotSupportedException("Date pivot grouping requires date source keys.");
                DateTime date = source.Date.Value;
                if (Start.HasValue && End.HasValue) {
                    int position = date < Start.Value ? 0 : date > End.Value ? Labels.Items.Count - 1
                        : GroupBy == ExcelPivotGroupBy.Years ? date.Year - Start.Value.Year + 1
                        : GroupBy == ExcelPivotGroupBy.Quarters ? (date.Month - 1) / 3 + 1
                        : date.Month;
                    if (position < 0 || position >= Labels.Items.Count)
                        throw new NotSupportedException("The date pivot group does not cover a source date.");
                    return Labels.Items[position];
                }
                string caption = FormatGeneratedDateGroupValue(date, GroupBy);
                return LabelsByCaption.TryGetValue(caption, out var label) ? label
                    : throw new NotSupportedException("The date pivot group items do not cover a source date.");
            }
        }

        private static PivotDateGrouping ReadPivotDateGrouping(CacheField[] fields, int index, int sourceFieldCount) {
            var field = fields[index];
            FieldGroup? group = field.FieldGroup;
            RangeProperties? range = group?.GetFirstChild<RangeProperties>();
            if (field.DatabaseField?.Value != false || field.Formula != null || group?.Base == null
                || group!.Base!.Value >= sourceFieldCount || group.ParentId != null && group.ParentId.Value >= fields.Length
                || range?.GroupBy == null || range.GroupBy.Value != GroupByValues.Years
                    && range.GroupBy.Value != GroupByValues.Months && range.GroupBy.Value != GroupByValues.Quarters
                || (range!.StartDate == null) != (range.EndDate == null))
                throw new NotSupportedException("This derived date pivot grouping is not qualified for materialization or lookup.");
            int source = (int)group.Base.Value;
            if (fields[source].DatabaseField?.Value == false || fields[source].Formula != null)
                throw new NotSupportedException("A derived date pivot group must reference a source date field.");
            OpenXmlElement[] items = group.GetFirstChild<GroupItems>()?.ChildElements.ToArray() ?? Array.Empty<OpenXmlElement>();
            if (items.Length == 0 || items.Length > 100_000 || items.Any(item => item is not StringItem))
                throw new NotSupportedException("The date pivot group requires saved text labels.");
            DateTime? start = range.StartDate?.Value, end = range.EndDate?.Value;
            var by = range.GroupBy.Value == GroupByValues.Years ? ExcelPivotGroupBy.Years
                : range.GroupBy.Value == GroupByValues.Quarters ? ExcelPivotGroupBy.Quarters : ExcelPivotGroupBy.Months;
            if (start.HasValue) {
                if (end < start) throw new NotSupportedException("The date pivot grouping bounds are invalid.");
                int expected = by == ExcelPivotGroupBy.Years ? end!.Value.Year - start.Value.Year + 3
                    : by == ExcelPivotGroupBy.Quarters ? 6 : 14;
                if (items.Length != expected)
                    throw new NotSupportedException("The date pivot group labels do not match its range boundaries.");
            }
            var labels = items.Cast<StringItem>().Select(item => PivotFieldValue.FromText(item.Val?.Value ?? string.Empty)).ToList();
            if (new HashSet<string>(labels.Select(item => item.Text), StringComparer.OrdinalIgnoreCase).Count != labels.Count)
                throw new NotSupportedException("The date pivot group labels are not distinct.");
            return new PivotDateGrouping { SourceField = source, GroupBy = by, Start = start, End = end,
                SavedItems = items, Labels = new PivotFieldValues(labels),
                LabelsByCaption = labels.ToDictionary(item => item.Text, StringComparer.OrdinalIgnoreCase) };
        }

        private static bool IsDerivedDateGroup(CacheField field) {
            var by = field.FieldGroup?.GetFirstChild<RangeProperties>()?.GroupBy?.Value;
            return field.DatabaseField?.Value == false && (by == GroupByValues.Years
                || by == GroupByValues.Quarters || by == GroupByValues.Months);
        }
    }
}
