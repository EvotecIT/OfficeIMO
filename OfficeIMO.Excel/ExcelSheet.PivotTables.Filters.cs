using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Text;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {

        private static PageField CreatePageField(int fieldIndex, ExcelPivotFieldOptions? options, IReadOnlyList<string> values) {
            var pageField = new PageField { Field = fieldIndex };
            if (options == null || string.IsNullOrWhiteSpace(options.SelectedItem)) {
                return pageField;
            }

            int selectedIndex = FindPivotItemIndex(options.SelectedItem!, values, options.FieldName, nameof(options.SelectedItem));
            pageField.Item = (uint)selectedIndex;
            return pageField;
        }

        private static void ApplyPivotFieldOptions(PivotField pivotField, ExcelPivotFieldOptions? options, WorkbookPart workbookPart,
            IReadOnlyList<string> values) {
            if (options == null) return;

            if (options.SortType.HasValue) pivotField.SortType = options.SortType.Value;
            if (options.DefaultSubtotal.HasValue) pivotField.DefaultSubtotal = options.DefaultSubtotal.Value;
            if (options.SubtotalTop.HasValue) pivotField.SubtotalTop = options.SubtotalTop.Value;
            if (options.InsertBlankRow.HasValue) pivotField.InsertBlankRow = options.InsertBlankRow.Value;
            if (options.InsertPageBreak.HasValue) pivotField.InsertPageBreak = options.InsertPageBreak.Value;
            if (options.Compact.HasValue) pivotField.Compact = options.Compact.Value;
            if (options.Outline.HasValue) pivotField.Outline = options.Outline.Value;
            if (options.ShowDropDowns.HasValue) pivotField.ShowDropDowns = options.ShowDropDowns.Value;
            if (options.MultipleItemSelectionAllowed.HasValue) pivotField.MultipleItemSelectionAllowed = options.MultipleItemSelectionAllowed.Value;
            if (options.IncludeNewItemsInFilter.HasValue) pivotField.IncludeNewItemsInFilter = options.IncludeNewItemsInFilter.Value;
            if (!string.IsNullOrWhiteSpace(options.SubtotalCaption)) pivotField.SubtotalCaption = options.SubtotalCaption;

            uint? numberFormatId = ResolveNumberFormatId(workbookPart, options.NumberFormatId, options.NumberFormat);
            if (numberFormatId.HasValue) pivotField.NumberFormatId = numberFormatId.Value;

            ApplyPivotFieldItemFilters(pivotField, options, values);
        }

        private static void ApplyPivotFieldItemFilters(PivotField pivotField, ExcelPivotFieldOptions options, IReadOnlyList<string> values) {
            if (options.HiddenItems.Count == 0 && options.VisibleItems.Count == 0) return;
            if (values.Count == 0) {
                throw new ArgumentException($"Field '{options.FieldName}' has no cache items to filter.", nameof(options));
            }

            var hidden = new HashSet<int>();
            if (options.HiddenItems.Count > 0) {
                foreach (string item in options.HiddenItems) {
                    hidden.Add(FindPivotItemIndex(item, values, options.FieldName, nameof(options.HiddenItems)));
                }
            } else {
                var visible = new HashSet<int>();
                foreach (string item in options.VisibleItems) {
                    visible.Add(FindPivotItemIndex(item, values, options.FieldName, nameof(options.VisibleItems)));
                }

                for (int i = 0; i < values.Count; i++) {
                    if (!visible.Contains(i)) hidden.Add(i);
                }
            }

            var items = new Items { Count = (uint)(values.Count + 1) };
            for (int i = 0; i < values.Count; i++) {
                var item = new Item { Index = (uint)i };
                if (hidden.Contains(i)) item.Hidden = true;
                items.Append(item);
            }
            items.Append(new Item { ItemType = ItemValues.Default });

            pivotField.Items = items;
            if (options.ShowAll == null) {
                pivotField.ShowAll = false;
            }
        }

        private static PivotFilters? CreatePivotFilters(IReadOnlyList<ExcelPivotFilter> filters,
            IDictionary<string, int> headerIndex,
            IReadOnlyList<ExcelPivotDataField> dataFields) {
            if (filters.Count == 0) return null;

            var pivotFilters = new PivotFilters { Count = (uint)filters.Count };
            for (int i = 0; i < filters.Count; i++) {
                var filter = filters[i];
                int fieldIndex = ResolveFieldIndex(filter.FieldName, headerIndex, nameof(filters));
                var pivotFilter = new PivotFilter {
                    Field = (uint)fieldIndex,
                    Type = filter.Type,
                    EvaluationOrder = i,
                    Id = (uint)(i + 1)
                };

                if (!string.IsNullOrWhiteSpace(filter.Value1)) pivotFilter.StringValue1 = filter.Value1;
                if (!string.IsNullOrWhiteSpace(filter.Value2)) pivotFilter.StringValue2 = filter.Value2;
                if (!string.IsNullOrWhiteSpace(filter.Name)) pivotFilter.Name = filter.Name;
                if (!string.IsNullOrWhiteSpace(filter.Description)) pivotFilter.Description = filter.Description;
                if (!string.IsNullOrWhiteSpace(filter.DataFieldName)) {
                    pivotFilter.MeasureField = (uint)ResolvePivotDataFieldIndex(filter.DataFieldName!, dataFields);
                }

                pivotFilter.AutoFilter = CreatePivotFilterAutoFilter(filter);
                pivotFilters.Append(pivotFilter);
            }

            return pivotFilters;
        }

        private static AutoFilter CreatePivotFilterAutoFilter(ExcelPivotFilter filter) {
            var autoFilter = new AutoFilter { Reference = "A1" };
            var filterColumn = new FilterColumn { ColumnId = 0U };

            if (TryCreateTop10Filter(filter, out var top10)) {
                filterColumn.Append(top10);
                autoFilter.Append(filterColumn);
                return autoFilter;
            }

            if (TryCreateDynamicFilter(filter, out var dynamicFilter)) {
                filterColumn.Append(dynamicFilter);
                autoFilter.Append(filterColumn);
                return autoFilter;
            }

            CustomFilters customFilters;

            if (TryResolveBetweenFilter(filter.Type, out var firstOperator, out var secondOperator, out bool matchAll)) {
                if (filter.Value1 == null || filter.Value2 == null) {
                    throw new ArgumentException($"Pivot filter '{filter.Type}' requires two values.", nameof(filter));
                }

                customFilters = new CustomFilters { And = matchAll };
                customFilters.Append(new CustomFilter {
                    Operator = firstOperator,
                    Val = filter.Value1
                });
                customFilters.Append(new CustomFilter {
                    Operator = secondOperator,
                    Val = filter.Value2
                });
            } else {
                if (filter.Value1 == null) {
                    throw new ArgumentException($"Pivot filter '{filter.Type}' requires a value.", nameof(filter));
                }

                customFilters = new CustomFilters();
                customFilters.Append(new CustomFilter {
                    Operator = ResolveSingleFilterOperator(filter.Type),
                    Val = NormalizePivotFilterAutoFilterValue(filter.Type, filter.Value1)
                });
            }

            filterColumn.Append(customFilters);
            autoFilter.Append(filterColumn);
            return autoFilter;
        }

        private static bool TryCreateTop10Filter(ExcelPivotFilter filter, out Top10 top10) {
            if (filter.Type != PivotFilterValues.Count && filter.Type != PivotFilterValues.Percent && filter.Type != PivotFilterValues.Sum) {
                top10 = new Top10();
                return false;
            }

            if (filter.Value1 == null) {
                throw new ArgumentException($"Pivot filter '{filter.Type}' requires a value.", nameof(filter));
            }

            if (!double.TryParse(filter.Value1, NumberStyles.Float, CultureInfo.InvariantCulture, out double value)) {
                throw new ArgumentException($"Pivot filter '{filter.Type}' value '{filter.Value1}' is not numeric.", nameof(filter));
            }

            top10 = new Top10 {
                Top = filter.IsTop ?? true,
                Percent = filter.IsPercent ?? filter.Type == PivotFilterValues.Percent,
                Val = value
            };

            if (!string.IsNullOrWhiteSpace(filter.FilterValue)) {
                if (!double.TryParse(filter.FilterValue, NumberStyles.Float, CultureInfo.InvariantCulture, out double filterValue)) {
                    throw new ArgumentException($"Pivot filter '{filter.Type}' filter value '{filter.FilterValue}' is not numeric.", nameof(filter));
                }

                top10.FilterValue = filterValue;
            }

            return true;
        }

        private static bool TryCreateDynamicFilter(ExcelPivotFilter filter, out DynamicFilter dynamicFilter) {
            DynamicFilterValues? dynamicType = ResolveDynamicFilterType(filter.Type);
            if (!dynamicType.HasValue) {
                dynamicFilter = new DynamicFilter();
                return false;
            }

            dynamicFilter = new DynamicFilter { Type = dynamicType.Value };
            return true;
        }

        private static DynamicFilterValues? ResolveDynamicFilterType(PivotFilterValues type) {
            if (type == PivotFilterValues.Today) return DynamicFilterValues.Today;
            if (type == PivotFilterValues.Yesterday) return DynamicFilterValues.Yesterday;
            if (type == PivotFilterValues.Tomorrow) return DynamicFilterValues.Tomorrow;
            if (type == PivotFilterValues.ThisWeek) return DynamicFilterValues.ThisWeek;
            if (type == PivotFilterValues.LastWeek) return DynamicFilterValues.LastWeek;
            if (type == PivotFilterValues.NextWeek) return DynamicFilterValues.NextWeek;
            if (type == PivotFilterValues.ThisMonth) return DynamicFilterValues.ThisMonth;
            if (type == PivotFilterValues.LastMonth) return DynamicFilterValues.LastMonth;
            if (type == PivotFilterValues.NextMonth) return DynamicFilterValues.NextMonth;
            if (type == PivotFilterValues.ThisQuarter) return DynamicFilterValues.ThisQuarter;
            if (type == PivotFilterValues.LastQuarter) return DynamicFilterValues.LastQuarter;
            if (type == PivotFilterValues.NextQuarter) return DynamicFilterValues.NextQuarter;
            if (type == PivotFilterValues.ThisYear) return DynamicFilterValues.ThisYear;
            if (type == PivotFilterValues.LastYear) return DynamicFilterValues.LastYear;
            if (type == PivotFilterValues.NextYear) return DynamicFilterValues.NextYear;
            if (type == PivotFilterValues.YearToDate) return DynamicFilterValues.YearToDate;
            if (type == PivotFilterValues.January) return DynamicFilterValues.January;
            if (type == PivotFilterValues.February) return DynamicFilterValues.February;
            if (type == PivotFilterValues.March) return DynamicFilterValues.March;
            if (type == PivotFilterValues.April) return DynamicFilterValues.April;
            if (type == PivotFilterValues.May) return DynamicFilterValues.May;
            if (type == PivotFilterValues.June) return DynamicFilterValues.June;
            if (type == PivotFilterValues.July) return DynamicFilterValues.July;
            if (type == PivotFilterValues.August) return DynamicFilterValues.August;
            if (type == PivotFilterValues.September) return DynamicFilterValues.September;
            if (type == PivotFilterValues.October) return DynamicFilterValues.October;
            if (type == PivotFilterValues.November) return DynamicFilterValues.November;
            if (type == PivotFilterValues.December) return DynamicFilterValues.December;
            if (type == PivotFilterValues.Quarter1) return DynamicFilterValues.Quarter1;
            if (type == PivotFilterValues.Quarter2) return DynamicFilterValues.Quarter2;
            if (type == PivotFilterValues.Quarter3) return DynamicFilterValues.Quarter3;
            if (type == PivotFilterValues.Quarter4) return DynamicFilterValues.Quarter4;

            return null;
        }

        private static bool TryResolveBetweenFilter(PivotFilterValues type,
            out FilterOperatorValues firstOperator,
            out FilterOperatorValues secondOperator,
            out bool matchAll) {
            if (type == PivotFilterValues.CaptionBetween || type == PivotFilterValues.ValueBetween || type == PivotFilterValues.DateBetween) {
                firstOperator = FilterOperatorValues.GreaterThanOrEqual;
                secondOperator = FilterOperatorValues.LessThanOrEqual;
                matchAll = true;
                return true;
            }

            if (type == PivotFilterValues.CaptionNotBetween || type == PivotFilterValues.ValueNotBetween || type == PivotFilterValues.DateNotBetween) {
                firstOperator = FilterOperatorValues.LessThan;
                secondOperator = FilterOperatorValues.GreaterThan;
                matchAll = false;
                return true;
            }

            firstOperator = FilterOperatorValues.Equal;
            secondOperator = FilterOperatorValues.Equal;
            matchAll = true;
            return false;
        }

        private static FilterOperatorValues ResolveSingleFilterOperator(PivotFilterValues type) {
            if (type == PivotFilterValues.CaptionNotEqual || type == PivotFilterValues.CaptionNotContains
                || type == PivotFilterValues.CaptionNotBeginsWith || type == PivotFilterValues.CaptionNotEndsWith
                || type == PivotFilterValues.ValueNotEqual || type == PivotFilterValues.DateNotEqual) {
                return FilterOperatorValues.NotEqual;
            }

            if (type == PivotFilterValues.CaptionGreaterThan || type == PivotFilterValues.ValueGreaterThan || type == PivotFilterValues.DateNewerThan) {
                return FilterOperatorValues.GreaterThan;
            }

            if (type == PivotFilterValues.CaptionGreaterThanOrEqual || type == PivotFilterValues.ValueGreaterThanOrEqual || type == PivotFilterValues.DateNewerThanOrEqual) {
                return FilterOperatorValues.GreaterThanOrEqual;
            }

            if (type == PivotFilterValues.CaptionLessThan || type == PivotFilterValues.ValueLessThan || type == PivotFilterValues.DateOlderThan) {
                return FilterOperatorValues.LessThan;
            }

            if (type == PivotFilterValues.CaptionLessThanOrEqual || type == PivotFilterValues.ValueLessThanOrEqual || type == PivotFilterValues.DateOlderThanOrEqual) {
                return FilterOperatorValues.LessThanOrEqual;
            }

            return FilterOperatorValues.Equal;
        }

        private static string NormalizePivotFilterAutoFilterValue(PivotFilterValues type, string value) {
            if (type == PivotFilterValues.CaptionContains || type == PivotFilterValues.CaptionNotContains) {
                return "*" + value + "*";
            }

            if (type == PivotFilterValues.CaptionBeginsWith || type == PivotFilterValues.CaptionNotBeginsWith) {
                return value + "*";
            }

            if (type == PivotFilterValues.CaptionEndsWith || type == PivotFilterValues.CaptionNotEndsWith) {
                return "*" + value;
            }

            return value;
        }

        private static int ResolvePivotDataFieldIndex(string dataFieldName, IReadOnlyList<ExcelPivotDataField> dataFields) {
            for (int i = 0; i < dataFields.Count; i++) {
                var dataField = dataFields[i];
                if (string.Equals(dataField.FieldName, dataFieldName, StringComparison.OrdinalIgnoreCase)
                    || string.Equals(dataField.DisplayName, dataFieldName, StringComparison.OrdinalIgnoreCase)) {
                    return i;
                }
            }

            throw new ArgumentException($"Data field '{dataFieldName}' was not found in pivot data fields.", nameof(dataFieldName));
        }

        private static int FindPivotItemIndex(string item, IReadOnlyList<string> values, string fieldName, string paramName) {
            for (int i = 0; i < values.Count; i++) {
                if (string.Equals(values[i], item, StringComparison.OrdinalIgnoreCase)) {
                    return i;
                }
            }

            throw new ArgumentException($"Item '{item}' was not found in pivot field '{fieldName}'.", paramName);
        }

        private static uint? ResolveNumberFormatId(WorkbookPart workbookPart, uint? numberFormatId, string? numberFormat) {
            if (numberFormatId.HasValue) return numberFormatId.Value;
            if (string.IsNullOrWhiteSpace(numberFormat)) return null;
            return GetOrCreateNumberFormatId(workbookPart, numberFormat!.Trim());
        }

        private static uint GetOrCreateNumberFormatId(WorkbookPart workbookPart, string numberFormat) {
            WorkbookStylesPart? stylesPart = workbookPart.WorkbookStylesPart;
            if (stylesPart == null) {
                stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
            }

            Stylesheet stylesheet = stylesPart.Stylesheet ??= new Stylesheet();
            EnsureDefaultStylePrimitives(stylesheet);

            stylesheet.NumberingFormats ??= new NumberingFormats();
            NumberingFormat? existingFormat = stylesheet.NumberingFormats.Elements<NumberingFormat>()
                .FirstOrDefault(n => n.FormatCode != null && string.Equals(n.FormatCode.Value, numberFormat, StringComparison.Ordinal));

            if (existingFormat?.NumberFormatId?.Value is uint existingId) {
                return existingId;
            }

            uint formatId = stylesheet.NumberingFormats.Elements<NumberingFormat>().Any()
                ? Math.Max(164U, stylesheet.NumberingFormats.Elements<NumberingFormat>().Max(n => n.NumberFormatId?.Value ?? 0U) + 1U)
                : 164U;

            stylesheet.NumberingFormats.Append(new NumberingFormat {
                NumberFormatId = formatId,
                FormatCode = StringValue.FromString(numberFormat)
            });
            stylesheet.NumberingFormats.Count = (uint)stylesheet.NumberingFormats.Count();
            stylesPart.Stylesheet.Save();
            return formatId;
        }
    }
}
