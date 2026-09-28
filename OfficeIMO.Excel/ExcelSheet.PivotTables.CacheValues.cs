using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Text;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private PivotFieldValue GetPivotFieldValue(int row, int column, ExcelPivotGrouping? grouping) {
            string text = TryGetCellText(row, column, out string cellText) ? cellText : string.Empty;
            if (TryGetExistingCell(row, column)?.DataType?.Value == DocumentFormat.OpenXml.Spreadsheet.CellValues.Error)
                return PivotFieldValue.FromError(text);
            if (string.IsNullOrEmpty(text)) {
                return PivotFieldValue.Blank();
            }

            if (grouping?.IsDateGrouping == true) {
                if (TryGetPivotDateValue(row, column, text, out var date)) {
                    return PivotFieldValue.FromDate(date);
                }
            }

            var snapshot = GetCellValueSnapshot(row, column);
            if (snapshot.Value is bool boolean) {
                return PivotFieldValue.FromBoolean(boolean);
            }

            if (snapshot.Value is double number) {
                if (TryGetPivotDateValueFromStyle(row, column, number, out var styledDate)) {
                    return PivotFieldValue.FromDate(grouping == null
                        ? ExcelPivotCacheDateCodec.FromSerial(number, _excelDocument.DateSystem) : styledDate);
                }

                return PivotFieldValue.FromNumber(number);
            }

            if (grouping?.GroupBy == ExcelPivotGroupBy.Range
                && double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out number)) {
                return PivotFieldValue.FromNumber(number);
            }

            return PivotFieldValue.FromText(text);
        }

        private bool TryGetPivotDateValueFromStyle(int row, int column, double serial, out DateTime date) {
            date = default;
            var cell = TryGetExistingCell(row, column);
            if (cell?.StyleIndex?.Value is not uint styleIndex) {
                return false;
            }

            var styles = _pivotStylesCache ??= StylesCache.Build(_spreadSheetDocument);
            if (!styles.HasDateStyles || !styles.IsDateLike(styleIndex)) {
                return false;
            }

            try {
                date = ExcelDateSystemConverter.FromSerial(serial, _excelDocument.DateSystem);
                return true;
            } catch (ArgumentException) {
                date = default;
                return false;
            }
        }

        internal bool IsPivotDateSourceValue(int row, int column) {
            ExcelCellData snapshot = GetCellValueSnapshot(row, column);
            return snapshot.Value is double serial
                && TryGetPivotDateValueFromStyle(row, column, serial, out _);
        }

        private PivotFieldValue GetPivotFieldValue(object? value) {
            if (value == null || value == DBNull.Value) {
                return PivotFieldValue.Blank();
            }

            return value switch {
                bool boolean => PivotFieldValue.FromBoolean(boolean),
                byte number => PivotFieldValue.FromNumber(number),
                sbyte number => PivotFieldValue.FromNumber(number),
                short number => PivotFieldValue.FromNumber(number),
                ushort number => PivotFieldValue.FromNumber(number),
                int number => PivotFieldValue.FromNumber(number),
                uint number => PivotFieldValue.FromNumber(number),
                long number => PivotFieldValue.FromNumber(number),
                ulong number when number <= long.MaxValue => PivotFieldValue.FromNumber(number),
                float number => PivotFieldValue.FromNumber(number),
                double number => PivotFieldValue.FromNumber(number),
                decimal number => PivotFieldValue.FromNumber((double)number),
                DateTime dateTime => CreatePivotFieldDateValue(dateTime),
                DateTimeOffset dateTimeOffset => CreatePivotFieldDateValue(_excelDocument.DateTimeOffsetWriteStrategy(dateTimeOffset)),
#if NET6_0_OR_GREATER
                DateOnly dateOnly => CreatePivotFieldDateValue(dateOnly.ToDateTime(TimeOnly.MinValue)),
#endif
                string text => CreatePivotFieldTextValue(text),
                _ => PivotFieldValue.FromText(FormatPivotFieldText(value, _excelDocument.DateTimeOffsetWriteStrategy, _excelDocument.DateSystem))
            };
        }

        private string GetPivotFieldText(object? value) {
            if (value == null || value == DBNull.Value) {
                return string.Empty;
            }

            return FormatPivotFieldText(value, _excelDocument.DateTimeOffsetWriteStrategy, _excelDocument.DateSystem);
        }

        private static PivotFieldValue CreatePivotFieldTextValue(string text)
            => text.Length == 0 ? PivotFieldValue.Blank() : PivotFieldValue.FromText(text);

        private PivotFieldValue CreatePivotFieldDateValue(DateTime date) => PivotFieldValue.FromDate(
            ExcelPivotCacheDateCodec.FromSerial(ExcelDateSystemConverter.ToSerial(date, _excelDocument.DateSystem), _excelDocument.DateSystem));

        private static string FormatPivotFieldText(object value, Func<DateTimeOffset, DateTime> dateTimeOffsetWriteStrategy, ExcelDateSystem dateSystem) {
            return value switch {
                string text => text,
                bool boolean => boolean ? "1" : "0",
                DateTime dateTime => InvariantNumberText.Get(ExcelDateSystemConverter.ToSerial(dateTime, dateSystem)),
                DateTimeOffset dateTimeOffset => InvariantNumberText.Get(ExcelDateSystemConverter.ToSerial(dateTimeOffsetWriteStrategy(dateTimeOffset), dateSystem)),
#if NET6_0_OR_GREATER
                DateOnly dateOnly => InvariantNumberText.Get(ExcelDateSystemConverter.ToSerial(dateOnly.ToDateTime(TimeOnly.MinValue), dateSystem)),
#endif
                double number => InvariantNumberText.Get(number),
                float number => InvariantNumberText.Get(number),
                decimal number => number.ToString(CultureInfo.InvariantCulture),
                IFormattable formattable => TrimPivotFieldText(formattable.ToString(null, CultureInfo.InvariantCulture)),
                _ => TrimPivotFieldText(value.ToString())
            };
        }

        private static string TrimPivotFieldText(string? text) {
            if (string.IsNullOrEmpty(text)) {
                return string.Empty;
            }

            string normalized = text!;
            int last = normalized.Length - 1;
            return char.IsWhiteSpace(normalized[0]) || char.IsWhiteSpace(normalized[last])
                ? normalized.Trim()
                : normalized;
        }

        private PivotFieldValue GetGeneratedPivotDateFieldValue(int row, int column, ExcelPivotGroupBy groupBy) {
            string text = TryGetCellText(row, column, out string cellText) ? cellText.Trim() : string.Empty;
            if (string.IsNullOrEmpty(text)) {
                return PivotFieldValue.Blank();
            }

            return TryGetPivotDateValue(row, column, text, out var date)
                ? PivotFieldValue.FromText(FormatGeneratedDateGroupValue(date, groupBy))
                : PivotFieldValue.FromText(text);
        }

        private bool TryGetPivotDateValue(int row, int column, string text, out DateTime date) {
            var snapshot = GetCellValueSnapshot(row, column);
            if (snapshot.Value is double serial) {
                try {
                    date = ExcelDateSystemConverter.FromSerial(serial, _excelDocument.DateSystem);
                    return true;
                } catch {
                    // Fall through to string parsing when a numeric value is not a valid Excel date.
                }
            }

            if (DateTime.TryParse(text, CultureInfo.CurrentCulture, DateTimeStyles.None, out date)
                || DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.None, out date)) {
                return true;
            }

            date = default;
            return false;
        }

        private static string FormatGeneratedDateGroupValue(DateTime date, ExcelPivotGroupBy groupBy) {
            if (groupBy == ExcelPivotGroupBy.Years) return date.Year.ToString(CultureInfo.InvariantCulture);
            if (groupBy == ExcelPivotGroupBy.Quarters) return $"Q{((date.Month - 1) / 3) + 1}";
            if (groupBy == ExcelPivotGroupBy.Months) return date.ToString("MMMM", CultureInfo.InvariantCulture);
            if (groupBy == ExcelPivotGroupBy.Days) return date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
            if (groupBy == ExcelPivotGroupBy.Hours) return date.Hour.ToString("00", CultureInfo.InvariantCulture);
            if (groupBy == ExcelPivotGroupBy.Minutes) return date.ToString("HH:mm", CultureInfo.InvariantCulture);
            if (groupBy == ExcelPivotGroupBy.Seconds) return date.ToString("HH:mm:ss", CultureInfo.InvariantCulture);
            return date.ToString("O", CultureInfo.InvariantCulture);
        }

        private static string GetDateGroupFieldSuffix(ExcelPivotGroupBy groupBy) {
            if (groupBy == ExcelPivotGroupBy.Years) return "Years";
            if (groupBy == ExcelPivotGroupBy.Quarters) return "Quarters";
            if (groupBy == ExcelPivotGroupBy.Months) return "Months";
            if (groupBy == ExcelPivotGroupBy.Days) return "Days";
            if (groupBy == ExcelPivotGroupBy.Hours) return "Hours";
            if (groupBy == ExcelPivotGroupBy.Minutes) return "Minutes";
            if (groupBy == ExcelPivotGroupBy.Seconds) return "Seconds";
            return groupBy.ToString();
        }

        private static SharedItems BuildSharedItems(PivotFieldValues values, ExcelPivotGrouping? grouping, bool appendItems = true) {
            bool hasBlank = false;
            bool hasDate = false;
            bool hasNumber = false;
            bool hasString = false;
            bool containsInteger = true;
            double minNumber = 0D;
            double maxNumber = 0D;
            DateTime minDate = default;
            DateTime maxDate = default;
            int numberCount = 0;
            int dateCount = 0;

            foreach (var value in values.Items) {
                switch (value.Kind) {
                    case PivotFieldValueKind.Blank:
                        hasBlank = true;
                        break;
                    case PivotFieldValueKind.Boolean:
                        hasString = true;
                        break;
                    case PivotFieldValueKind.Number:
                        hasNumber = true;
                        if (value.Number.HasValue) {
                            double number = value.Number.Value;
                            if (numberCount == 0) {
                                minNumber = number;
                                maxNumber = number;
                            } else {
                                if (number < minNumber) minNumber = number;
                                if (number > maxNumber) maxNumber = number;
                            }

                            if (Math.Abs(number - Math.Round(number)) >= 0.0000001d) {
                                containsInteger = false;
                            }

                            numberCount++;
                        }
                        break;
                    case PivotFieldValueKind.Date:
                        hasDate = true;
                        if (value.Date.HasValue) {
                            DateTime date = value.Date.Value;
                            if (dateCount == 0) {
                                minDate = date;
                                maxDate = date;
                            } else {
                                if (date < minDate) minDate = date;
                                if (date > maxDate) maxDate = date;
                            }

                            dateCount++;
                        }
                        break;
                    default:
                        hasString = true;
                        break;
                }
            }

            var sharedItems = new SharedItems {
                ContainsString = hasString,
                ContainsSemiMixedTypes = hasString || hasBlank,
                ContainsMixedTypes = (hasNumber && hasString) || (hasDate && (hasNumber || hasString)),
                ContainsBlank = hasBlank,
                ContainsDate = hasDate,
                ContainsNonDate = !hasDate || hasNumber || hasString,
                // Excel classifies fields containing date items as date fields even
                // when numeric items coexist. Numeric flags make those caches unreadable.
                ContainsNumber = hasNumber && !hasDate
            };

            if (appendItems) {
                sharedItems.Count = (uint)values.Items.Count;
            }

            if (numberCount > 0) {
                if (dateCount == 0 && appendItems) {
                    sharedItems.MinValue = minNumber;
                    sharedItems.MaxValue = maxNumber;
                }
                if (!hasDate) sharedItems.ContainsInteger = containsInteger;
            }

            if (dateCount > 0 && appendItems) {
                sharedItems.MinDate = minDate;
                sharedItems.MaxDate = maxDate;
            }

            if (appendItems) {
                foreach (var value in values.Items) {
                    sharedItems.Append(value.Kind switch {
                        PivotFieldValueKind.Blank => new MissingItem(),
                        PivotFieldValueKind.Boolean => new BooleanItem { Val = value.Boolean!.Value },
                        PivotFieldValueKind.Number => new NumberItem { Val = value.Number!.Value },
                        PivotFieldValueKind.Date => new DateTimeItem { Val = value.Date!.Value },
                        PivotFieldValueKind.Error => new ErrorItem { Val = value.Text },
                        _ => new StringItem { Val = value.Text }
                    });
                }
            }

            return sharedItems;
        }

        private static FieldGroup CreatePivotFieldGroup(ExcelPivotGrouping grouping, PivotFieldValues? groupItems = null, uint? baseFieldIndex = null, uint? parentFieldIndex = null) {
            var range = new RangeProperties {
                AutoStart = grouping.AutoStart,
                AutoEnd = grouping.AutoEnd,
                GroupBy = grouping.GroupBy.ToOpenXml()
            };

            if (grouping.StartDate.HasValue) range.StartDate = grouping.StartDate.Value;
            if (grouping.EndDate.HasValue) range.EndDate = grouping.EndDate.Value;
            if (grouping.StartNumber.HasValue) range.StartNumber = grouping.StartNumber.Value;
            if (grouping.EndNumber.HasValue) range.EndNum = grouping.EndNumber.Value;
            if (grouping.Interval.HasValue) range.GroupInterval = grouping.Interval.Value;

            var fieldGroup = new FieldGroup(range);
            if (baseFieldIndex.HasValue) fieldGroup.Base = baseFieldIndex.Value;
            if (parentFieldIndex.HasValue) fieldGroup.ParentId = parentFieldIndex.Value;
            if (groupItems != null) {
                fieldGroup.Append(BuildGroupItems(BuildAuthorNumericGroupLabels(grouping) ?? groupItems));
            }

            return fieldGroup;
        }

        private static GroupItems BuildGroupItems(PivotFieldValues values) {
            var groupItems = new GroupItems { Count = (uint)values.Items.Count };
            foreach (var value in values.Items) {
                groupItems.Append(value.Kind switch {
                    PivotFieldValueKind.Blank => new MissingItem(),
                    PivotFieldValueKind.Boolean => new BooleanItem { Val = value.Boolean!.Value },
                    PivotFieldValueKind.Number => new NumberItem { Val = value.Number!.Value },
                    PivotFieldValueKind.Date => new DateTimeItem { Val = value.Date!.Value },
                    PivotFieldValueKind.Error => new ErrorItem { Val = value.Text },
                    _ => new StringItem { Val = value.Text }
                });
            }

            return groupItems;
        }

        private enum PivotFieldValueKind {
            Blank,
            Text,
            Boolean,
            Number,
            Date,
            Error
        }

        private sealed class PivotFieldValue : IEquatable<PivotFieldValue> {
            private PivotFieldValue(PivotFieldValueKind kind, string text, bool? boolean = null, double? number = null, DateTime? date = null) {
                Kind = kind;
                Text = text;
                Boolean = boolean;
                Number = number;
                Date = date;
            }

            public PivotFieldValueKind Kind { get; }

            public string Text { get; }

            public bool? Boolean { get; }

            public double? Number { get; }

            public DateTime? Date { get; }

            public bool Equals(PivotFieldValue? other) => other != null && Kind == other.Kind && StringComparer.OrdinalIgnoreCase.Equals(Text, other.Text);

            public override bool Equals(object? other) => other is PivotFieldValue value && Equals(value);

            public override int GetHashCode() => ((int)Kind * 397) ^ StringComparer.OrdinalIgnoreCase.GetHashCode(Text);

            public static PivotFieldValue FromError(string text) => new(PivotFieldValueKind.Error, text);

            public static PivotFieldValue Blank() => new(PivotFieldValueKind.Blank, string.Empty);

            public static PivotFieldValue FromText(string text) => new(PivotFieldValueKind.Text, text);

            public static PivotFieldValue FromBoolean(bool boolean) => new(PivotFieldValueKind.Boolean, boolean ? "1" : "0", boolean: boolean);

            public static PivotFieldValue FromNumber(double number) => new(PivotFieldValueKind.Number, InvariantNumberText.Get(number), number: number);

            public static PivotFieldValue FromDate(DateTime date) => new(PivotFieldValueKind.Date, date.ToString("O", CultureInfo.InvariantCulture), date: date);
        }

        private sealed class PivotFieldValues {
            public PivotFieldValues(IReadOnlyList<PivotFieldValue> items) {
                Items = items;
                TextValues = CreateTextValues(items);
            }

            public IReadOnlyList<PivotFieldValue> Items { get; }

            public IReadOnlyList<string> TextValues { get; }

            private static IReadOnlyList<string> CreateTextValues(IReadOnlyList<PivotFieldValue> items) {
                if (items.Count == 0) {
                    return Array.Empty<string>();
                }

                var textValues = new string[items.Count];
                for (int i = 0; i < items.Count; i++) {
                    textValues[i] = items[i].Text;
                }

                return textValues;
            }
        }

        private sealed class GeneratedPivotGroupingField {
            public GeneratedPivotGroupingField(int sourceIndex, int fieldIndex, int? parentFieldIndex, string fieldName, ExcelPivotGroupBy groupBy, ExcelPivotGrouping grouping) {
                SourceIndex = sourceIndex;
                FieldIndex = fieldIndex;
                ParentFieldIndex = parentFieldIndex;
                FieldName = fieldName;
                GroupBy = groupBy;
                Grouping = grouping;
            }

            public int SourceIndex { get; }

            public int FieldIndex { get; }

            public int? ParentFieldIndex { get; }

            public string FieldName { get; }

            public ExcelPivotGroupBy GroupBy { get; }

            public ExcelPivotGrouping Grouping { get; }
        }
    }
}
