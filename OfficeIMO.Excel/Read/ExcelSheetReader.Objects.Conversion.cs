using System.Globalization;
using System.Diagnostics.CodeAnalysis;
using System.Reflection;
using System.Threading;
using System.Text;
using System.ComponentModel;
using System.Linq.Expressions;
using System.Runtime.Serialization;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Xml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Object-mapping readers for <see cref="ExcelSheetReader"/>.
    /// </summary>
    public sealed partial class ExcelSheetReader {
        private object? TryChangeType<TTarget>(object value, TypedPropertyBinding<TTarget> binding, CultureInfo culture) {
            if (value == null) return null;

            var hook = _opt.TypeConverter;
            if (hook != null) {
                var (ok, v) = hook(value, binding.DestinationType, culture);
                if (ok) return v;
            }

            var srcType = value.GetType();
            if (binding.PropertyType.IsAssignableFrom(srcType)) return value;

            return binding.ConvertValue(value, culture);
        }

        private bool TryConvertCellForBinding<TTarget>(
            DocumentFormat.OpenXml.Spreadsheet.Cell cell,
            TypedPropertyBinding<TTarget> binding,
            out object? converted) {
            converted = null;

            if (_opt.CellValueConverter != null || _opt.TypeConverter != null) {
                object? value = ConvertCell(cell);
                if (value is null) {
                    return binding.IsNullable;
                }

                converted = TryChangeType(value, binding, _opt.Culture);
                return converted is not null || binding.IsNullable;
            }

            CellValues? typeHint = cell.DataType?.Value;
            bool hasFormula = cell.CellFormula is not null;
            string? formulaText = hasFormula ? ExtractFormulaText(cell) : null;
            bool preferFormulaText = hasFormula && !_opt.UseCachedFormulaResult && formulaText != null;
            string? rawText = preferFormulaText ? null : ExtractRawText(cell);
            string? inlineText = preferFormulaText ? null : ExtractInlineString(cell, typeHint);

            if (rawText == null && inlineText == null && formulaText == null && !CellHasExplicitBlank(cell)) {
                return binding.IsNullable;
            }

            if (hasFormula && (!_opt.UseCachedFormulaResult || rawText == null)) {
                if (binding.DestinationType == typeof(string)) {
                    converted = formulaText ?? rawText ?? inlineText;
                    return converted is not null || binding.IsNullable;
                }

                return TryConvertCellForBindingFallback(cell, binding, out converted);
            }

            if (!string.IsNullOrEmpty(inlineText)
                && ReturnBindingConversion(TryConvertStringForBinding(inlineText, binding, out converted), binding, converted)) {
                return true;
            }

            if (typeHint == CellValues.SharedString) {
                string? text = rawText;
                if (TryParseSharedStringIndex(rawText, out int sstIndex)) {
                    text = GetSharedString(sstIndex);
                }

                if (ReturnBindingConversion(TryConvertStringForBinding(text, binding, out converted), binding, converted)) {
                    return true;
                }

                return TryConvertCellForBindingFallback(cell, binding, out converted);
            }

            if (typeHint == CellValues.Boolean && rawText != null) {
                if (ReturnBindingConversion(TryConvertBooleanForBinding(rawText == "1", binding, out converted), binding, converted)) {
                    return true;
                }

                return TryConvertCellForBindingFallback(cell, binding, out converted);
            }

            if (typeHint == CellValues.String || typeHint == CellValues.InlineString) {
                if (ReturnBindingConversion(TryConvertStringForBinding(rawText ?? inlineText, binding, out converted), binding, converted)) {
                    return true;
                }

                return TryConvertCellForBindingFallback(cell, binding, out converted);
            }

            if (typeHint == CellValues.Date && rawText != null) {
                if (DateTime.TryParse(rawText, _opt.Culture, DateTimeStyles.AssumeLocal, out var dt)
                    && ReturnBindingConversion(TryConvertDateTimeForBinding(dt, binding, out converted), binding, converted)) {
                    return true;
                }

                if (ReturnBindingConversion(TryConvertStringForBinding(rawText, binding, out converted), binding, converted)) {
                    return true;
                }

                return TryConvertCellForBindingFallback(cell, binding, out converted);
            }

            if (rawText == null) {
                return binding.IsNullable;
            }

            uint? styleIndex = null;
            if (_opt.TreatDatesUsingNumberFormat && binding.NeedsDateStyleConversion) {
                styleIndex = cell.StyleIndex?.Value;
                if (styleIndex is not null && Styles.IsDateLike(styleIndex.Value)) {
                    if ((TryParseInvariantDoubleFast(rawText, out var oa)
                            || double.TryParse(rawText, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out oa))
                        && ReturnBindingConversion(TryConvertDateTimeForBinding(DateTime.FromOADate(oa), binding, out converted), binding, converted)) {
                        return true;
                    }

                    if (ReturnBindingConversion(TryConvertStringForBinding(rawText, binding, out converted), binding, converted)) {
                        return true;
                    }

                    return TryConvertCellForBindingFallback(cell, binding, out converted);
                }
            }

            if (ReturnBindingConversion(TryConvertNumericTextForBinding(rawText, binding, out converted), binding, converted)) {
                return true;
            }

            return TryConvertCellForBindingFallback(cell, binding, out converted);
        }

        private bool TryConvertCellForBindingFallback<TTarget>(
            DocumentFormat.OpenXml.Spreadsheet.Cell cell,
            TypedPropertyBinding<TTarget> binding,
            out object? converted) {
            converted = null;
            var raw = SnapshotCell(cell);
            if (raw.RawText == null && raw.InlineText == null && raw.FormulaText == null && !CellHasExplicitBlank(cell)) {
                return binding.IsNullable;
            }

            if (TryConvertRawForBinding(raw, binding, out converted)) {
                return converted is not null || binding.IsNullable;
            }

            object? typedValue = ConvertRaw(raw).TypedValue;
            if (typedValue is null) {
                return binding.IsNullable;
            }

            converted = TryChangeType(typedValue, binding, _opt.Culture);
            return converted is not null || binding.IsNullable;
        }

        private static bool ReturnBindingConversion<TTarget>(
            bool convertedByFastPath,
            TypedPropertyBinding<TTarget> binding,
            object? converted) {
            return convertedByFastPath && (converted is not null || binding.IsNullable);
        }

        private bool TryConvertRawForBinding<TTarget>(
            CellRaw raw,
            TypedPropertyBinding<TTarget> binding,
            out object? converted) {
            converted = null;

            if (raw.HasFormula && (!_opt.UseCachedFormulaResult || raw.RawText == null)) {
                if (binding.DestinationType == typeof(string)) {
                    converted = raw.FormulaText ?? raw.RawText ?? raw.InlineText;
                    return true;
                }

                return false;
            }

            if (!string.IsNullOrEmpty(raw.InlineText)) {
                return TryConvertStringForBinding(raw.InlineText, binding, out converted);
            }

            if (raw.TypeHint == DocumentFormat.OpenXml.Spreadsheet.CellValues.SharedString) {
                if (!TryParseSharedStringIndex(raw.RawText, out int sstIndex)) {
                    return TryConvertStringForBinding(raw.RawText, binding, out converted);
                }

                return TryConvertStringForBinding(GetSharedString(sstIndex), binding, out converted);
            }

            if (raw.TypeHint == DocumentFormat.OpenXml.Spreadsheet.CellValues.Boolean && raw.RawText != null) {
                return TryConvertBooleanForBinding(raw.RawText == "1", binding, out converted);
            }

            if (raw.TypeHint == DocumentFormat.OpenXml.Spreadsheet.CellValues.String
                || raw.TypeHint == DocumentFormat.OpenXml.Spreadsheet.CellValues.InlineString) {
                return TryConvertStringForBinding(raw.RawText ?? raw.InlineText, binding, out converted);
            }

            if (raw.TypeHint == DocumentFormat.OpenXml.Spreadsheet.CellValues.Date && raw.RawText != null) {
                if (DateTime.TryParse(raw.RawText, _opt.Culture, DateTimeStyles.AssumeLocal, out var dt)) {
                    return TryConvertDateTimeForBinding(dt, binding, out converted);
                }

                return TryConvertStringForBinding(raw.RawText, binding, out converted);
            }

            if (raw.RawText == null) {
                return false;
            }

            if (_opt.TreatDatesUsingNumberFormat
                && binding.NeedsDateStyleConversion
                && raw.StyleIndex is not null
                && Styles.IsDateLike(raw.StyleIndex.Value)) {
                if (TryParseInvariantDoubleFast(raw.RawText, out var oa)
                    || double.TryParse(raw.RawText, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out oa)) {
                    return TryConvertDateTimeForBinding(DateTime.FromOADate(oa), binding, out converted);
                }

                return TryConvertStringForBinding(raw.RawText, binding, out converted);
            }

            return TryConvertNumericTextForBinding(raw.RawText, binding, out converted);
        }

        private bool TryConvertRawCellForBinding<TTarget>(
            CellRaw raw,
            TypedPropertyBinding<TTarget> binding,
            out object? converted) {
            converted = null;

            if (raw.RawText == null && raw.InlineText == null && raw.FormulaText == null) {
                return binding.IsNullable;
            }

            if (TryConvertRawForBinding(raw, binding, out converted)) {
                return converted is not null || binding.IsNullable;
            }

            object? typedValue = ConvertRaw(raw).TypedValue;
            if (typedValue is null) {
                return binding.IsNullable;
            }

            converted = TryChangeType(typedValue, binding, _opt.Culture);
            return converted is not null || binding.IsNullable;
        }

        private bool TrySetRawCellForBinding<TTarget>(
            CellRaw raw,
            TypedPropertyBinding<TTarget> binding,
            TTarget target) {
            if (_opt.CellValueConverter != null || _opt.TypeConverter != null) {
                object? typedValue = ConvertRaw(raw).TypedValue;
                if (typedValue is null) {
                    if (binding.IsNullable) {
                        binding.SetValue(target, null);
                        return true;
                    }

                    return false;
                }

                object? converted = TryChangeType(typedValue, binding, _opt.Culture);
                if (converted is not null || binding.IsNullable) {
                    binding.SetValue(target, converted);
                    return true;
                }

                return false;
            }

            if (raw.RawText == null && raw.InlineText == null && raw.FormulaText == null) {
                if (binding.IsNullable) {
                    binding.SetValue(target, null);
                    return true;
                }

                return false;
            }

            if (raw.HasFormula && (!_opt.UseCachedFormulaResult || raw.RawText == null)) {
                if (binding.BindingKind == TypedBindingKind.String) {
                    string? formulaValue = raw.FormulaText ?? raw.RawText ?? raw.InlineText;
                    SetStringBinding(binding, target, formulaValue);
                    return formulaValue is not null || binding.IsNullable;
                }

                return false;
            }

            if (!string.IsNullOrEmpty(raw.InlineText)) {
                if (TrySetStringTextBinding(raw.InlineText, binding, target)) {
                    return true;
                }

                return TrySetRawCellForBindingFallback(raw, binding, target);
            }

            if (raw.TypeHint == DocumentFormat.OpenXml.Spreadsheet.CellValues.SharedString) {
                string? text = TryParseSharedStringIndex(raw.RawText, out int sstIndex)
                    ? GetSharedString(sstIndex)
                    : raw.RawText;
                if (TrySetStringTextBinding(text, binding, target)) {
                    return true;
                }

                return TrySetRawCellForBindingFallback(raw, binding, target);
            }

            if (raw.TypeHint == DocumentFormat.OpenXml.Spreadsheet.CellValues.Boolean && raw.RawText != null) {
                bool boolValue = raw.RawText == "1";
                if (binding.SetBoolean != null && binding.BindingKind == TypedBindingKind.Boolean) {
                    binding.SetBoolean(target, boolValue);
                    return true;
                }

                if (binding.SetString != null && binding.BindingKind == TypedBindingKind.String) {
                    binding.SetString(target, boolValue.ToString());
                    return true;
                }

                return TrySetRawCellForBindingFallback(raw, binding, target);
            }

            if (raw.TypeHint == DocumentFormat.OpenXml.Spreadsheet.CellValues.String
                || raw.TypeHint == DocumentFormat.OpenXml.Spreadsheet.CellValues.InlineString) {
                if (TrySetStringTextBinding(raw.RawText ?? raw.InlineText, binding, target)) {
                    return true;
                }

                return TrySetRawCellForBindingFallback(raw, binding, target);
            }

            if (raw.TypeHint == DocumentFormat.OpenXml.Spreadsheet.CellValues.Date && raw.RawText != null) {
                if (binding.SetDateTime != null
                    && DateTime.TryParse(raw.RawText, _opt.Culture, DateTimeStyles.AssumeLocal, out var dateValue)) {
                    binding.SetDateTime(target, dateValue);
                    return true;
                }

                if (TrySetStringTextBinding(raw.RawText, binding, target)) {
                    return true;
                }

                return TrySetRawCellForBindingFallback(raw, binding, target);
            }

            if (raw.RawText == null) {
                return false;
            }

            if (_opt.TreatDatesUsingNumberFormat
                && binding.NeedsDateStyleConversion
                && raw.StyleIndex is not null
                && Styles.IsDateLike(raw.StyleIndex.Value)) {
                if (TryParseInvariantDoubleFast(raw.RawText, out var oa)
                    || double.TryParse(raw.RawText, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out oa)) {
                    DateTime dateValue = DateTime.FromOADate(oa);
                    if (binding.SetDateTime != null && binding.BindingKind == TypedBindingKind.DateTime) {
                        binding.SetDateTime(target, dateValue);
                        return true;
                    }

                    if (binding.SetString != null && binding.BindingKind == TypedBindingKind.String) {
                        binding.SetString(target, dateValue.ToString(_opt.Culture));
                        return true;
                    }
                }

                if (TrySetStringTextBinding(raw.RawText, binding, target)) {
                    return true;
                }

                return TrySetRawCellForBindingFallback(raw, binding, target);
            }

            if (TrySetNumericTextBinding(raw.RawText, binding, target)) {
                return true;
            }

            return TrySetRawCellForBindingFallback(raw, binding, target);
        }

        private bool TrySetRawCellForBindingFallback<TTarget>(
            CellRaw raw,
            TypedPropertyBinding<TTarget> binding,
            TTarget target) {
            if (TryConvertRawCellForBinding(raw, binding, out object? converted)) {
                binding.SetValue(target, converted);
                return true;
            }

            return false;
        }

        private bool TrySetStringTextBinding<TTarget>(
            string? text,
            TypedPropertyBinding<TTarget> binding,
            TTarget target) {
            if (text == null) {
                if (binding.IsNullable) {
                    binding.SetValue(target, null);
                    return true;
                }

                return false;
            }

            if (binding.SetString != null && binding.BindingKind == TypedBindingKind.String) {
                binding.SetString(target, text);
                return true;
            }

            if (binding.SetBoolean != null
                && binding.BindingKind == TypedBindingKind.Boolean
                && bool.TryParse(text, out bool boolValue)) {
                binding.SetBoolean(target, boolValue);
                return true;
            }

            if (binding.SetDateTime != null
                && binding.BindingKind == TypedBindingKind.DateTime
                && DateTime.TryParse(text, _opt.Culture, DateTimeStyles.AssumeLocal, out var dateValue)) {
                binding.SetDateTime(target, dateValue);
                return true;
            }

            return TrySetNumericTextBinding(text, binding, target);
        }

        private bool TrySetNumericTextBinding<TTarget>(
            string rawText,
            TypedPropertyBinding<TTarget> binding,
            TTarget target) {
            switch (binding.BindingKind) {
                case TypedBindingKind.Int32: {
                    if (binding.SetInt32 == null) {
                        return false;
                    }

                    if (TryParseRawInt32(rawText, out int intValue)) {
                        binding.SetInt32(target, intValue);
                        return true;
                    }

                    if (TryParseRawDouble(rawText, out double doubleValue)
                        && doubleValue >= int.MinValue
                        && doubleValue <= int.MaxValue
                        && Math.Truncate(doubleValue) == doubleValue) {
                        binding.SetInt32(target, (int)doubleValue);
                        return true;
                    }

                    return false;
                }

                case TypedBindingKind.Int64: {
                    if (binding.SetInt64 == null) {
                        return false;
                    }

                    if (TryParseRawInt64(rawText, out long longValue)) {
                        binding.SetInt64(target, longValue);
                        return true;
                    }

                    if (TryParseRawDouble(rawText, out double doubleValue)
                        && doubleValue >= long.MinValue
                        && doubleValue <= long.MaxValue
                        && Math.Truncate(doubleValue) == doubleValue) {
                        binding.SetInt64(target, (long)doubleValue);
                        return true;
                    }

                    return false;
                }

                case TypedBindingKind.Double: {
                    if (binding.SetDouble == null) {
                        return false;
                    }

                    if (TryParseRawDouble(rawText, out double doubleValue)) {
                        binding.SetDouble(target, doubleValue);
                        return true;
                    }

                    return false;
                }

                case TypedBindingKind.Decimal: {
                    if (binding.SetDecimal == null) {
                        return false;
                    }

                    if (TryParseRawDecimal(rawText, out decimal decimalValue)) {
                        binding.SetDecimal(target, decimalValue);
                        return true;
                    }

                    return false;
                }

                case TypedBindingKind.Boolean: {
                    if (binding.SetBoolean == null) {
                        return false;
                    }

                    if (rawText == "1") {
                        binding.SetBoolean(target, true);
                        return true;
                    }

                    if (rawText == "0") {
                        binding.SetBoolean(target, false);
                        return true;
                    }

                    return false;
                }

                case TypedBindingKind.String: {
                    if (binding.SetString == null) {
                        return false;
                    }

                    binding.SetString(target, rawText);
                    return true;
                }

                default:
                    return false;
            }
        }

        private static void SetStringBinding<TTarget>(
            TypedPropertyBinding<TTarget> binding,
            TTarget target,
            string? value) {
            if (binding.SetString != null && binding.BindingKind == TypedBindingKind.String) {
                binding.SetString(target, value);
            } else {
                binding.SetValue(target, value);
            }
        }

        private bool ShouldRetryRawDateStyledNumericBinding<TTarget>(
            CellRaw raw,
            TypedPropertyBinding<TTarget> binding) {
            if (!_opt.TreatDatesUsingNumberFormat
                || binding.NeedsDateStyleConversion
                || !IsNumericBindingDestination(binding.BindingKind)
                || raw.RawText == null
                || raw.StyleIndex is null
                || !Styles.IsDateLike(raw.StyleIndex.Value)) {
                return false;
            }

            return TryParseInvariantDoubleFast(raw.RawText, out _)
                || double.TryParse(raw.RawText, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out _);
        }

        private static bool IsNumericBindingDestination(TypedBindingKind bindingKind) {
            return bindingKind == TypedBindingKind.Int32
                || bindingKind == TypedBindingKind.Int64
                || bindingKind == TypedBindingKind.Double
                || bindingKind == TypedBindingKind.Decimal;
        }

        private bool TryConvertNumericTextForBinding<TTarget>(
            string rawText,
            TypedPropertyBinding<TTarget> binding,
            out object? converted) {
            converted = null;
            Type destinationType = binding.DestinationType;

            if (destinationType == typeof(int)) {
                if (TryParseRawInt32(rawText, out int intValue)) {
                    converted = intValue;
                    return true;
                }

                if (TryParseRawDouble(rawText, out double doubleValue)
                    && doubleValue >= int.MinValue
                    && doubleValue <= int.MaxValue
                    && Math.Truncate(doubleValue) == doubleValue) {
                    converted = (int)doubleValue;
                    return true;
                }

                return false;
            }

            if (destinationType == typeof(long)) {
                if (TryParseRawInt64(rawText, out long longValue)) {
                    converted = longValue;
                    return true;
                }

                if (TryParseRawDouble(rawText, out double doubleValue)
                    && doubleValue >= long.MinValue
                    && doubleValue <= long.MaxValue
                    && Math.Truncate(doubleValue) == doubleValue) {
                    converted = (long)doubleValue;
                    return true;
                }

                return false;
            }

            if (destinationType == typeof(double)) {
                if (TryParseRawDouble(rawText, out double doubleValue)) {
                    converted = doubleValue;
                    return true;
                }

                return false;
            }

            if (destinationType == typeof(decimal)) {
                if (TryParseRawDecimal(rawText, out decimal decimalValue)) {
                    converted = decimalValue;
                    return true;
                }

                return false;
            }

            if (destinationType == typeof(bool)) {
                if (rawText == "1") {
                    converted = true;
                    return true;
                }

                if (rawText == "0") {
                    converted = false;
                    return true;
                }

                return false;
            }

            if (destinationType == typeof(string)) {
                converted = rawText;
                return true;
            }

            return false;
        }

        private bool TryConvertStringForBinding<TTarget>(
            string? text,
            TypedPropertyBinding<TTarget> binding,
            out object? converted) {
            converted = null;
            if (text == null) {
                return binding.IsNullable;
            }

            Type destinationType = binding.DestinationType;
            if (destinationType == typeof(string)) {
                converted = text;
                return true;
            }

            if (destinationType == typeof(bool) && bool.TryParse(text, out bool boolValue)) {
                converted = boolValue;
                return true;
            }

            if (destinationType == typeof(DateTime)
                && DateTime.TryParse(text, _opt.Culture, DateTimeStyles.AssumeLocal, out var dt)) {
                converted = dt;
                return true;
            }

            return TryConvertNumericTextForBinding(text, binding, out converted);
        }

        private static bool TryConvertBooleanForBinding<TTarget>(
            bool value,
            TypedPropertyBinding<TTarget> binding,
            out object? converted) {
            converted = null;
            Type destinationType = binding.DestinationType;
            if (destinationType == typeof(bool)) {
                converted = value;
                return true;
            }

            if (destinationType == typeof(string)) {
                converted = value.ToString();
                return true;
            }

            return false;
        }

        private bool TryConvertDateTimeForBinding<TTarget>(
            DateTime value,
            TypedPropertyBinding<TTarget> binding,
            out object? converted) {
            converted = null;
            Type destinationType = binding.DestinationType;
            if (destinationType == typeof(DateTime)) {
                converted = value;
                return true;
            }

            if (destinationType == typeof(string)) {
                converted = value.ToString(_opt.Culture);
                return true;
            }

            return false;
        }

        private static Dictionary<string, PropertyInfo> BuildPropertyMap(
            IEnumerable<PropertyInfo> props,
            Func<PropertyInfo, IEnumerable<string>> candidateFactory,
            ICollection<string> diagnostics,
            string typeName,
            string mappingKind) {
            var map = new Dictionary<string, PropertyInfo>(StringComparer.OrdinalIgnoreCase);
            var ambiguous = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var ambiguousProps = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);

            foreach (var prop in props) {
                foreach (string rawCandidate in candidateFactory(prop)) {
                    if (string.IsNullOrWhiteSpace(rawCandidate)) {
                        continue;
                    }

                    string candidate = rawCandidate;
                    if (candidate.Length == 0 || ambiguous.Contains(candidate)) {
                        continue;
                    }

                    if (map.TryGetValue(candidate, out var existing) && existing != prop) {
                        map.Remove(candidate);
                        ambiguous.Add(candidate);
                        if (!ambiguousProps.TryGetValue(candidate, out var propNames)) {
                            propNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { existing.Name };
                            ambiguousProps[candidate] = propNames;
                        }

                        propNames.Add(prop.Name);
                        continue;
                    }

                    if (ambiguousProps.TryGetValue(candidate, out var existingAmbiguousNames)) {
                        existingAmbiguousNames.Add(prop.Name);
                        continue;
                    }

                    map[candidate] = prop;
                }
            }

            foreach (var pair in ambiguousProps.OrderBy(pair => pair.Key, StringComparer.OrdinalIgnoreCase)) {
                diagnostics.Add(
                    $"[TypedRead AmbiguousMapping] Type='{typeName}', match='{mappingKind}', header='{pair.Key}', properties='{string.Join(", ", pair.Value.OrderBy(name => name, StringComparer.OrdinalIgnoreCase))}'.");
            }

            return map;
        }

        private static IEnumerable<string> GetPropertyAliases(PropertyInfo propertyInfo) {
            var yielded = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            void YieldIfUnique(string? candidate, List<string> buffer) {
                if (!string.IsNullOrWhiteSpace(candidate)) {
                    string text = candidate!;
                    if (yielded.Add(text)) {
                        buffer.Add(text);
                    }
                }
            }

            var aliases = new List<string>();

            var displayName = propertyInfo.GetCustomAttribute<DisplayNameAttribute>(inherit: true);
            YieldIfUnique(displayName?.DisplayName, aliases);

            var dataMember = propertyInfo.GetCustomAttribute<DataMemberAttribute>(inherit: true);
            YieldIfUnique(dataMember?.Name, aliases);

            var excelColumn = propertyInfo.GetCustomAttribute<ExcelColumnAttribute>(inherit: true);
            if (excelColumn != null) {
                YieldIfUnique(excelColumn.Name, aliases);
                foreach (string alias in excelColumn.Aliases) {
                    YieldIfUnique(alias, aliases);
                }
            }

            return aliases;
        }

        private static string CanonicalizeMemberName(string? value) {
            if (string.IsNullOrWhiteSpace(value)) {
                return string.Empty;
            }

            string text = value ?? string.Empty;
            var builder = new StringBuilder(text.Length);
            foreach (char character in text) {
                if (char.IsLetterOrDigit(character)) {
                    builder.Append(char.ToUpperInvariant(character));
                }
            }

            return builder.ToString();
        }

        private object? TryChangeType(object value, Type targetType, CultureInfo culture) {
            if (value == null) return null;

            var nullable = Nullable.GetUnderlyingType(targetType);
            var destType = nullable ?? targetType;

            var hook = _opt.TypeConverter;
            if (hook != null) {
                var (ok, v) = hook(value, destType, culture);
                if (ok) return v;
            }

            var srcType = value.GetType();
            if (targetType.IsAssignableFrom(srcType)) return value;

            return ConvertToDestinationType(value, destType, culture);
        }

        private static object? ConvertToDestinationType(object value, Type destType, CultureInfo culture) {
            try {
                if (destType == typeof(string)) {
                    return value as string ?? Convert.ToString(value, culture);
                }

                if (destType == typeof(bool)) {
                    if (value is bool boolValue) return boolValue;
                    return Convert.ToBoolean(value, culture);
                }

                if (destType == typeof(int)) {
                    if (value is int intValue) return intValue;
                    if (value is double doubleValue
                        && doubleValue >= int.MinValue
                        && doubleValue <= int.MaxValue
                        && Math.Truncate(doubleValue) == doubleValue) {
                        return (int)doubleValue;
                    }

                    return Convert.ToInt32(value, culture);
                }

                if (destType == typeof(long)) {
                    if (value is long longValue) return longValue;
                    if (value is double doubleValue
                        && doubleValue >= long.MinValue
                        && doubleValue <= long.MaxValue
                        && Math.Truncate(doubleValue) == doubleValue) {
                        return (long)doubleValue;
                    }

                    return Convert.ToInt64(value, culture);
                }

                if (destType == typeof(double)) {
                    if (value is double doubleValue) return doubleValue;
                    return Convert.ToDouble(value, culture);
                }

                if (destType == typeof(decimal)) {
                    if (value is decimal decimalValue) return decimalValue;
                    return Convert.ToDecimal(value, culture);
                }

                if (destType == typeof(DateTime)) {
                    if (value is DateTime dt) return dt;
                    if (value is double oa) return DateTime.FromOADate(oa);
                    if (DateTime.TryParse(Convert.ToString(value, culture), culture, DateTimeStyles.AssumeLocal, out var parsed)) return parsed;
                    return null;
                }

                return Convert.ChangeType(value, destType, culture);
            } catch {
                return null;
            }
        }
    }
}
