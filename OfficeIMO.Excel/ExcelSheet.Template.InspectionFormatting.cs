using System.Globalization;
using System.Diagnostics.CodeAnalysis;
using System.Reflection;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Replaces {{Marker}} placeholders in text cells on this worksheet using public properties from the supplied model.
        /// Nested properties are exposed as dotted marker names, for example {{Customer.Name}}.
        /// </summary>
        [RequiresUnreferencedCode("Object-model template binding walks runtime properties, including nested values. Use the IDictionary overload in NativeAOT applications.")]
        public int ApplyTemplate(object model, IFormatProvider? provider = null, bool throwOnMissing = false) {
            if (model == null) throw new ArgumentNullException(nameof(model));
            return ApplyTemplate(ExcelTemplateBindingHelper.Create(model), provider, throwOnMissing);
        }

        /// <summary>
        /// Replaces {{Marker}} placeholders in text cells on this worksheet using public properties from the supplied model and options.
        /// Nested properties are exposed as dotted marker names, for example {{Customer.Name}}.
        /// </summary>
        [RequiresUnreferencedCode("Object-model template binding walks runtime properties, including nested values. Use the IDictionary overload in NativeAOT applications.")]
        public int ApplyTemplate(object model, ExcelTemplateOptions options) {
            if (model == null) throw new ArgumentNullException(nameof(model));
            if (options == null) throw new ArgumentNullException(nameof(options));
            return ApplyTemplateCore(ExcelTemplateBindingHelper.Create(model), options);
        }

        /// <summary>
        /// Inspects {{Marker}} placeholders on this worksheet without modifying it.
        /// </summary>
        public ExcelTemplateInspection InspectTemplate() {
            return new ExcelTemplateInspection(GetTemplateMarkers(), hasBindingInfo: false);
        }

        /// <summary>
        /// Inspects {{Marker}} placeholders on this worksheet and reports which markers are missing from the supplied values.
        /// </summary>
        public ExcelTemplateInspection InspectTemplate(IDictionary<string, object?> values) {
            if (values == null) throw new ArgumentNullException(nameof(values));
            return new ExcelTemplateInspection(GetTemplateMarkers(ExcelTemplateBindingHelper.Create(values)), hasBindingInfo: true);
        }

        /// <summary>
        /// Inspects {{Marker}} placeholders on this worksheet and reports which markers are missing from public properties on the supplied model.
        /// </summary>
        [RequiresUnreferencedCode("Object-model template inspection walks runtime properties, including nested values. Use the IDictionary overload in NativeAOT applications.")]
        public ExcelTemplateInspection InspectTemplate(object model) {
            if (model == null) throw new ArgumentNullException(nameof(model));
            return InspectTemplate(ExcelTemplateBindingHelper.Create(model));
        }

        internal IReadOnlyList<ExcelTemplateMarkerInfo> GetTemplateMarkers(IReadOnlyDictionary<string, object?>? bindings = null) {
            MaterializeDeferredDataSetImportIfNeeded();
            return Locking.ExecuteRead(_excelDocument.EnsureLock(), () => {
                var markers = new List<ExcelTemplateMarkerInfo>();
                foreach (var cell in WorksheetRoot.Descendants<Cell>()) {
                    var value = GetCellValueSnapshot(cell);
                    if (value.Value is not string text || text.IndexOf("{{", StringComparison.Ordinal) < 0) {
                        continue;
                    }

                    bool wholeCell = WholeCellTemplateMarkerRegex.IsMatch(text);
                    foreach (Match match in TemplateMarkerRegex.Matches(text)) {
                        string name = match.Groups["name"].Value;
                        string? format = match.Groups["format"].Success ? match.Groups["format"].Value.Trim() : null;
                        bool? isBound = null;
                        string? boundValueKind = null;
                        string? boundValueTypeName = null;
                        if (bindings != null) {
                            isBound = bindings.TryGetValue(name, out object? boundValue);
                            if (isBound.Value) {
                                boundValueKind = DescribeTemplateBoundValueKind(boundValue);
                                boundValueTypeName = boundValue?.GetType().Name;
                            }
                        }

                        markers.Add(new ExcelTemplateMarkerInfo(Name, cell.CellReference?.Value ?? string.Empty, name, format, text, wholeCell, isBound, boundValueKind, boundValueTypeName));
                    }
                }

                return markers;
            });
        }

        private static string DescribeTemplateBoundValueKind(object? value) {
            if (value == null) {
                return "null";
            }

            Type type = Nullable.GetUnderlyingType(value.GetType()) ?? value.GetType();
            if (value is ExcelTemplateImage) {
                return "image";
            }

            if (value is string || value is char || value is Guid) {
                return "text";
            }

            if (value is bool) {
                return "boolean";
            }

            if (type == typeof(DateTime) || type == typeof(DateTimeOffset)) {
                return "date/time";
            }

            if (type == typeof(TimeSpan)) {
                return "duration";
            }

            if (IsNumericTemplateValue(type)) {
                return "number";
            }

            return "object";
        }

        private static bool IsNumericTemplateValue(Type type) {
            type = Nullable.GetUnderlyingType(type) ?? type;
            return type == typeof(byte)
                || type == typeof(sbyte)
                || type == typeof(short)
                || type == typeof(ushort)
                || type == typeof(int)
                || type == typeof(uint)
                || type == typeof(long)
                || type == typeof(ulong)
                || type == typeof(float)
                || type == typeof(double)
                || type == typeof(decimal);
        }

        private static string FormatTemplateValue(object? value, string? format, ExcelTemplateOptions options) {
            if (!string.IsNullOrWhiteSpace(format)
                && options.Formatters.TryGetValue(format!.Trim(), out ExcelTemplateValueFormatter? formatter)) {
                return formatter(value, options.FormatProvider) ?? string.Empty;
            }

            if (value == null) {
                return string.Empty;
            }

            if (value is TimeSpan timeSpan && IsDurationFormatAlias(format)) {
                return FormatDurationText(timeSpan);
            }

            string? resolvedFormat = ResolveTemplateFormatAlias(format);
            if (value is IFormattable formattable) {
                return formattable.ToString(resolvedFormat, options.FormatProvider ?? CultureInfo.CurrentCulture) ?? string.Empty;
            }

            return Convert.ToString(value, options.FormatProvider as CultureInfo ?? CultureInfo.CurrentCulture) ?? string.Empty;
        }

        private static bool IsDurationFormatAlias(string? format) {
            if (string.IsNullOrWhiteSpace(format)) {
                return false;
            }

            string trimmed = format!.Trim();
            return string.Equals(trimmed, "duration", StringComparison.OrdinalIgnoreCase)
                || string.Equals(trimmed, "durationhours", StringComparison.OrdinalIgnoreCase)
                || string.Equals(trimmed, "elapsed", StringComparison.OrdinalIgnoreCase);
        }

        private static string FormatDurationText(TimeSpan value) {
            string sign = value < TimeSpan.Zero ? "-" : string.Empty;
            value = value.Duration();
            int totalHours = (int)Math.Floor(value.TotalHours);
            return string.Format(CultureInfo.InvariantCulture, "{0}{1}:{2:00}:{3:00}", sign, totalHours, value.Minutes, value.Seconds);
        }

        private static string? ResolveTemplateFormatAlias(string? format) {
            if (string.IsNullOrWhiteSpace(format)) {
                return null;
            }

            string trimmed = format!.Trim();
            switch (trimmed.ToLowerInvariant()) {
                case "currency":
                case "money":
                    return "C";
                case "percent":
                    return "P2";
                case "integer":
                case "int":
                    return "N0";
                case "decimal":
                case "number":
                    return "N2";
                case "date":
                    return "yyyy-MM-dd";
                case "datetime":
                    return "yyyy-MM-dd HH:mm:ss";
                case "time":
                    return "HH:mm:ss";
                case "duration":
                case "durationhours":
                case "elapsed":
                    return @"hh\:mm\:ss";
                default:
                    return trimmed;
            }
        }

        private static string? ResolveTemplateNumberFormatAlias(string? format, IFormatProvider? provider) {
            if (string.IsNullOrWhiteSpace(format)) {
                return null;
            }

            string trimmed = format!.Trim();
            switch (trimmed.ToLowerInvariant()) {
                case "currency":
                case "money":
                    return ExcelNumberFormats.Get(ExcelNumberPreset.Currency, decimals: 2, provider as CultureInfo);
                case "percent":
                    return ExcelNumberFormats.Get(ExcelNumberPreset.Percent, decimals: 2);
                case "integer":
                case "int":
                    return ExcelNumberFormats.Get(ExcelNumberPreset.Integer);
                case "decimal":
                case "number":
                    return ExcelNumberFormats.Get(ExcelNumberPreset.Decimal, decimals: 2);
                case "date":
                    return ExcelNumberFormats.Get(ExcelNumberPreset.DateShort);
                case "datetime":
                    return ExcelNumberFormats.Get(ExcelNumberPreset.DateTime);
                case "time":
                    return ExcelNumberFormats.Get(ExcelNumberPreset.Time);
                case "duration":
                case "durationhours":
                case "elapsed":
                    return ExcelNumberFormats.Get(ExcelNumberPreset.DurationHours);
                default:
                    return null;
            }
        }
    }
}
