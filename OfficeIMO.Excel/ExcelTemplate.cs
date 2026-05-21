using System.Globalization;
using System.Reflection;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Formats a template marker value for a named marker format.
    /// </summary>
    public delegate string ExcelTemplateValueFormatter(object? value, IFormatProvider? provider);

    /// <summary>
    /// Options used when applying workbook or worksheet template markers.
    /// </summary>
    public sealed class ExcelTemplateOptions {
        /// <summary>Format provider used by built-in aliases and custom formatters.</summary>
        public IFormatProvider? FormatProvider { get; set; }

        /// <summary>Throws when a marker is not supplied by the values/model.</summary>
        public bool ThrowOnMissing { get; set; }

        /// <summary>Named custom formatters, keyed by marker format such as "upper" in {{Name:upper}}.</summary>
        public IDictionary<string, ExcelTemplateValueFormatter> Formatters { get; } =
            new Dictionary<string, ExcelTemplateValueFormatter>(StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// Adds or replaces a named custom formatter and returns this options instance.
        /// </summary>
        public ExcelTemplateOptions AddFormatter(string name, ExcelTemplateValueFormatter formatter) {
            if (string.IsNullOrWhiteSpace(name)) throw new ArgumentNullException(nameof(name));
            Formatters[name.Trim()] = formatter ?? throw new ArgumentNullException(nameof(formatter));
            return this;
        }

        internal static ExcelTemplateOptions Create(IFormatProvider? provider, bool throwOnMissing) {
            return new ExcelTemplateOptions {
                FormatProvider = provider,
                ThrowOnMissing = throwOnMissing
            };
        }
    }

    /// <summary>
    /// Simple workbook template helpers for replacing {{Marker}} placeholders in text cells.
    /// </summary>
    public partial class ExcelDocument {
        /// <summary>
        /// Replaces {{Marker}} placeholders across all worksheets using the supplied values.
        /// </summary>
        public int ApplyTemplate(IDictionary<string, object?> values, IFormatProvider? provider = null, bool throwOnMissing = false) {
            if (values == null) throw new ArgumentNullException(nameof(values));
            return ApplyTemplateCore(ExcelTemplateBindingHelper.Create(values), ExcelTemplateOptions.Create(provider, throwOnMissing));
        }

        /// <summary>
        /// Replaces {{Marker}} placeholders across all worksheets using the supplied values and options.
        /// </summary>
        public int ApplyTemplate(IDictionary<string, object?> values, ExcelTemplateOptions options) {
            if (values == null) throw new ArgumentNullException(nameof(values));
            if (options == null) throw new ArgumentNullException(nameof(options));
            return ApplyTemplateCore(ExcelTemplateBindingHelper.Create(values), options);
        }

        private int ApplyTemplateCore(IReadOnlyDictionary<string, object?> bindings, ExcelTemplateOptions options) {
            int replacements = 0;
            foreach (var sheet in Sheets) {
                replacements += sheet.ApplyTemplateCore(bindings, options);
            }

            return replacements;
        }

        /// <summary>
        /// Replaces {{Marker}} placeholders across all worksheets using public properties from the supplied model.
        /// Nested properties are exposed as dotted marker names, for example {{Customer.Name}}.
        /// </summary>
        public int ApplyTemplate(object model, IFormatProvider? provider = null, bool throwOnMissing = false) {
            if (model == null) throw new ArgumentNullException(nameof(model));
            return ApplyTemplateCore(ExcelTemplateBindingHelper.Create(model), ExcelTemplateOptions.Create(provider, throwOnMissing));
        }

        /// <summary>
        /// Replaces {{Marker}} placeholders across all worksheets using public properties from the supplied model and options.
        /// Nested properties are exposed as dotted marker names, for example {{Customer.Name}}.
        /// </summary>
        public int ApplyTemplate(object model, ExcelTemplateOptions options) {
            if (model == null) throw new ArgumentNullException(nameof(model));
            if (options == null) throw new ArgumentNullException(nameof(options));
            return ApplyTemplateCore(ExcelTemplateBindingHelper.Create(model), options);
        }

        /// <summary>
        /// Inspects {{Marker}} placeholders across all worksheets without modifying the workbook.
        /// </summary>
        public ExcelTemplateInspection InspectTemplate() {
            var markers = new List<ExcelTemplateMarkerInfo>();
            foreach (var sheet in Sheets) {
                markers.AddRange(sheet.GetTemplateMarkers());
            }

            return new ExcelTemplateInspection(markers, hasBindingInfo: false);
        }

        /// <summary>
        /// Inspects {{Marker}} placeholders across all worksheets and reports which markers are missing from the supplied values.
        /// </summary>
        public ExcelTemplateInspection InspectTemplate(IDictionary<string, object?> values) {
            if (values == null) throw new ArgumentNullException(nameof(values));

            var bindings = ExcelTemplateBindingHelper.Create(values);
            var markers = new List<ExcelTemplateMarkerInfo>();
            foreach (var sheet in Sheets) {
                markers.AddRange(sheet.GetTemplateMarkers(bindings));
            }

            return new ExcelTemplateInspection(markers, hasBindingInfo: true);
        }

        /// <summary>
        /// Inspects {{Marker}} placeholders across all worksheets and reports which markers are missing from public properties on the supplied model.
        /// </summary>
        public ExcelTemplateInspection InspectTemplate(object model) {
            if (model == null) throw new ArgumentNullException(nameof(model));
            return InspectTemplate(ExcelTemplateBindingHelper.Create(model));
        }
    }

    public partial class ExcelSheet {
        private static readonly Regex TemplateMarkerRegex = new Regex(
            @"\{\{\s*(?<name>[A-Za-z0-9_.-]+)(?:\s*:\s*(?<format>[^}]+?))?\s*\}\}",
            RegexOptions.Compiled,
            TimeSpan.FromMilliseconds(100));

        private static readonly Regex WholeCellTemplateMarkerRegex = new Regex(
            @"^\s*\{\{\s*(?<name>[A-Za-z0-9_.-]+)(?:\s*:\s*(?<format>[^}]+?))?\s*\}\}\s*$",
            RegexOptions.Compiled,
            TimeSpan.FromMilliseconds(100));

        /// <summary>
        /// Replaces {{Marker}} placeholders in text cells on this worksheet using the supplied values.
        /// </summary>
        public int ApplyTemplate(IDictionary<string, object?> values, IFormatProvider? provider = null, bool throwOnMissing = false) {
            if (values == null) throw new ArgumentNullException(nameof(values));
            return ApplyTemplateCore(ExcelTemplateBindingHelper.Create(values), ExcelTemplateOptions.Create(provider, throwOnMissing));
        }

        /// <summary>
        /// Replaces {{Marker}} placeholders in text cells on this worksheet using the supplied values and options.
        /// </summary>
        public int ApplyTemplate(IDictionary<string, object?> values, ExcelTemplateOptions options) {
            if (values == null) throw new ArgumentNullException(nameof(values));
            if (options == null) throw new ArgumentNullException(nameof(options));
            return ApplyTemplateCore(ExcelTemplateBindingHelper.Create(values), options);
        }

        internal int ApplyTemplateCore(IReadOnlyDictionary<string, object?> bindings, ExcelTemplateOptions options) {
            if (bindings == null) throw new ArgumentNullException(nameof(bindings));
            if (options == null) throw new ArgumentNullException(nameof(options));

            int replacements = 0;
            WriteLock(() => {
                foreach (var cell in WorksheetRoot.Descendants<Cell>().ToList()) {
                    var value = GetCellValueSnapshot(cell);
                    if (value.Value is not string text || text.IndexOf("{{", StringComparison.Ordinal) < 0) {
                        continue;
                    }

                    var wholeMarker = WholeCellTemplateMarkerRegex.Match(text);
                    if (wholeMarker.Success) {
                        string marker = wholeMarker.Groups["name"].Value;
                        if (!bindings.TryGetValue(marker, out object? replacement)) {
                            if (options.ThrowOnMissing) {
                                throw new InvalidOperationException($"Template marker '{marker}' was not supplied.");
                            }

                            continue;
                        }

                        string? format = wholeMarker.Groups["format"].Success ? wholeMarker.Groups["format"].Value.Trim() : null;
                        string? numberFormat = ResolveTemplateNumberFormatAlias(format, options.FormatProvider);
                        if (numberFormat != null && replacement != null) {
                            var reference = A1.ParseCellRef(cell.CellReference?.Value ?? string.Empty);
                            if (reference.Row > 0 && reference.Col > 0) {
                                CellValueCore(reference.Row, reference.Col, replacement);
                                FormatCellCore(reference.Row, reference.Col, numberFormat);
                                replacements++;
                                continue;
                            }
                        }
                    }

                    int cellReplacements = 0;
                    string replaced = TemplateMarkerRegex.Replace(text, match => {
                        string marker = match.Groups["name"].Value;
                        if (!bindings.TryGetValue(marker, out object? replacement)) {
                            if (options.ThrowOnMissing) {
                                throw new InvalidOperationException($"Template marker '{marker}' was not supplied.");
                            }

                            return match.Value;
                        }

                        cellReplacements++;
                        string? format = match.Groups["format"].Success ? match.Groups["format"].Value.Trim() : null;
                        return FormatTemplateValue(replacement, format, options);
                    });

                    if (cellReplacements == 0 || string.Equals(replaced, text, StringComparison.Ordinal)) {
                        continue;
                    }

                    cell.CellFormula = null;
                    cell.CellValue = null;
                    cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.InlineString;
                    cell.InlineString = new InlineString(new Text(Utilities.ExcelSanitizer.SanitizeString(replaced)));
                    replacements += cellReplacements;
                }

                WorksheetRoot.Save();
            });

            return replacements;
        }

        /// <summary>
        /// Replaces {{Marker}} placeholders in text cells on this worksheet using public properties from the supplied model.
        /// Nested properties are exposed as dotted marker names, for example {{Customer.Name}}.
        /// </summary>
        public int ApplyTemplate(object model, IFormatProvider? provider = null, bool throwOnMissing = false) {
            if (model == null) throw new ArgumentNullException(nameof(model));
            return ApplyTemplate(ExcelTemplateBindingHelper.Create(model), provider, throwOnMissing);
        }

        /// <summary>
        /// Replaces {{Marker}} placeholders in text cells on this worksheet using public properties from the supplied model and options.
        /// Nested properties are exposed as dotted marker names, for example {{Customer.Name}}.
        /// </summary>
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
        public ExcelTemplateInspection InspectTemplate(object model) {
            if (model == null) throw new ArgumentNullException(nameof(model));
            return InspectTemplate(ExcelTemplateBindingHelper.Create(model));
        }

        internal IReadOnlyList<ExcelTemplateMarkerInfo> GetTemplateMarkers(IReadOnlyDictionary<string, object?>? bindings = null) {
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
                        bool? isBound = bindings == null ? null : bindings.ContainsKey(name);
                        markers.Add(new ExcelTemplateMarkerInfo(Name, cell.CellReference?.Value ?? string.Empty, name, format, text, wholeCell, isBound));
                    }
                }

                return markers;
            });
        }

        private static string FormatTemplateValue(object? value, string? format, ExcelTemplateOptions options) {
            if (!string.IsNullOrWhiteSpace(format)
                && options.Formatters.TryGetValue(format!.Trim(), out ExcelTemplateValueFormatter? formatter)) {
                return formatter(value, options.FormatProvider) ?? string.Empty;
            }

            if (value == null) {
                return string.Empty;
            }

            string? resolvedFormat = ResolveTemplateFormatAlias(format);
            if (value is IFormattable formattable) {
                return formattable.ToString(resolvedFormat, options.FormatProvider ?? CultureInfo.CurrentCulture) ?? string.Empty;
            }

            return Convert.ToString(value, options.FormatProvider as CultureInfo ?? CultureInfo.CurrentCulture) ?? string.Empty;
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
                default:
                    return null;
            }
        }
    }

    internal static class ExcelTemplateBindingHelper {
        internal static Dictionary<string, object?> Create(IDictionary<string, object?> values) {
            return new Dictionary<string, object?>(values, StringComparer.OrdinalIgnoreCase);
        }

        internal static Dictionary<string, object?> Create(object model) {
            if (model is IDictionary<string, object?> objectDictionary) {
                return Create(objectDictionary);
            }

            var values = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
            AddBindings(values, prefix: null, model, depth: 0);
            return values;
        }

        private static void AddBindings(IDictionary<string, object?> values, string? prefix, object? source, int depth) {
            if (source == null || depth > 8) {
                return;
            }

            Type type = source.GetType();
            if (IsSimpleValue(type)) {
                if (prefix != null) {
                    values[prefix] = source;
                }
                return;
            }

            foreach (PropertyInfo property in type.GetProperties(BindingFlags.Instance | BindingFlags.Public)) {
                if (property.GetIndexParameters().Length != 0) {
                    continue;
                }

                object? value;
                try {
                    value = property.GetValue(source);
                } catch (TargetInvocationException) {
                    continue;
                }

                string key = prefix == null ? property.Name : prefix + "." + property.Name;
                values[key] = value;

                if (value != null && !IsSimpleValue(value.GetType())) {
                    AddBindings(values, key, value, depth + 1);
                }
            }
        }

        private static bool IsSimpleValue(Type type) {
            type = Nullable.GetUnderlyingType(type) ?? type;
            return type.IsPrimitive
                || type.IsEnum
                || type == typeof(string)
                || type == typeof(decimal)
                || type == typeof(DateTime)
                || type == typeof(DateTimeOffset)
                || type == typeof(TimeSpan)
                || type == typeof(Guid);
        }
    }
}
