using System.Globalization;
using System.Diagnostics.CodeAnalysis;
using System.Reflection;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
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
        /// Repeats a template worksheet for each supplied value dictionary and applies markers to each generated sheet.
        /// </summary>
        /// <param name="templateSheetName">Worksheet to use as the sheet template.</param>
        /// <param name="sheets">Per-sheet value dictionaries.</param>
        /// <param name="sheetNameSelector">Optional generated sheet name selector. The index is zero-based.</param>
        /// <param name="options">Optional template binding options.</param>
        /// <returns>Total marker replacements across all generated sheets.</returns>
        public int ApplyTemplateSheets(
            string templateSheetName,
            IEnumerable<IDictionary<string, object?>> sheets,
            Func<IDictionary<string, object?>, int, string>? sheetNameSelector = null,
            ExcelTemplateOptions? options = null) {
            if (sheets == null) throw new ArgumentNullException(nameof(sheets));
            var items = sheets
                .Select((sheet, index) => new ExcelTemplateSheetBinding(
                    ExcelTemplateBindingHelper.Create(sheet),
                    sheetNameSelector?.Invoke(sheet, index)))
                .ToList();
            return ApplyTemplateSheetsCore(templateSheetName, items, options ?? new ExcelTemplateOptions());
        }

        /// <summary>
        /// Repeats a template worksheet for each supplied model and applies markers to each generated sheet.
        /// Public properties are exposed as marker names, including nested dotted names.
        /// </summary>
        /// <param name="templateSheetName">Worksheet to use as the sheet template.</param>
        /// <param name="models">Per-sheet models.</param>
        /// <param name="sheetNameSelector">Optional generated sheet name selector. The index is zero-based.</param>
        /// <param name="options">Optional template binding options.</param>
        /// <returns>Total marker replacements across all generated sheets.</returns>
        [RequiresUnreferencedCode("Object-model template binding walks runtime properties, including nested values. Use the IDictionary overload in NativeAOT applications.")]
        public int ApplyTemplateSheets<T>(
            string templateSheetName,
            IEnumerable<T> models,
            Func<T, int, string>? sheetNameSelector = null,
            ExcelTemplateOptions? options = null) {
            if (models == null) throw new ArgumentNullException(nameof(models));
            var items = models
                .Select((model, index) => new ExcelTemplateSheetBinding(
                    model == null
                        ? new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase)
                        : ExcelTemplateBindingHelper.Create(model),
                    sheetNameSelector?.Invoke(model, index)))
                .ToList();
            return ApplyTemplateSheetsCore(templateSheetName, items, options ?? new ExcelTemplateOptions());
        }

        private int ApplyTemplateSheetsCore(string templateSheetName, IReadOnlyList<ExcelTemplateSheetBinding> items, ExcelTemplateOptions options) {
            if (string.IsNullOrWhiteSpace(templateSheetName)) throw new ArgumentNullException(nameof(templateSheetName));
            if (items.Count == 0) return 0;

            ExcelSheet? templateSheet = Sheets.FirstOrDefault(sheet => string.Equals(sheet.Name, templateSheetName, StringComparison.OrdinalIgnoreCase));
            if (templateSheet == null) {
                throw new ArgumentException($"Worksheet '{templateSheetName}' was not found.", nameof(templateSheetName));
            }

            var generatedSheets = new List<ExcelSheet>(items.Count) { templateSheet };
            for (int index = 1; index < items.Count; index++) {
                string generatedName = ResolveTemplateSheetName(templateSheetName, items[index].SheetName, index);
                generatedSheets.Add(templateSheet.CopyTemplateWorksheet(generatedName));
            }

            string firstSheetName = ResolveTemplateSheetName(templateSheetName, items[0].SheetName, 0);
            if (!string.Equals(templateSheet.Name, firstSheetName, StringComparison.Ordinal)) {
                RenameWorksheet(templateSheet, firstSheetName, SheetNameValidationMode.Sanitize);
            }

            int replacements = 0;
            for (int index = 0; index < generatedSheets.Count; index++) {
                replacements += generatedSheets[index].ApplyTemplateCore(items[index].Bindings, options);
            }

            return replacements;
        }

        private static string ResolveTemplateSheetName(string templateSheetName, string? requestedName, int index) {
            if (!string.IsNullOrWhiteSpace(requestedName)) {
                return requestedName!.Trim();
            }

            return index == 0
                ? templateSheetName
                : templateSheetName + " " + (index + 1).ToString(CultureInfo.InvariantCulture);
        }

        /// <summary>
        /// Replaces {{Marker}} placeholders across all worksheets using public properties from the supplied model.
        /// Nested properties are exposed as dotted marker names, for example {{Customer.Name}}.
        /// </summary>
        [RequiresUnreferencedCode("Object-model template binding walks runtime properties, including nested values. Use the IDictionary overload in NativeAOT applications.")]
        public int ApplyTemplate(object model, IFormatProvider? provider = null, bool throwOnMissing = false) {
            if (model == null) throw new ArgumentNullException(nameof(model));
            return ApplyTemplateCore(ExcelTemplateBindingHelper.Create(model), ExcelTemplateOptions.Create(provider, throwOnMissing));
        }

        /// <summary>
        /// Replaces {{Marker}} placeholders across all worksheets using public properties from the supplied model and options.
        /// Nested properties are exposed as dotted marker names, for example {{Customer.Name}}.
        /// </summary>
        [RequiresUnreferencedCode("Object-model template binding walks runtime properties, including nested values. Use the IDictionary overload in NativeAOT applications.")]
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
        [RequiresUnreferencedCode("Object-model template inspection walks runtime properties, including nested values. Use the IDictionary overload in NativeAOT applications.")]
        public ExcelTemplateInspection InspectTemplate(object model) {
            if (model == null) throw new ArgumentNullException(nameof(model));
            return InspectTemplate(ExcelTemplateBindingHelper.Create(model));
        }

        private sealed class ExcelTemplateSheetBinding {
            internal ExcelTemplateSheetBinding(IReadOnlyDictionary<string, object?> bindings, string? sheetName) {
                Bindings = bindings;
                SheetName = sheetName;
            }

            internal IReadOnlyDictionary<string, object?> Bindings { get; }
            internal string? SheetName { get; }
        }
    }
}
