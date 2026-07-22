using System.Globalization;
using System.Diagnostics.CodeAnalysis;
using System.Reflection;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
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
                MaterializeDeferredDataSetImportIfNeeded();
                replacements = ApplyTemplateCellsCore(bindings, options, rowFilter: null);
                WorksheetRoot.Save();
            });

            return replacements;
        }

        /// <summary>
        /// Repeats a single template row for each supplied row value dictionary, inserting additional worksheet rows as needed.
        /// </summary>
        /// <param name="templateRow">1-based row number containing template markers.</param>
        /// <param name="rows">Row value dictionaries. Each dictionary is bound to one copied row.</param>
        /// <param name="options">Optional template binding options.</param>
        public int ApplyTemplateRows(int templateRow, IEnumerable<IDictionary<string, object?>> rows, ExcelTemplateOptions? options = null) {
            if (rows == null) throw new ArgumentNullException(nameof(rows));
            var bindings = rows.Select(ExcelTemplateBindingHelper.Create).Cast<IReadOnlyDictionary<string, object?>>().ToList();
            return ApplyTemplateRowsCore(templateRow, bindings, options ?? new ExcelTemplateOptions());
        }

        /// <summary>
        /// Repeats a single template row for each supplied row model, inserting additional worksheet rows as needed.
        /// Public properties are exposed as marker names, including nested dotted names.
        /// </summary>
        /// <param name="templateRow">1-based row number containing template markers.</param>
        /// <param name="rows">Row models. Each model is bound to one copied row.</param>
        /// <param name="options">Optional template binding options.</param>
        [RequiresUnreferencedCode("Object-model template binding walks runtime properties, including nested values. Use the IDictionary overload in NativeAOT applications.")]
        public int ApplyTemplateRows<T>(int templateRow, IEnumerable<T> rows, ExcelTemplateOptions? options = null) {
            if (rows == null) throw new ArgumentNullException(nameof(rows));
            var bindings = rows
                .Select(row => row == null
                    ? new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase)
                    : ExcelTemplateBindingHelper.Create(row))
                .Cast<IReadOnlyDictionary<string, object?>>()
                .ToList();
            return ApplyTemplateRowsCore(templateRow, bindings, options ?? new ExcelTemplateOptions());
        }
    }
}
