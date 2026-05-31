using System.Globalization;
using System.Reflection;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Formats a template marker value for a named marker format.
    /// </summary>
    public delegate string ExcelTemplateValueFormatter(object? value, IFormatProvider? provider);

    /// <summary>
    /// Controls how template binding handles markers that are not supplied by the values/model.
    /// </summary>
    public enum ExcelTemplateMissingValueBehavior {
        /// <summary>Leave the marker text unchanged.</summary>
        PreserveMarker,

        /// <summary>Replace the marker with an empty string.</summary>
        EmptyString,

        /// <summary>Throw an exception when a marker is missing.</summary>
        Throw
    }

    /// <summary>
    /// Options used when applying workbook or worksheet template markers.
    /// </summary>
    public sealed class ExcelTemplateOptions {
        /// <summary>Format provider used by built-in aliases and custom formatters.</summary>
        public IFormatProvider? FormatProvider { get; set; }

        /// <summary>Throws when a marker is not supplied by the values/model. Equivalent to <see cref="ExcelTemplateMissingValueBehavior.Throw"/>.</summary>
        public bool ThrowOnMissing { get; set; }

        /// <summary>Behavior used when a marker is not supplied by the values/model.</summary>
        public ExcelTemplateMissingValueBehavior MissingValueBehavior { get; set; }

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
    /// Image value that can be bound to a whole-cell template marker.
    /// </summary>
    public sealed class ExcelTemplateImage {
        private ExcelTemplateImage(byte[]? bytes, string? url, string contentType, int widthPixels, int heightPixels, int offsetXPixels, int offsetYPixels, string? name, string? altText, bool lockAspectRatio) {
            Bytes = bytes;
            Url = url;
            ContentType = contentType;
            WidthPixels = widthPixels;
            HeightPixels = heightPixels;
            OffsetXPixels = offsetXPixels;
            OffsetYPixels = offsetYPixels;
            Name = name;
            AltText = altText;
            LockAspectRatio = lockAspectRatio;
        }

        /// <summary>Image bytes when the image is supplied directly.</summary>
        public byte[]? Bytes { get; }

        /// <summary>Remote image URL when the image should be downloaded during binding.</summary>
        public string? Url { get; }

        /// <summary>Image content type, such as image/png or image/jpeg.</summary>
        public string ContentType { get; }

        /// <summary>Image width in pixels.</summary>
        public int WidthPixels { get; }

        /// <summary>Image height in pixels.</summary>
        public int HeightPixels { get; }

        /// <summary>Horizontal pixel offset from the target cell.</summary>
        public int OffsetXPixels { get; }

        /// <summary>Vertical pixel offset from the target cell.</summary>
        public int OffsetYPixels { get; }

        /// <summary>Optional drawing name.</summary>
        public string? Name { get; }

        /// <summary>Optional alternative text description.</summary>
        public string? AltText { get; }

        /// <summary>Whether Excel should keep the picture aspect ratio locked.</summary>
        public bool LockAspectRatio { get; }

        /// <summary>
        /// Creates a template image from bytes.
        /// </summary>
        public static ExcelTemplateImage FromBytes(byte[] bytes, string contentType = "image/png", int widthPixels = 96, int heightPixels = 32, int offsetXPixels = 0, int offsetYPixels = 0, string? name = null, string? altText = null, bool lockAspectRatio = true) {
            if (bytes == null || bytes.Length == 0) throw new ArgumentException("Image bytes are required.", nameof(bytes));
            return new ExcelTemplateImage(bytes.ToArray(), null, NormalizeContentType(contentType), widthPixels, heightPixels, offsetXPixels, offsetYPixels, name, altText, lockAspectRatio);
        }

        /// <summary>
        /// Creates a template image from a stream.
        /// </summary>
        public static ExcelTemplateImage FromStream(Stream stream, string contentType = "image/png", int widthPixels = 96, int heightPixels = 32, int offsetXPixels = 0, int offsetYPixels = 0, string? name = null, string? altText = null, bool lockAspectRatio = true) {
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            using var buffer = new MemoryStream();
            stream.CopyTo(buffer);
            return FromBytes(buffer.ToArray(), contentType, widthPixels, heightPixels, offsetXPixels, offsetYPixels, name, altText, lockAspectRatio);
        }

        /// <summary>
        /// Creates a template image from a remote URL. The image is downloaded when the template is applied.
        /// </summary>
        public static ExcelTemplateImage FromUrl(string url, int widthPixels = 96, int heightPixels = 32, int offsetXPixels = 0, int offsetYPixels = 0, string? name = null, string? altText = null, bool lockAspectRatio = true) {
            if (string.IsNullOrWhiteSpace(url)) throw new ArgumentNullException(nameof(url));
            return new ExcelTemplateImage(null, url.Trim(), "image/png", widthPixels, heightPixels, offsetXPixels, offsetYPixels, name, altText, lockAspectRatio);
        }

        internal bool TryAddToSheet(ExcelSheet sheet, int row, int column) {
            if (Bytes != null) {
                sheet.AddImage(row, column, Bytes, ContentType, WidthPixels, HeightPixels, OffsetXPixels, OffsetYPixels, Name, AltText, LockAspectRatio);
                return true;
            }

            if (!string.IsNullOrWhiteSpace(Url)
                && ImageDownloader.TryFetch(Url!, timeoutSeconds: 5, maxBytes: 2_000_000, out var bytes, out var contentType)
                && bytes != null) {
                sheet.AddImage(row, column, bytes, string.IsNullOrWhiteSpace(contentType) ? ContentType : contentType!, WidthPixels, HeightPixels, OffsetXPixels, OffsetYPixels, Name, AltText, LockAspectRatio);
                return true;
            }

            return false;
        }

        private static string NormalizeContentType(string? contentType) {
            return string.IsNullOrWhiteSpace(contentType) ? "image/png" : contentType!.Trim();
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
                RenameWorkSheet(templateSheet, firstSheetName, SheetNameValidationMode.Sanitize);
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

        private sealed class ExcelTemplateSheetBinding {
            internal ExcelTemplateSheetBinding(IReadOnlyDictionary<string, object?> bindings, string? sheetName) {
                Bindings = bindings;
                SheetName = sheetName;
            }

            internal IReadOnlyDictionary<string, object?> Bindings { get; }
            internal string? SheetName { get; }
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

        internal ExcelSheet CopyTemplateWorksheet(string sheetName) {
            ExcelSheet target = _excelDocument.AddWorkSheet(sheetName, SheetNameValidationMode.Sanitize);
            Worksheet worksheet = (Worksheet)WorksheetRoot.CloneNode(true);
            RemoveRelationshipBackedTemplateCopyElements(worksheet);
            target.WorksheetRoot = worksheet;
            RewriteTemplateWorksheetLocalReferences(target.WorksheetRoot, Name, target.Name);
            CopyTemplateWorksheetTableParts(target);
            CopyTemplateWorksheetHyperlinks(target);
            CopyTemplateWorksheetDrawings(target);
            CopyTemplateWorksheetComments(target);
            CopyTemplateWorksheetScopedDefinedNames(target);
            var workbookDefinedNameMap = CopyTemplateWorksheetWorkbookScopedDefinedNames(target);
            RewriteTemplateWorkbookDefinedNameReferences(target, workbookDefinedNameMap);
            target.WorksheetRoot.Save();
            target._sheetDataCache = null;
            target._lastAccessedRow = null;
            target._lastAccessedRowIndex = 0;
            target._lastAccessedCell = null;
            target._lastAccessedCellRowIndex = 0;
            target._lastAccessedCellColumnIndex = 0;
            target.ClearHeaderCache();
            target.MarkRequiresSavePreparation();
            return target;
        }

        private static void RewriteTemplateWorksheetLocalReferences(Worksheet worksheet, string sourceSheetName, string targetSheetName) {
            if (string.Equals(sourceSheetName, targetSheetName, StringComparison.Ordinal)) {
                return;
            }

            foreach (var formula in worksheet.Descendants<CellFormula>()) {
                RewriteTemplateFormulaText(formula, sourceSheetName, targetSheetName);
            }

            foreach (var formula in worksheet.Descendants<Formula>()) {
                RewriteTemplateFormulaText(formula, sourceSheetName, targetSheetName);
            }

            foreach (var formula in worksheet.Descendants<Formula1>()) {
                RewriteTemplateFormulaText(formula, sourceSheetName, targetSheetName);
            }

            foreach (var formula in worksheet.Descendants<Formula2>()) {
                RewriteTemplateFormulaText(formula, sourceSheetName, targetSheetName);
            }

            foreach (var formula in worksheet.Descendants<DocumentFormat.OpenXml.Office.Excel.Formula>()) {
                RewriteTemplateFormulaText(formula, sourceSheetName, targetSheetName);
            }

            foreach (var hyperlink in worksheet.Descendants<Hyperlink>()) {
                string? location = hyperlink.Location?.Value;
                if (string.IsNullOrEmpty(location)) {
                    continue;
                }

                string updated = ExcelDocument.ReplaceSheetNameReferences(location!, sourceSheetName, targetSheetName);
                if (!string.Equals(updated, location, StringComparison.Ordinal)) {
                    hyperlink.Location = updated;
                }
            }
        }

        private static void RewriteTemplateTableLocalReferences(Table table, string sourceSheetName, string targetSheetName) {
            if (string.Equals(sourceSheetName, targetSheetName, StringComparison.Ordinal)) {
                return;
            }

            foreach (var formula in table.Descendants<CalculatedColumnFormula>()) {
                RewriteTemplateFormulaText(formula, sourceSheetName, targetSheetName);
            }

            foreach (var formula in table.Descendants<TotalsRowFormula>()) {
                RewriteTemplateFormulaText(formula, sourceSheetName, targetSheetName);
            }
        }

        private static void RewriteTemplateFormulaText(OpenXmlLeafTextElement formula, string sourceSheetName, string targetSheetName) {
            string? text = formula.Text;
            if (string.IsNullOrEmpty(text)) {
                return;
            }

            string updated = ExcelDocument.ReplaceSheetNameReferences(text, sourceSheetName, targetSheetName);
            if (!string.Equals(updated, text, StringComparison.Ordinal)) {
                formula.Text = updated;
            }
        }

        private void CopyTemplateWorksheetScopedDefinedNames(ExcelSheet target) {
            DefinedNames? definedNames = WorkbookRoot.DefinedNames;
            if (definedNames == null) {
                return;
            }

            ushort sourceSheetPosition = GetTemplateSheetPositionIndex(Name);
            ushort targetSheetPosition = GetTemplateSheetPositionIndex(target.Name);
            var sourceNames = definedNames.Elements<DefinedName>()
                .Where(name => name.LocalSheetId != null && name.LocalSheetId.Value == sourceSheetPosition)
                .ToList();
            if (sourceNames.Count == 0) {
                return;
            }

            foreach (DefinedName sourceName in sourceNames) {
                string? definedName = sourceName.Name?.Value;
                if (string.IsNullOrWhiteSpace(definedName)) {
                    continue;
                }

                foreach (DefinedName existing in definedNames.Elements<DefinedName>()
                    .Where(name => name.LocalSheetId != null
                        && name.LocalSheetId.Value == targetSheetPosition
                        && string.Equals(name.Name?.Value, definedName, StringComparison.OrdinalIgnoreCase))
                    .ToList()) {
                    existing.Remove();
                }

                var clone = (DefinedName)sourceName.CloneNode(true);
                clone.LocalSheetId = targetSheetPosition;
                if (!string.IsNullOrEmpty(clone.Text)) {
                    clone.Text = ExcelDocument.ReplaceSheetNameReferences(clone.Text!, Name, target.Name);
                }

                definedNames.Append(clone);
            }

            WorkbookRoot.Save();
        }

        private Dictionary<string, string> CopyTemplateWorksheetWorkbookScopedDefinedNames(ExcelSheet target) {
            var copiedNames = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            DefinedNames? definedNames = WorkbookRoot.DefinedNames;
            if (definedNames == null) {
                return copiedNames;
            }

            var existingNames = new HashSet<string>(definedNames.Elements<DefinedName>()
                .Where(name => name.LocalSheetId == null && !string.IsNullOrWhiteSpace(name.Name?.Value))
                .Select(name => name.Name!.Value!), StringComparer.OrdinalIgnoreCase);
            var sourceNames = definedNames.Elements<DefinedName>()
                .Where(name => name.LocalSheetId == null
                    && !IsTemplateBuiltInDefinedName(name.Name?.Value)
                    && !string.IsNullOrWhiteSpace(name.Name?.Value)
                    && DefinedNameReferencesTemplateSheet(name.Text, Name, target.Name))
                .ToList();
            if (sourceNames.Count == 0) {
                return copiedNames;
            }

            foreach (DefinedName sourceName in sourceNames) {
                string sourceDefinedName = sourceName.Name!.Value!;
                string targetDefinedName = CreateTemplateWorkbookDefinedName(sourceDefinedName, target.Name, existingNames);
                existingNames.Add(targetDefinedName);

                var clone = (DefinedName)sourceName.CloneNode(true);
                clone.Name = targetDefinedName;
                clone.LocalSheetId = null;
                if (!string.IsNullOrEmpty(clone.Text)) {
                    clone.Text = ExcelDocument.ReplaceSheetNameReferences(clone.Text!, Name, target.Name);
                }

                definedNames.Append(clone);
                copiedNames[sourceDefinedName] = targetDefinedName;
            }

            WorkbookRoot.Save();
            return copiedNames;
        }

        private static bool DefinedNameReferencesTemplateSheet(string? reference, string sourceSheetName, string targetSheetName) {
            if (string.IsNullOrWhiteSpace(reference)) {
                return false;
            }

            string updated = ExcelDocument.ReplaceSheetNameReferences(reference!, sourceSheetName, targetSheetName);
            return !string.Equals(updated, reference, StringComparison.Ordinal);
        }

        private static bool IsTemplateBuiltInDefinedName(string? name) {
            return !string.IsNullOrWhiteSpace(name)
                && name!.StartsWith("_xlnm.", StringComparison.OrdinalIgnoreCase);
        }

        private static string CreateTemplateWorkbookDefinedName(string sourceDefinedName, string targetSheetName, ISet<string> existingNames) {
            string baseName = Regex.Replace(sourceDefinedName + "_" + targetSheetName, @"[^A-Za-z0-9_\.]", "_");
            if (string.IsNullOrWhiteSpace(baseName) || (!char.IsLetter(baseName[0]) && baseName[0] != '_' && baseName[0] != '\\')) {
                baseName = "_" + baseName;
            }

            if (baseName.Length > 240) {
                baseName = baseName.Substring(0, 240);
            }

            string candidate = baseName;
            int suffix = 2;
            while (existingNames.Contains(candidate)) {
                string suffixText = "_" + suffix.ToString(CultureInfo.InvariantCulture);
                int maxBaseLength = Math.Max(1, 255 - suffixText.Length);
                candidate = (baseName.Length > maxBaseLength ? baseName.Substring(0, maxBaseLength) : baseName) + suffixText;
                suffix++;
            }

            return candidate;
        }

        private static void RewriteTemplateWorkbookDefinedNameReferences(ExcelSheet target, IReadOnlyDictionary<string, string> definedNameMap) {
            if (definedNameMap.Count == 0) {
                return;
            }

            RewriteDefinedNameReferences(target.WorksheetRoot, definedNameMap);

            foreach (TableDefinitionPart tablePart in target._worksheetPart.TableDefinitionParts) {
                if (tablePart.Table != null) {
                    RewriteDefinedNameReferences(tablePart.Table, definedNameMap);
                    tablePart.Table.Save();
                }
            }

            DrawingsPart? drawingsPart = target._worksheetPart.DrawingsPart;
            if (drawingsPart != null) {
                foreach (ChartPart chartPart in drawingsPart.ChartParts) {
                    if (chartPart.ChartSpace != null) {
                        RewriteDefinedNameReferences(chartPart.ChartSpace, definedNameMap);
                        chartPart.ChartSpace.Save();
                    }
                }
            }
        }

        private static void RewriteDefinedNameReferences(OpenXmlElement root, IReadOnlyDictionary<string, string> definedNameMap) {
            foreach (var formula in root.Descendants<CellFormula>()) {
                RewriteDefinedNameFormulaText(formula, definedNameMap);
            }

            foreach (var formula in root.Descendants<Formula>()) {
                RewriteDefinedNameFormulaText(formula, definedNameMap);
            }

            foreach (var formula in root.Descendants<Formula1>()) {
                RewriteDefinedNameFormulaText(formula, definedNameMap);
            }

            foreach (var formula in root.Descendants<Formula2>()) {
                RewriteDefinedNameFormulaText(formula, definedNameMap);
            }

            foreach (var formula in root.Descendants<DocumentFormat.OpenXml.Office.Excel.Formula>()) {
                RewriteDefinedNameFormulaText(formula, definedNameMap);
            }

            foreach (var formula in root.Descendants<DocumentFormat.OpenXml.Drawing.Charts.Formula>()) {
                RewriteDefinedNameFormulaText(formula, definedNameMap);
            }
        }

        private static void RewriteDefinedNameFormulaText(OpenXmlLeafTextElement formula, IReadOnlyDictionary<string, string> definedNameMap) {
            string? text = formula.Text;
            if (string.IsNullOrEmpty(text)) {
                return;
            }

            string updated = text!;
            foreach (var pair in definedNameMap) {
                updated = Regex.Replace(
                    updated,
                    @"(?<![A-Za-z0-9_\.])" + Regex.Escape(pair.Key) + @"(?![A-Za-z0-9_\.])",
                    pair.Value,
                    RegexOptions.IgnoreCase,
                    TimeSpan.FromMilliseconds(100));
            }

            if (!string.Equals(updated, text, StringComparison.Ordinal)) {
                formula.Text = updated;
            }
        }

        private ushort GetTemplateSheetPositionIndex(string sheetName) {
            var sheets = WorkbookRoot.Sheets?.OfType<Sheet>().ToList() ?? new List<Sheet>();
            for (ushort index = 0; index < sheets.Count; index++) {
                if (string.Equals(sheets[index].Name?.Value, sheetName, StringComparison.Ordinal)) {
                    return index;
                }
            }

            throw new ArgumentException($"Worksheet '{sheetName}' was not found.", nameof(sheetName));
        }

        private static void RemoveRelationshipBackedTemplateCopyElements(Worksheet worksheet) {
            foreach (var drawing in worksheet.Descendants<DocumentFormat.OpenXml.Spreadsheet.Drawing>().ToList()) {
                drawing.Remove();
            }

            foreach (var legacyDrawing in worksheet.Descendants<LegacyDrawing>().ToList()) {
                legacyDrawing.Remove();
            }

            foreach (var legacyHeaderFooterDrawing in worksheet.Descendants<LegacyDrawingHeaderFooter>().ToList()) {
                legacyHeaderFooterDrawing.Remove();
            }

            foreach (var picture in worksheet.Descendants<Picture>().ToList()) {
                picture.Remove();
            }

            foreach (var oleObjects in worksheet.Descendants<OleObjects>().ToList()) {
                oleObjects.Remove();
            }

            foreach (var controls in worksheet.Descendants<Controls>().ToList()) {
                controls.Remove();
            }

        }

        private void CopyTemplateWorksheetTableParts(ExcelSheet target) {
            TableParts? clonedTableParts = target.WorksheetRoot.GetFirstChild<TableParts>();
            if (clonedTableParts == null) {
                return;
            }

            foreach (TablePart tablePart in clonedTableParts.Elements<TablePart>().ToList()) {
                string? sourceRelationshipId = tablePart.Id?.Value;
                if (string.IsNullOrWhiteSpace(sourceRelationshipId)) {
                    tablePart.Remove();
                    continue;
                }

                if (_worksheetPart.GetPartById(sourceRelationshipId!) is not TableDefinitionPart sourceTablePart
                    || sourceTablePart.Table == null) {
                    tablePart.Remove();
                    continue;
                }

                string targetRelationshipId = GetUnusedRelationshipId(target._worksheetPart);
                TableDefinitionPart targetTablePart = target._worksheetPart.AddNewPart<TableDefinitionPart>(targetRelationshipId);
                Table clonedTable = (Table)sourceTablePart.Table.CloneNode(true);
                RewriteTemplateTableLocalReferences(clonedTable, Name, target.Name);
                clonedTable.Id = _excelDocument.AllocateTableId();

                string requestedName = clonedTable.Name?.Value ?? clonedTable.DisplayName?.Value ?? "Table";
                string tableName = EnsureValidUniqueTableName(requestedName, TableNameValidationMode.Sanitize);
                clonedTable.Name = tableName;
                clonedTable.DisplayName = tableName;
                _excelDocument.ReserveTableName(tableName);

                targetTablePart.Table = clonedTable;
                targetTablePart.Table.Save();
                tablePart.Id = targetRelationshipId;
            }

            uint count = (uint)clonedTableParts.Elements<TablePart>().Count();
            if (count == 0) {
                clonedTableParts.Remove();
            } else {
                clonedTableParts.Count = count;
            }
        }

        private static string GetUnusedRelationshipId(OpenXmlPartContainer partContainer) {
            var existing = new HashSet<string>(StringComparer.Ordinal);
            foreach (var part in partContainer.Parts) {
                if (!string.IsNullOrWhiteSpace(part.RelationshipId)) {
                    existing.Add(part.RelationshipId);
                }
            }

            foreach (var relationship in partContainer.ExternalRelationships) {
                if (!string.IsNullOrWhiteSpace(relationship.Id)) {
                    existing.Add(relationship.Id);
                }
            }

            if (partContainer is WorksheetPart worksheetPart) {
                foreach (var relationship in worksheetPart.HyperlinkRelationships) {
                    if (!string.IsNullOrWhiteSpace(relationship.Id)) {
                        existing.Add(relationship.Id);
                    }
                }
            }

            foreach (var relationship in partContainer.DataPartReferenceRelationships) {
                if (!string.IsNullOrWhiteSpace(relationship.Id)) {
                    existing.Add(relationship.Id);
                }
            }

            int index = 1;
            string id;
            do {
                id = "rId" + index.ToString(CultureInfo.InvariantCulture);
                index++;
            } while (existing.Contains(id));

            return id;
        }

        private void CopyTemplateWorksheetHyperlinks(ExcelSheet target) {
            Hyperlinks? hyperlinks = target.WorksheetRoot.GetFirstChild<Hyperlinks>();
            if (hyperlinks == null) {
                return;
            }

            foreach (Hyperlink hyperlink in hyperlinks.Elements<Hyperlink>().Where(hyperlink => hyperlink.Id != null).ToList()) {
                string? sourceRelationshipId = hyperlink.Id?.Value;
                HyperlinkRelationship? sourceRelationship = string.IsNullOrWhiteSpace(sourceRelationshipId)
                    ? null
                    : _worksheetPart.HyperlinkRelationships.FirstOrDefault(relationship =>
                        string.Equals(relationship.Id, sourceRelationshipId, StringComparison.OrdinalIgnoreCase));

                if (sourceRelationship == null) {
                    hyperlink.Remove();
                    continue;
                }

                string targetRelationshipId = GetUnusedRelationshipId(target._worksheetPart);
                target._worksheetPart.AddHyperlinkRelationship(sourceRelationship.Uri, sourceRelationship.IsExternal, targetRelationshipId);
                hyperlink.Id = targetRelationshipId;
            }

            if (!hyperlinks.Elements<Hyperlink>().Any()) {
                hyperlinks.Remove();
            }
        }

        private void CopyTemplateWorksheetDrawings(ExcelSheet target) {
            DocumentFormat.OpenXml.Spreadsheet.Drawing? sourceDrawing = WorksheetRoot.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Drawing>();
            if (sourceDrawing?.Id?.Value is not string sourceRelationshipId || string.IsNullOrWhiteSpace(sourceRelationshipId)) {
                return;
            }

            OpenXmlPart? sourcePart;
            try {
                sourcePart = _worksheetPart.GetPartById(sourceRelationshipId);
            } catch {
                return;
            }

            if (sourcePart is not DrawingsPart sourceDrawingsPart || sourceDrawingsPart.WorksheetDrawing == null) {
                return;
            }

            string targetRelationshipId = GetUnusedRelationshipId(target._worksheetPart);
            DrawingsPart targetDrawingsPart = target._worksheetPart.AddNewPart<DrawingsPart>(targetRelationshipId);
            CopyPartStream(sourceDrawingsPart, targetDrawingsPart);
            targetDrawingsPart.WorksheetDrawing ??= new DocumentFormat.OpenXml.Drawing.Spreadsheet.WorksheetDrawing();
            CopyTemplateDrawingPartRelationships(sourceDrawingsPart, targetDrawingsPart, Name, target.Name);
            targetDrawingsPart.WorksheetDrawing.Save();

            var drawing = new DocumentFormat.OpenXml.Spreadsheet.Drawing { Id = targetRelationshipId };
            LegacyDrawing? legacyDrawing = target.WorksheetRoot.GetFirstChild<LegacyDrawing>();
            LegacyDrawingHeaderFooter? legacyHeaderFooter = target.WorksheetRoot.GetFirstChild<LegacyDrawingHeaderFooter>();
            if (legacyDrawing != null) {
                target.WorksheetRoot.InsertBefore(drawing, legacyDrawing);
            } else if (legacyHeaderFooter != null) {
                target.WorksheetRoot.InsertBefore(drawing, legacyHeaderFooter);
            } else {
                target.WorksheetRoot.Append(drawing);
            }
        }

        private static void CopyTemplateDrawingPartRelationships(
            DrawingsPart sourceDrawingsPart,
            DrawingsPart targetDrawingsPart,
            string sourceSheetName,
            string targetSheetName) {
            foreach (var relationship in sourceDrawingsPart.Parts.ToList()) {
                if (relationship.OpenXmlPart is ChartPart sourceChartPart) {
                    string targetRelationshipId = GetUnusedRelationshipId(targetDrawingsPart);
                    ChartPart targetChartPart = targetDrawingsPart.AddNewPart<ChartPart>(targetRelationshipId);
                    if (sourceChartPart.ChartSpace != null) {
                        targetChartPart.ChartSpace = (DocumentFormat.OpenXml.Drawing.Charts.ChartSpace)sourceChartPart.ChartSpace.CloneNode(true);
                        RewriteChartSheetReferences(targetChartPart, sourceSheetName, targetSheetName);
                        targetChartPart.ChartSpace.Save();
                    }

                    foreach (var chartRelationship in sourceChartPart.Parts.ToList()) {
                        CopyTemplateKnownPartRelationship(chartRelationship.OpenXmlPart, targetChartPart, chartRelationship.RelationshipId);
                    }

                    RewriteDrawingRelationshipId(targetDrawingsPart.WorksheetDrawing!, relationship.RelationshipId, targetRelationshipId);
                    continue;
                }

                if (relationship.OpenXmlPart is ImagePart sourceImagePart) {
                    string targetRelationshipId = GetUnusedRelationshipId(targetDrawingsPart);
                    CopyTemplateImagePart(sourceImagePart, targetDrawingsPart, targetRelationshipId);
                    RewriteDrawingRelationshipId(targetDrawingsPart.WorksheetDrawing!, relationship.RelationshipId, targetRelationshipId);
                    continue;
                }

                if (IsTemplateDiagramPart(relationship.OpenXmlPart)) {
                    string targetRelationshipId = GetUnusedRelationshipId(targetDrawingsPart);
                    CopyTemplateKnownPartRelationship(relationship.OpenXmlPart, targetDrawingsPart, targetRelationshipId);
                    RewriteDrawingRelationshipId(targetDrawingsPart.WorksheetDrawing!, relationship.RelationshipId, targetRelationshipId);
                    continue;
                }

                CopyTemplateKnownPartRelationship(relationship.OpenXmlPart, targetDrawingsPart, relationship.RelationshipId);
            }

            CopyReferencedDrawingImages(sourceDrawingsPart, targetDrawingsPart);
        }

        private static void RewriteChartSheetReferences(ChartPart chartPart, string sourceSheetName, string targetSheetName) {
            if (chartPart.ChartSpace == null || string.Equals(sourceSheetName, targetSheetName, StringComparison.Ordinal)) {
                return;
            }

            foreach (var formula in chartPart.ChartSpace.Descendants<DocumentFormat.OpenXml.Drawing.Charts.Formula>()) {
                string? text = formula.Text;
                if (string.IsNullOrEmpty(text)) {
                    continue;
                }

                string updated = ExcelDocument.ReplaceSheetNameReferences(text, sourceSheetName, targetSheetName);
                if (!string.Equals(updated, text, StringComparison.Ordinal)) {
                    formula.Text = updated;
                }
            }
        }

        private static void CopyTemplateKnownPartRelationship(OpenXmlPart sourcePart, OpenXmlPartContainer targetContainer, string sourceRelationshipId) {
            if (sourcePart is ImagePart sourceImagePart) {
                ImagePart? targetImagePart = targetContainer switch {
                    ChartPart chartPart => chartPart.AddImagePart(sourceImagePart.ContentType, sourceRelationshipId),
                    ChartDrawingPart chartDrawingPart => chartDrawingPart.AddImagePart(sourceImagePart.ContentType, sourceRelationshipId),
                    DiagramDataPart diagramDataPart => diagramDataPart.AddImagePart(sourceImagePart.ContentType, sourceRelationshipId),
                    DiagramLayoutDefinitionPart diagramLayoutPart => diagramLayoutPart.AddImagePart(sourceImagePart.ContentType, sourceRelationshipId),
                    DiagramPersistLayoutPart diagramPersistLayoutPart => diagramPersistLayoutPart.AddImagePart(sourceImagePart.ContentType, sourceRelationshipId),
                    _ => null
                };
                if (targetImagePart == null) {
                    return;
                }

                CopyPartStream(sourceImagePart, targetImagePart);
                return;
            }

            if (sourcePart is ChartStylePart sourceChartStylePart && targetContainer is ChartPart targetChartPartForStyle) {
                ChartStylePart targetChartStylePart = targetChartPartForStyle.AddNewPart<ChartStylePart>(sourceRelationshipId);
                CopyPartStream(sourceChartStylePart, targetChartStylePart);
                return;
            }

            if (sourcePart is ChartColorStylePart sourceChartColorStylePart && targetContainer is ChartPart targetChartPartForColorStyle) {
                ChartColorStylePart targetChartColorStylePart = targetChartPartForColorStyle.AddNewPart<ChartColorStylePart>(sourceRelationshipId);
                CopyPartStream(sourceChartColorStylePart, targetChartColorStylePart);
                return;
            }

            if (sourcePart is ChartDrawingPart sourceChartDrawingPart && targetContainer is ChartPart targetChartPartForDrawing) {
                ChartDrawingPart targetChartDrawingPart = targetChartPartForDrawing.AddNewPart<ChartDrawingPart>(sourceRelationshipId);
                CopyPartStream(sourceChartDrawingPart, targetChartDrawingPart);
                CopyTemplateChildPartRelationships(sourceChartDrawingPart, targetChartDrawingPart);

                return;
            }

            if (sourcePart is DiagramColorsPart sourceDiagramColorsPart && targetContainer is DrawingsPart targetDrawingsPartForColors) {
                DiagramColorsPart targetDiagramColorsPart = targetDrawingsPartForColors.AddNewPart<DiagramColorsPart>(sourceRelationshipId);
                CopyPartStream(sourceDiagramColorsPart, targetDiagramColorsPart);
                CopyTemplateChildPartRelationships(sourceDiagramColorsPart, targetDiagramColorsPart);
                return;
            }

            if (sourcePart is DiagramDataPart sourceDiagramDataPart && targetContainer is DrawingsPart targetDrawingsPartForData) {
                DiagramDataPart targetDiagramDataPart = targetDrawingsPartForData.AddNewPart<DiagramDataPart>(sourceRelationshipId);
                CopyPartStream(sourceDiagramDataPart, targetDiagramDataPart);
                CopyTemplateChildPartRelationships(sourceDiagramDataPart, targetDiagramDataPart);
                return;
            }

            if (sourcePart is DiagramLayoutDefinitionPart sourceDiagramLayoutPart && targetContainer is DrawingsPart targetDrawingsPartForLayout) {
                DiagramLayoutDefinitionPart targetDiagramLayoutPart = targetDrawingsPartForLayout.AddNewPart<DiagramLayoutDefinitionPart>(sourceRelationshipId);
                CopyPartStream(sourceDiagramLayoutPart, targetDiagramLayoutPart);
                CopyTemplateChildPartRelationships(sourceDiagramLayoutPart, targetDiagramLayoutPart);
                return;
            }

            if (sourcePart is DiagramPersistLayoutPart sourceDiagramPersistLayoutPart && targetContainer is DrawingsPart targetDrawingsPartForPersistLayout) {
                DiagramPersistLayoutPart targetDiagramPersistLayoutPart = targetDrawingsPartForPersistLayout.AddNewPart<DiagramPersistLayoutPart>(sourceRelationshipId);
                CopyPartStream(sourceDiagramPersistLayoutPart, targetDiagramPersistLayoutPart);
                CopyTemplateChildPartRelationships(sourceDiagramPersistLayoutPart, targetDiagramPersistLayoutPart);
                return;
            }

            if (sourcePart is DiagramStylePart sourceDiagramStylePart && targetContainer is DrawingsPart targetDrawingsPartForStyle) {
                DiagramStylePart targetDiagramStylePart = targetDrawingsPartForStyle.AddNewPart<DiagramStylePart>(sourceRelationshipId);
                CopyPartStream(sourceDiagramStylePart, targetDiagramStylePart);
                CopyTemplateChildPartRelationships(sourceDiagramStylePart, targetDiagramStylePart);
                return;
            }

            if (sourcePart is EmbeddedPackagePart sourceEmbeddedPackagePart && targetContainer is ChartPart targetChartPartForEmbeddedPackage) {
                EmbeddedPackagePart targetEmbeddedPackagePart = targetChartPartForEmbeddedPackage.AddEmbeddedPackagePart(sourceEmbeddedPackagePart.ContentType, sourceRelationshipId);
                CopyPartStream(sourceEmbeddedPackagePart, targetEmbeddedPackagePart);
            }
        }

        private static bool IsTemplateDiagramPart(OpenXmlPart sourcePart) {
            return sourcePart is DiagramColorsPart
                || sourcePart is DiagramDataPart
                || sourcePart is DiagramLayoutDefinitionPart
                || sourcePart is DiagramPersistLayoutPart
                || sourcePart is DiagramStylePart;
        }

        private static void CopyTemplateChildPartRelationships(OpenXmlPart sourcePart, OpenXmlPart targetPart) {
            foreach (var relationship in sourcePart.Parts.ToList()) {
                CopyTemplateKnownPartRelationship(relationship.OpenXmlPart, targetPart, relationship.RelationshipId);
            }
        }

        private static void CopyTemplateImagePart(ImagePart sourceImagePart, DrawingsPart targetDrawingsPart, string targetRelationshipId) {
            ImagePart targetImagePart = targetDrawingsPart.AddImagePart(sourceImagePart.ContentType, targetRelationshipId);
            CopyPartStream(sourceImagePart, targetImagePart);
        }

        private static void CopyReferencedDrawingImages(DrawingsPart sourceDrawingsPart, DrawingsPart targetDrawingsPart) {
            if (targetDrawingsPart.WorksheetDrawing == null) {
                return;
            }

            foreach (var blip in targetDrawingsPart.WorksheetDrawing.Descendants<DocumentFormat.OpenXml.Drawing.Blip>()) {
                string? sourceRelationshipId = blip.Embed?.Value;
                if (string.IsNullOrWhiteSpace(sourceRelationshipId)) {
                    continue;
                }

                if (TryGetPartById(targetDrawingsPart, sourceRelationshipId!) != null) {
                    continue;
                }

                if (TryGetPartById(sourceDrawingsPart, sourceRelationshipId!) is not ImagePart sourceImagePart) {
                    continue;
                }

                string targetRelationshipId = GetUnusedRelationshipId(targetDrawingsPart);
                CopyTemplateImagePart(sourceImagePart, targetDrawingsPart, targetRelationshipId);
                blip.Embed = targetRelationshipId;
            }
        }

        private static OpenXmlPart? TryGetPartById(OpenXmlPartContainer container, string relationshipId) {
            try {
                return container.GetPartById(relationshipId);
            } catch {
                return null;
            }
        }

        private static void CopyPartStream(OpenXmlPart sourcePart, OpenXmlPart targetPart) {
            using (Stream sourceStream = sourcePart.GetStream(FileMode.Open, FileAccess.Read))
            using (Stream targetStream = targetPart.GetStream(FileMode.Create, FileAccess.Write)) {
                sourceStream.CopyTo(targetStream);
            }
        }

        private static void RewriteDrawingRelationshipId(OpenXmlElement root, string oldRelationshipId, string newRelationshipId) {
            foreach (var element in root.Descendants<OpenXmlElement>()) {
                foreach (var attribute in element.GetAttributes()) {
                    if (string.Equals(attribute.NamespaceUri, "http://schemas.openxmlformats.org/officeDocument/2006/relationships", StringComparison.Ordinal)
                        && string.Equals(attribute.Value, oldRelationshipId, StringComparison.Ordinal)) {
                        element.SetAttribute(new OpenXmlAttribute(attribute.Prefix, attribute.LocalName, attribute.NamespaceUri, newRelationshipId));
                    }
                }
            }
        }

        private void CopyTemplateWorksheetComments(ExcelSheet target) {
            WorksheetCommentsPart? sourceCommentsPart = _worksheetPart.WorksheetCommentsPart;
            if (sourceCommentsPart?.Comments?.CommentList?.Elements<Comment>().Any() != true) {
                return;
            }

            string commentsRelationshipId = GetUnusedRelationshipId(target._worksheetPart);
            WorksheetCommentsPart targetCommentsPart = target._worksheetPart.AddNewPart<WorksheetCommentsPart>(commentsRelationshipId);
            targetCommentsPart.Comments = (Comments)sourceCommentsPart.Comments.CloneNode(true);
            targetCommentsPart.Comments.Save();

            LegacyDrawing? sourceLegacyDrawing = WorksheetRoot.GetFirstChild<LegacyDrawing>();
            if (sourceLegacyDrawing?.Id?.Value is not string sourceLegacyRelationshipId
                || string.IsNullOrWhiteSpace(sourceLegacyRelationshipId)) {
                return;
            }

            OpenXmlPart? sourceLegacyPart;
            try {
                sourceLegacyPart = _worksheetPart.GetPartById(sourceLegacyRelationshipId);
            } catch {
                return;
            }

            if (sourceLegacyPart is not VmlDrawingPart sourceVmlPart) {
                return;
            }

            string vmlRelationshipId = GetUnusedRelationshipId(target._worksheetPart);
            VmlDrawingPart targetVmlPart = target._worksheetPart.AddNewPart<VmlDrawingPart>(vmlRelationshipId);
            using (Stream sourceStream = sourceVmlPart.GetStream(FileMode.Open, FileAccess.Read))
            using (Stream targetStream = targetVmlPart.GetStream(FileMode.Create, FileAccess.Write)) {
                sourceStream.CopyTo(targetStream);
            }

            var legacyDrawing = new LegacyDrawing { Id = vmlRelationshipId };
            LegacyDrawingHeaderFooter? legacyHeaderFooter = target.WorksheetRoot.GetFirstChild<LegacyDrawingHeaderFooter>();
            if (legacyHeaderFooter != null) {
                target.WorksheetRoot.InsertBefore(legacyDrawing, legacyHeaderFooter);
            } else {
                target.WorksheetRoot.Append(legacyDrawing);
            }
        }

        /// <summary>
        /// Includes or removes an optional worksheet row block. When included, markers in the block are bound with the supplied values.
        /// When removed, following worksheet rows are shifted up.
        /// </summary>
        /// <param name="firstRow">1-based first row in the optional block.</param>
        /// <param name="rowCount">Number of rows in the optional block.</param>
        /// <param name="include">True to keep and bind the block; false to remove it.</param>
        /// <param name="values">Values used when the block is included.</param>
        /// <param name="options">Optional template binding options.</param>
        public int ApplyTemplateOptionalRows(int firstRow, int rowCount, bool include, IDictionary<string, object?> values, ExcelTemplateOptions? options = null) {
            if (values == null) throw new ArgumentNullException(nameof(values));
            return ApplyTemplateOptionalRowsCore(firstRow, rowCount, include, ExcelTemplateBindingHelper.Create(values), options ?? new ExcelTemplateOptions());
        }

        /// <summary>
        /// Includes or removes an optional worksheet row block. When included, markers in the block are bound from public properties on the supplied model.
        /// When removed, following worksheet rows are shifted up.
        /// </summary>
        /// <param name="firstRow">1-based first row in the optional block.</param>
        /// <param name="rowCount">Number of rows in the optional block.</param>
        /// <param name="include">True to keep and bind the block; false to remove it.</param>
        /// <param name="model">Model used when the block is included.</param>
        /// <param name="options">Optional template binding options.</param>
        public int ApplyTemplateOptionalRows(int firstRow, int rowCount, bool include, object model, ExcelTemplateOptions? options = null) {
            if (model == null) throw new ArgumentNullException(nameof(model));
            return ApplyTemplateOptionalRowsCore(firstRow, rowCount, include, ExcelTemplateBindingHelper.Create(model), options ?? new ExcelTemplateOptions());
        }

        /// <summary>
        /// Removes an optional worksheet row block and shifts following worksheet rows up.
        /// </summary>
        /// <param name="firstRow">1-based first row in the optional block.</param>
        /// <param name="rowCount">Number of rows in the optional block.</param>
        public int RemoveTemplateOptionalRows(int firstRow, int rowCount) {
            return ApplyTemplateOptionalRowsCore(firstRow, rowCount, include: false, bindings: null, new ExcelTemplateOptions());
        }

        private int ApplyTemplateRowsCore(int templateRow, IReadOnlyList<IReadOnlyDictionary<string, object?>> rowBindings, ExcelTemplateOptions options) {
            if (templateRow <= 0) throw new ArgumentOutOfRangeException(nameof(templateRow));
            if (rowBindings.Count == 0) return 0;

            int replacements = 0;
            WriteLockConditional(() => {
                var bounds = GetTemplateRowBounds(templateRow);
                if (bounds == null) {
                    return;
                }

                var snapshot = CaptureRow(templateRow, bounds.Value.FirstColumn, bounds.Value.LastColumn, bounds.Value.FirstColumn);
                if (rowBindings.Count > 1) {
                    ShiftRowsDown(templateRow + 1, rowBindings.Count - 1);
                }

                for (int index = 0; index < rowBindings.Count; index++) {
                    int targetRow = templateRow + index;
                    var rowMap = targetRow == templateRow
                        ? new Dictionary<int, int>()
                        : new Dictionary<int, int> { [templateRow] = targetRow };
                    WriteRowSnapshot(targetRow, bounds.Value.FirstColumn, bounds.Value.LastColumn, snapshot, rowMap, targetRow - templateRow);
                    replacements += ApplyTemplateCellsCore(rowBindings[index], options, targetRow);
                }

                WorksheetRoot.Save();
            });

            return replacements;
        }

        private int ApplyTemplateOptionalRowsCore(int firstRow, int rowCount, bool include, IReadOnlyDictionary<string, object?>? bindings, ExcelTemplateOptions options) {
            if (firstRow <= 0) throw new ArgumentOutOfRangeException(nameof(firstRow));
            if (rowCount <= 0) throw new ArgumentOutOfRangeException(nameof(rowCount));

            int replacements = 0;
            WriteLockConditional(() => {
                if (include) {
                    if (bindings != null) {
                        int lastRow = firstRow + rowCount - 1;
                        for (int row = firstRow; row <= lastRow; row++) {
                            replacements += ApplyTemplateCellsCore(bindings, options, row);
                        }
                    }
                } else {
                    RemoveRowsAndShiftUp(firstRow, rowCount);
                }

                WorksheetRoot.Save();
            });

            return replacements;
        }

        private int ApplyTemplateCellsCore(IReadOnlyDictionary<string, object?> bindings, ExcelTemplateOptions options, int? rowFilter) {
            int replacements = 0;
            foreach (var cell in WorksheetRoot.Descendants<Cell>().ToList()) {
                var reference = A1.ParseCellRef(cell.CellReference?.Value ?? string.Empty);
                if (rowFilter.HasValue && reference.Row != rowFilter.Value) {
                    continue;
                }

                var value = GetCellValueSnapshot(cell);
                if (value.Value is not string text || text.IndexOf("{{", StringComparison.Ordinal) < 0) {
                    continue;
                }

                var wholeMarker = WholeCellTemplateMarkerRegex.Match(text);
                if (wholeMarker.Success) {
                    string marker = wholeMarker.Groups["name"].Value;
                    if (!bindings.TryGetValue(marker, out object? replacement)) {
                        if (ShouldThrowOnMissing(options)) {
                            ThrowMissingMarker(marker);
                        }

                        if (options.MissingValueBehavior == ExcelTemplateMissingValueBehavior.EmptyString
                            && reference.Row > 0
                            && reference.Col > 0) {
                            CellValueCore(reference.Row, reference.Col, string.Empty);
                            replacements++;
                        }

                        continue;
                    }

                    string? format = wholeMarker.Groups["format"].Success ? wholeMarker.Groups["format"].Value.Trim() : null;
                    if (replacement is ExcelTemplateImage templateImage && reference.Row > 0 && reference.Col > 0) {
                        using (Locking.EnterNoLockScope()) {
                            if (!templateImage.TryAddToSheet(this, reference.Row, reference.Col)) {
                                throw new InvalidOperationException($"Template marker '{marker}' image could not be loaded.");
                            }
                        }

                        CellValueCore(reference.Row, reference.Col, string.Empty);
                        replacements++;
                        continue;
                    }

                    string? numberFormat = ResolveTemplateNumberFormatAlias(format, options.FormatProvider);
                    if (numberFormat != null && replacement != null && reference.Row > 0 && reference.Col > 0) {
                        CellValueCore(reference.Row, reference.Col, replacement);
                        FormatCellCore(reference.Row, reference.Col, numberFormat);
                        replacements++;
                        continue;
                    }
                }

                int cellReplacements = 0;
                string replaced = TemplateMarkerRegex.Replace(text, match => {
                    string marker = match.Groups["name"].Value;
                    if (!bindings.TryGetValue(marker, out object? replacement)) {
                        if (ShouldThrowOnMissing(options)) {
                            ThrowMissingMarker(marker);
                        }

                        if (options.MissingValueBehavior == ExcelTemplateMissingValueBehavior.EmptyString) {
                            cellReplacements++;
                            return string.Empty;
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

            return replacements;
        }

        private (int FirstColumn, int LastColumn)? GetTemplateRowBounds(int templateRow) {
            var sheetData = WorksheetRoot.GetFirstChild<SheetData>();
            var row = sheetData?.Elements<Row>().FirstOrDefault(item => item.RowIndex?.Value == (uint)templateRow);
            if (row == null) {
                return null;
            }

            int firstColumn = int.MaxValue;
            int lastColumn = 0;
            foreach (var cell in row.Elements<Cell>()) {
                if (cell.CellReference?.Value is not string reference || reference.Length == 0) {
                    continue;
                }

                int column = GetColumnIndex(reference);
                if (column <= 0) {
                    continue;
                }

                firstColumn = Math.Min(firstColumn, column);
                lastColumn = Math.Max(lastColumn, column);
            }

            return lastColumn == 0 ? null : (firstColumn, lastColumn);
        }

        private void ShiftRowsDown(int firstRow, int count) {
            if (count <= 0) {
                return;
            }

            var sheetData = WorksheetRoot.GetFirstChild<SheetData>();
            if (sheetData == null) {
                return;
            }

            foreach (var row in sheetData.Elements<Row>()
                .Where(item => item.RowIndex?.Value >= (uint)firstRow)
                .OrderByDescending(item => item.RowIndex?.Value ?? 0U)
                .ToList()) {
                int newRowIndex = (int)(row.RowIndex!.Value + (uint)count);
                row.RowIndex = (uint)newRowIndex;
                foreach (var cell in row.Elements<Cell>()) {
                    if (cell.CellReference?.Value is not string reference || reference.Length == 0) {
                        continue;
                    }

                    int column = GetColumnIndex(reference);
                    if (column > 0) {
                        cell.CellReference = BuildCellReference(newRowIndex, column);
                    }
                }
            }

            RewriteWorksheetFormulaReferences(firstRow, count);
            RemapShiftedRowMetadata(firstRow, count);
            ShiftMergeCellsRows(firstRow, count);

            _lastAccessedRow = null;
            _lastAccessedRowIndex = 0;
            _lastAccessedCell = null;
            _lastAccessedCellRowIndex = 0;
            _lastAccessedCellColumnIndex = 0;
            ClearHeaderCache();
        }

        private void RemoveRowsAndShiftUp(int firstRow, int count) {
            if (count <= 0) {
                return;
            }

            var sheetData = WorksheetRoot.GetFirstChild<SheetData>();
            if (sheetData == null) {
                return;
            }

            int lastRemovedRow = firstRow + count - 1;
            foreach (var row in sheetData.Elements<Row>().ToList()) {
                if (row.RowIndex == null) {
                    continue;
                }

                int rowIndex = checked((int)row.RowIndex.Value);
                if (rowIndex >= firstRow && rowIndex <= lastRemovedRow) {
                    row.Remove();
                    continue;
                }

                if (rowIndex > lastRemovedRow) {
                    int newRowIndex = rowIndex - count;
                    row.RowIndex = (uint)newRowIndex;
                    foreach (var cell in row.Elements<Cell>()) {
                        if (cell.CellReference?.Value is not string reference || reference.Length == 0) {
                            continue;
                        }

                        int column = GetColumnIndex(reference);
                        if (column > 0) {
                            cell.CellReference = BuildCellReference(newRowIndex, column);
                        }
                    }
                }
            }

            RewriteDeletedWorksheetFormulaReferences(firstRow, lastRemovedRow, -count);
            RemapDeletedRowMetadata(firstRow, lastRemovedRow, -count);
            ShiftMergeCellsRows(firstRow, -count, lastRemovedRow);

            _lastAccessedRow = null;
            _lastAccessedRowIndex = 0;
            _lastAccessedCell = null;
            _lastAccessedCellRowIndex = 0;
            _lastAccessedCellColumnIndex = 0;
            ClearHeaderCache();
        }

        private static bool ShouldThrowOnMissing(ExcelTemplateOptions options) {
            return options.ThrowOnMissing || options.MissingValueBehavior == ExcelTemplateMissingValueBehavior.Throw;
        }

        private static void ThrowMissingMarker(string marker) {
            throw new InvalidOperationException($"Template marker '{marker}' was not supplied.");
        }

        private void ShiftMergeCellsRows(int firstAffectedRow, int delta, int? lastDeletedRow = null) {
            var merges = WorksheetRoot.GetFirstChild<MergeCells>();
            if (merges == null || delta == 0) {
                return;
            }

            uint count = 0;
            foreach (var merge in merges.Elements<MergeCell>().ToList()) {
                if (merge.Reference?.Value is not string reference
                    || !TryParseReference(reference, out var bounds)) {
                    count++;
                    continue;
                }

                if (!TryRemapShiftedReferenceRows(bounds, firstAffectedRow, delta, lastDeletedRow, out var remappedBounds)) {
                    count++;
                    continue;
                }

                if (remappedBounds == null) {
                    merge.Remove();
                    continue;
                }

                merge.Reference = ToReference(remappedBounds.Value.r1, remappedBounds.Value.c1, remappedBounds.Value.r2, remappedBounds.Value.c2);
                count++;
            }

            merges.Count = count;
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
                || type == typeof(Guid)
                || type == typeof(ExcelTemplateImage);
        }
    }
}
