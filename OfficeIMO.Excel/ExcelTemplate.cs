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
                || type == typeof(Guid);
        }
    }
}
