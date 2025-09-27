using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Concurrent;
using System.Diagnostics;
using System.Globalization;
using System.Threading.Tasks;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Inserts objects into the worksheet by flattening their properties into columns.
        /// </summary>
        /// <typeparam name="T">Type of objects being inserted.</typeparam>
        /// <param name="items">Collection of objects to insert.</param>
        /// <param name="includeHeaders">Whether to include column headers.</param>
        /// <param name="startRow">1-based starting row.</param>
        public void InsertObjects<T>(IEnumerable<T> items, bool includeHeaders = true, int startRow = 1) {
            if (items == null) {
                throw new ArgumentNullException(nameof(items));
            }

            var list = items.Cast<object?>().ToList();
            if (list.Count == 0) {
                return;
            }

            var flattenedItems = new List<Dictionary<string, object?>>();
            List<string> headers = new List<string>();
            HashSet<string> headerSet = new HashSet<string>();

            foreach (var item in list) {
                var dict = new Dictionary<string, object?>();
                FlattenObject(item, null, dict);
                flattenedItems.Add(dict);
                foreach (var key in dict.Keys) {
                    if (headerSet.Add(key)) {
                        headers.Add(key);
                    }
                }
            }

            List<(int Row, int Column, object Value)> cells = new List<(int Row, int Column, object Value)>();
            int row = startRow;
            if (includeHeaders) {
                for (int c = 0; c < headers.Count; c++) {
                    cells.Add((row, c + 1, headers[c]));
                }
                row++;
            }

            foreach (var dict in flattenedItems) {
                for (int c = 0; c < headers.Count; c++) {
                    object value = dict.TryGetValue(headers[c], out var entry) ? entry ?? string.Empty : string.Empty;
                    cells.Add((row, c + 1, value));
                }
                row++;
            }

            // Use the batch CellValues path with planner + execution policy
            SetCellValues(cells, null);
        }

        private static void FlattenObject(object? value, string? prefix, IDictionary<string, object?> result) {
            if (value == null) {
                if (!string.IsNullOrEmpty(prefix)) {
                    result[prefix!] = null;
                }
                return;
            }

            if (value is IDictionary dictionary) {
                foreach (DictionaryEntry entry in dictionary) {
                    string key = entry.Key?.ToString() ?? string.Empty;
                    string childPrefix = string.IsNullOrEmpty(prefix) ? key : prefix + "." + key;
                    FlattenObject(entry.Value, childPrefix, result);
                }
                return;
            }

            if (value is IEnumerable enumerable && value is not string) {
                var values = new List<string>();
                foreach (var item in enumerable) {
                    values.Add(item?.ToString() ?? string.Empty);
                }
                if (!string.IsNullOrEmpty(prefix)) {
                    result[prefix!] = string.Join(", ", values);
                }
                return;
            }

            Type type = value.GetType();
            if (type.IsPrimitive || value is string || value is decimal || value is DateTime || value is DateTimeOffset || value is Guid) {
                if (!string.IsNullOrEmpty(prefix)) {
                    result[prefix!] = value;
                }
                return;
            }

            var props = type.GetProperties().Where(p => p.CanRead);
            bool hasAny = false;
            foreach (var prop in props) {
                hasAny = true;
                string childPrefix = string.IsNullOrEmpty(prefix) ? prop.Name : prefix + "." + prop.Name;
                FlattenObject(prop.GetValue(value, null), childPrefix, result);
            }

            if (!hasAny && !string.IsNullOrEmpty(prefix)) {
                result[prefix!] = value.ToString();
            }
        }

        private class CellUpdate {
            public int Row { get; }
            public int Column { get; }
            public string Text { get; }
            public DocumentFormat.OpenXml.Spreadsheet.CellValues DataType { get; }
            public bool IsSharedString { get; }

            public CellUpdate(int row, int column, string text, DocumentFormat.OpenXml.Spreadsheet.CellValues dataType, bool isSharedString) {
                Row = row;
                Column = column;
                Text = text;
                DataType = dataType;
                IsSharedString = isSharedString;
            }
        }

        /// <summary>
        /// Sets multiple cell values in parallel without mutating the DOM concurrently.
        /// Cell data is prepared in parallel, then applied sequentially under write lock
        /// to prevent XML corruption and ensure thread safety.
        /// </summary>
        /// <param name="cells">Collection of cell coordinates and values.</param>
        public void CellValuesParallel(IEnumerable<(int Row, int Column, object Value)> cells) {
            if (cells == null) {
                throw new ArgumentNullException(nameof(cells));
            }

            var cellList = cells as IList<(int Row, int Column, object Value)> ?? cells.ToList();
            int cellCount = cellList.Count;
            bool monitor = cellCount > 5000;
            Stopwatch? prepWatch = null;
            Stopwatch? applyWatch = null;
            if (monitor) {
                prepWatch = Stopwatch.StartNew();
            }

            var bag = new ConcurrentBag<CellUpdate>();

            Parallel.ForEach(cellList, cell => {
                bag.Add(PrepareCellUpdate(cell.Row, cell.Column, cell.Value));
            });

            if (monitor && prepWatch != null) {
                prepWatch.Stop();
                applyWatch = Stopwatch.StartNew();
            }

            WriteLock(() => {
                _isBatchOperation = true;
                try {
                    foreach (var update in bag) {
                        ApplyCellUpdate(update);
                    }
                    ValidateWorksheetXml();
                } finally {
                    _isBatchOperation = false;
                }
            });

            if (monitor && applyWatch != null && prepWatch != null) {
                applyWatch.Stop();
                Debug.WriteLine($"CellValuesParallel: prepared {cellCount} cells in {prepWatch.ElapsedMilliseconds} ms, applied in {applyWatch.ElapsedMilliseconds} ms.");
            }
        }

        private CellUpdate PrepareCellUpdate(int row, int column, object value) {
            var dateTimeOffsetStrategy = _excelDocument.DateTimeOffsetWriteStrategy;
            var (cellValue, dataType) = CoerceValueHelper.Coerce(
                value,
                s => new CellValue(s),
                dateTimeOffsetStrategy);

            bool isSharedString = dataType == DocumentFormat.OpenXml.Spreadsheet.CellValues.SharedString;
            return new CellUpdate(row, column, cellValue.Text ?? string.Empty, dataType, isSharedString);
        }

        private void ApplyCellUpdate(CellUpdate update) {
            Cell cell = GetCell(update.Row, update.Column);
            if (update.IsSharedString) {
                int sharedStringIndex = _excelDocument.GetSharedStringIndex(update.Text);
                cell.CellValue = new CellValue(sharedStringIndex.ToString(CultureInfo.InvariantCulture));
                cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.SharedString;
            } else {
                cell.CellValue = new CellValue(update.Text);
                cell.DataType = update.DataType;
            }
        }

        private void ValidateWorksheetXml()
            => WorksheetIntegrityValidator.Validate(_worksheetPart, EffectiveExecution, Name);
    }
}

