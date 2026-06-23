using System.Globalization;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Creates a grouped subtotal summary block from a contiguous worksheet data range.
        /// </summary>
        /// <param name="options">Subtotal generation options.</param>
        /// <returns>Information about the generated summary block.</returns>
        public ExcelSubtotalResult AddSubtotalSummary(ExcelSubtotalOptions options) {
            if (options == null) throw new ArgumentNullException(nameof(options));
            ValidateSubtotalOptions(options);

            _excelDocument.MaterializeDeferredDataSetImport();

            var groups = GetContiguousSubtotalGroups(options).ToArray();
            if (groups.Length == 0) {
                throw new InvalidOperationException("No data rows were available for subtotal summary generation.");
            }
            ValidateSubtotalOutputBounds(options, groups.Length);

            int summaryStartRow = options.SummaryStartRow ?? checked(options.DataEndRow + 2);
            int currentRow = summaryStartRow;
            int firstSummaryColumn = Math.Min(options.GroupColumn, options.ValueColumns.Min());
            int lastSummaryColumn = Math.Max(options.GroupColumn, options.ValueColumns.Max());
            int subtotalCode = GetSubtotalFunctionCode(options.Function);
            var results = new List<ExcelSubtotalGroupResult>(groups.Length);

            if (options.IncludeHeader) {
                WriteSubtotalHeaderRow(options, currentRow);
                currentRow++;
            }

            foreach (var group in groups) {
                WriteSubtotalGroupRow(options, currentRow, group, subtotalCode);
                if (options.OutlineDetailRows) {
                    GroupRows(group.StartRow, group.EndRow, options.OutlineLevel, collapsed: false, hidden: options.HideDetailRows);
                }

                results.Add(new ExcelSubtotalGroupResult(group.Key, group.StartRow, group.EndRow, currentRow));
                currentRow++;
            }

            bool grandTotalWritten = false;
            if (options.IncludeGrandTotal) {
                WriteSubtotalGrandTotalRow(options, currentRow, subtotalCode);
                grandTotalWritten = true;
                currentRow++;
            }

            int summaryEndRow = currentRow - 1;
            string summaryRange = $"{A1.CellReference(summaryStartRow, firstSummaryColumn)}:{A1.CellReference(summaryEndRow, lastSummaryColumn)}";
            return new ExcelSubtotalResult(summaryRange, summaryStartRow, summaryEndRow, results, grandTotalWritten);
        }

        private void WriteSubtotalHeaderRow(ExcelSubtotalOptions options, int row) {
            string groupHeader = TryGetCellText(options.HeaderRow, options.GroupColumn, out string headerText) && !string.IsNullOrWhiteSpace(headerText)
                ? headerText
                : "Group";
            CellValue(row, options.GroupColumn, groupHeader);
            CellBold(row, options.GroupColumn, true);

            foreach (int column in options.ValueColumns) {
                string valueHeader = TryGetCellText(options.HeaderRow, column, out string valueText) && !string.IsNullOrWhiteSpace(valueText)
                    ? valueText
                    : A1.ColumnIndexToLetters(column);
                CellValue(row, column, valueHeader);
                CellBold(row, column, true);
            }
        }

        private void WriteSubtotalGroupRow(ExcelSubtotalOptions options, int row, SubtotalGroup group, int subtotalCode) {
            string label = string.Concat(group.Key, options.LabelSuffix ?? string.Empty);
            CellValue(row, options.GroupColumn, label);
            CellBold(row, options.GroupColumn, true);

            foreach (int column in options.ValueColumns) {
                string columnName = A1.ColumnIndexToLetters(column);
                CellFormula(row, column, string.Format(CultureInfo.InvariantCulture, "SUBTOTAL({0},{1}{2}:{1}{3})", subtotalCode, columnName, group.StartRow, group.EndRow));
                CellBold(row, column, true);
            }
        }

        private void WriteSubtotalGrandTotalRow(ExcelSubtotalOptions options, int row, int subtotalCode) {
            CellValue(row, options.GroupColumn, options.GrandTotalLabel);
            CellBold(row, options.GroupColumn, true);

            foreach (int column in options.ValueColumns) {
                string columnName = A1.ColumnIndexToLetters(column);
                CellFormula(row, column, string.Format(CultureInfo.InvariantCulture, "SUBTOTAL({0},{1}{2}:{1}{3})", subtotalCode, columnName, options.DataStartRow, options.DataEndRow));
                CellBold(row, column, true);
            }
        }

        private IEnumerable<SubtotalGroup> GetContiguousSubtotalGroups(ExcelSubtotalOptions options) {
            string? currentKey = null;
            int currentStart = 0;

            for (int row = options.DataStartRow; row <= options.DataEndRow; row++) {
                string key = GetSubtotalGroupKey(options, row);
                if (currentKey == null) {
                    currentKey = key;
                    currentStart = row;
                    continue;
                }

                if (string.Equals(currentKey, key, StringComparison.Ordinal)) {
                    continue;
                }

                yield return new SubtotalGroup(currentKey, currentStart, row - 1);
                currentKey = key;
                currentStart = row;
            }

            if (currentKey != null) {
                yield return new SubtotalGroup(currentKey, currentStart, options.DataEndRow);
            }
        }

        private string GetSubtotalGroupKey(ExcelSubtotalOptions options, int row) {
            if (TryGetCellText(row, options.GroupColumn, out string text) && !string.IsNullOrWhiteSpace(text)) {
                return text;
            }

            var value = GetCellValueSnapshot(row, options.GroupColumn).Value;
            string? converted = Convert.ToString(value, CultureInfo.InvariantCulture);
            return string.IsNullOrWhiteSpace(converted) ? options.BlankGroupLabel : converted!;
        }

        private static void ValidateSubtotalOptions(ExcelSubtotalOptions options) {
            if (options.HeaderRow <= 0) throw new ArgumentOutOfRangeException(nameof(options.HeaderRow), "Header row must be 1 or greater.");
            if (options.DataStartRow <= 0) throw new ArgumentOutOfRangeException(nameof(options.DataStartRow), "Data start row must be 1 or greater.");
            if (options.DataEndRow < options.DataStartRow) throw new ArgumentOutOfRangeException(nameof(options.DataEndRow), "Data end row must be greater than or equal to the data start row.");
            if (options.GroupColumn <= 0) throw new ArgumentOutOfRangeException(nameof(options.GroupColumn), "Group column must be 1 or greater.");
            if (options.HeaderRow > A1.MaxRows || options.DataStartRow > A1.MaxRows || options.DataEndRow > A1.MaxRows) throw new ArgumentOutOfRangeException(nameof(options.DataEndRow), "Subtotal source rows must not exceed the Excel row limit.");
            if (options.GroupColumn > A1.MaxColumns) throw new ArgumentOutOfRangeException(nameof(options.GroupColumn), "Group column must not exceed the Excel column limit.");
            if (options.ValueColumns == null || options.ValueColumns.Count == 0) throw new ArgumentException("At least one value column is required.", nameof(options.ValueColumns));
            if (options.ValueColumns.Any(column => column <= 0)) throw new ArgumentOutOfRangeException(nameof(options.ValueColumns), "Value columns must be 1 or greater.");
            if (options.ValueColumns.Any(column => column > A1.MaxColumns)) throw new ArgumentOutOfRangeException(nameof(options.ValueColumns), "Value columns must not exceed the Excel column limit.");
            if (options.OutlineLevel < 1 || options.OutlineLevel > 7) throw new ArgumentOutOfRangeException(nameof(options.OutlineLevel), "Excel outline level must be between 1 and 7.");
            int summaryStartRow = options.SummaryStartRow ?? checked(options.DataEndRow + 2);
            if (summaryStartRow <= options.DataEndRow) throw new ArgumentOutOfRangeException(nameof(options.SummaryStartRow), "Summary start row must be below the source data range.");
            if (summaryStartRow > A1.MaxRows) throw new ArgumentOutOfRangeException(nameof(options.SummaryStartRow), "Summary start row must not exceed the Excel row limit.");
        }

        private static void ValidateSubtotalOutputBounds(ExcelSubtotalOptions options, int groupCount) {
            int summaryStartRow = options.SummaryStartRow ?? checked(options.DataEndRow + 2);
            int rowCount = groupCount
                + (options.IncludeHeader ? 1 : 0)
                + (options.IncludeGrandTotal ? 1 : 0);
            int summaryEndRow = checked(summaryStartRow + rowCount - 1);
            if (summaryEndRow > A1.MaxRows) {
                throw new ArgumentOutOfRangeException(nameof(options.SummaryStartRow), "Subtotal summary output must not exceed the Excel row limit.");
            }
        }

        private static int GetSubtotalFunctionCode(ExcelSubtotalFunction function) {
            return function switch {
                ExcelSubtotalFunction.Average => 1,
                ExcelSubtotalFunction.Count => 2,
                ExcelSubtotalFunction.CountNonBlank => 3,
                ExcelSubtotalFunction.Max => 4,
                ExcelSubtotalFunction.Min => 5,
                ExcelSubtotalFunction.Sum => 9,
                _ => throw new ArgumentOutOfRangeException(nameof(function), function, null)
            };
        }

        private readonly struct SubtotalGroup {
            internal SubtotalGroup(string key, int startRow, int endRow) {
                Key = key;
                StartRow = startRow;
                EndRow = endRow;
            }

            internal string Key { get; }

            internal int StartRow { get; }

            internal int EndRow { get; }
        }
    }
}
