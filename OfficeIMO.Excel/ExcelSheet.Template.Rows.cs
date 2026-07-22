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
        [RequiresUnreferencedCode("Object-model template binding walks runtime properties, including nested values. Use the IDictionary overload in NativeAOT applications.")]
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
            if ((long)templateRow + rowBindings.Count - 1L > A1.MaxRows) {
                throw new ArgumentOutOfRangeException(nameof(rowBindings), "Template row expansion must fit inside Excel's worksheet row limit.");
            }

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
            if ((long)firstRow + rowCount - 1L > A1.MaxRows) throw new ArgumentOutOfRangeException(nameof(rowCount), "Template optional row block must fit inside Excel's worksheet row limit.");

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

            uint maxShiftedRow = sheetData.Elements<Row>()
                .Where(item => item.RowIndex?.Value >= (uint)firstRow)
                .Select(item => item.RowIndex?.Value ?? 0U)
                .DefaultIfEmpty(0U)
                .Max();
            if (maxShiftedRow > 0U && (long)maxShiftedRow + count > A1.MaxRows) {
                throw new InvalidOperationException("Template row expansion would move worksheet rows beyond Excel's row limit.");
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
    }
}
