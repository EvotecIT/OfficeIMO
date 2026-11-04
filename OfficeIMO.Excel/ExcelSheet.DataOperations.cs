using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System.IO;
using System.Globalization;
using System.Text;

namespace OfficeIMO.Excel {
    // Data operations: autofilter, sorting (values-only), find/replace, and validation.
    // Kept small and focused; advanced formatting lives in dedicated files.
    public partial class ExcelSheet {
        // -------- AutoFilter --------

        /// <summary>Returns the AutoFilter on the range, keeping its criteria when it already exists.</summary>
        internal ExcelAutoFilter EnsureAutoFilter(string range) {
            if (string.IsNullOrWhiteSpace(range)) throw new ArgumentNullException(nameof(range));
            if (_excelDocument.TryApplyDirectWorksheetAutoFilterMetadata(this, range)) {
                return new ExcelAutoFilter(this, range);
            }

            bool exists = Locking.ExecuteRead(_excelDocument.EnsureLock(), () => TryFindAutoFilter(range, out _, out _));
            return exists ? new ExcelAutoFilter(this, range) : AddAutoFilterCore(range);
        }

        private bool TryFindAutoFilter(string range, out AutoFilter? autoFilter, out Table? tableOwner) {
            AutoFilter? worksheetFilter = WorksheetRoot.GetFirstChild<AutoFilter>();
            if (worksheetFilter != null && AutoFilterRangesMatch(worksheetFilter.Reference?.Value, range)) {
                autoFilter = worksheetFilter;
                tableOwner = null;
                return true;
            }

            foreach (var tablePart in _worksheetPart.TableDefinitionParts) {
                Table? table = tablePart.Table;
                AutoFilter? tableFilter = table?.GetFirstChild<AutoFilter>();
                if (table != null && tableFilter != null
                    && (AutoFilterRangesMatch(table.Reference?.Value, range) || AutoFilterRangesMatch(tableFilter.Reference?.Value, range))) {
                    autoFilter = tableFilter;
                    tableOwner = table;
                    return true;
                }
            }

            autoFilter = null;
            tableOwner = null;
            return false;
        }

        /// <summary>Criteria and removals edit the worksheet, so a deferred direct export materializes first, as AddAutoFilter with criteria did.</summary>
        private void MaterializeDeferredAutoFilterWorkbook() {
            if (!_excelDocument.IsMaterializingDeferredDataSetImport) _excelDocument.MaterializeDeferredDataSetImport();
        }

        /// <summary>Removes the worksheet or table AutoFilter that covers the range.</summary>
        internal bool RemoveAutoFilter(string range) {
            MaterializeDeferredAutoFilterWorkbook();
            bool removed = false;
            WriteLock(() => {
                if (!TryFindAutoFilter(range, out AutoFilter? autoFilter, out Table? tableOwner)) return;
                autoFilter!.Remove();
                DeferAutoFilterOwner(tableOwner);
                removed = true;
            });
            return removed;
        }

        /// <summary>Removes the criteria of one zero-based AutoFilter column offset.</summary>
        internal bool ClearAutoFilterColumn(string range, uint columnOffset) {
            MaterializeDeferredAutoFilterWorkbook();
            bool removed = false;
            WriteLock(() => {
                if (!TryFindAutoFilter(range, out AutoFilter? autoFilter, out Table? tableOwner)) return;
                FilterColumn? column = autoFilter!.Elements<FilterColumn>()
                    .FirstOrDefault(item => item.ColumnId?.Value == columnOffset);
                if (column == null) return;
                column.Remove();
                DeferAutoFilterOwner(tableOwner);
                removed = true;
            });
            return removed;
        }

        /// <summary>Resolves a header in the first row of the AutoFilter range to its zero-based column offset.</summary>
        internal uint ResolveAutoFilterColumnOffset(string range, string header) {
            if (string.IsNullOrWhiteSpace(header)) throw new ArgumentNullException(nameof(header));
            var (r1, c1, _, c2) = ExcelA1.ParseRange(range);
            string wanted = ExcelHeaderNameHelper.NormalizeHeader(header, true);
            for (int column = c1; column <= c2; column++) {
                if (TryGetCellText(r1, column, out string text)
                    && string.Equals(ExcelHeaderNameHelper.NormalizeHeader(text, true), wanted, StringComparison.OrdinalIgnoreCase)) {
                    return (uint)(column - c1);
                }
            }

            throw new ArgumentException($"Header '{header}' was not found in the first row of AutoFilter range {range}.", nameof(header));
        }

        internal void ApplyAutoFilterValueCriteria(string range, uint columnId, IEnumerable<string> values) {
            MaterializeDeferredAutoFilterWorkbook();
            if (string.IsNullOrWhiteSpace(range)) throw new ArgumentNullException(nameof(range));
            ArgumentNullException.ThrowIfNull(values);
            range = ValidateAutoFilterColumnOffset(range, columnId);
            List<string> filterValues = BuildDistinctFilterValues(values);

            WriteLock(() => {
                AutoFilter autoFilter = GetOrReplaceAutoFilter(WorksheetRoot, range, out Table? tableOwner);
                FilterColumn? existingColumn = autoFilter.Elements<FilterColumn>().FirstOrDefault(fc => fc.ColumnId?.Value == columnId);
                if (existingColumn != null && FilterColumnMatchesValues(existingColumn, filterValues)) return;
                existingColumn?.Remove();

                var filterColumn = new FilterColumn { ColumnId = columnId };
                var filters = new Filters();
                foreach (string value in filterValues) {
                    filters.Append(new Filter { Val = value });
                }
                filterColumn.Append(filters);
                autoFilter.Append(filterColumn);
                DeferAutoFilterOwner(tableOwner);
            });
        }

        private static List<string> BuildDistinctFilterValues(IEnumerable<string> values) {
            var result = new List<string>();
            HashSet<string>? seen = null;
            foreach (var value in values) {
                if (value == null) continue;
                seen ??= new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                if (seen.Add(value)) {
                    result.Add(value);
                }
            }

            return result;
        }

        private static bool FilterColumnMatchesValues(FilterColumn filterColumn, IReadOnlyList<string> values) {
            var filtersNode = filterColumn.GetFirstChild<Filters>();
            if (filtersNode == null || filterColumn.GetFirstChild<CustomFilters>() != null) {
                return false;
            }

            int index = 0;
            foreach (var filter in filtersNode.Elements<Filter>()) {
                if ((uint)index >= (uint)values.Count || !string.Equals(filter.Val?.Value, values[index], StringComparison.Ordinal)) {
                    return false;
                }

                index++;
            }

            return index == values.Count;
        }

        // -------- Find/Replace --------

        /// <summary>
        /// Finds the first cell text that contains the specified value. Returns the A1 address or null.
        /// Searches values rendered as text (shared strings, inline strings, numbers as invariant strings).
        /// </summary>
        public string? FindFirst(string text) {
            if (string.IsNullOrEmpty(text)) return null;
            MaterializeDeferredDataSetImportIfNeeded();
            // Text-mutating paths must clear this cache before rendered cell text changes.
            bool canUseCache = !_hasWorksheetMutations;
            if (canUseCache && TryGetFindFirstCache(text, out string? cachedAddress)) {
                return cachedAddress;
            }

            var ws = WorksheetRoot;
            var sd = ws.GetFirstChild<SheetData>();
            if (sd == null) {
                SetFindFirstCacheIfAllowed(null);
                return null;
            }

            var sharedStringCache = BuildCellTextSharedStringSnapshot();
            var sharedStringMatches = sharedStringCache.FindIndexesContaining(text, StringComparison.OrdinalIgnoreCase);
            foreach (var row in sd.Elements<Row>()) {
                foreach (var cell in row.Elements<Cell>()) {
                    if (TryGetSharedStringCellIndex(cell, out int sharedStringIndex)) {
                        if (sharedStringMatches != null && sharedStringMatches.Contains(sharedStringIndex)) {
                            string? address = cell.CellReference?.Value;
                            SetFindFirstCacheIfAllowed(address);
                            return address;
                        }

                        continue;
                    }

                    var t = GetCellText(cell);
                    if (!string.IsNullOrEmpty(t) && t.IndexOf(text, StringComparison.OrdinalIgnoreCase) >= 0) {
                        string? address = cell.CellReference?.Value;
                        SetFindFirstCacheIfAllowed(address);
                        return address;
                    }
                }
            }

            SetFindFirstCacheIfAllowed(null);
            return null;

            void SetFindFirstCacheIfAllowed(string? address) {
                if (canUseCache) {
                    SetFindFirstCache(text, address);
                }
            }
        }

        /// <summary>
        /// Replaces all occurrences of <paramref name="oldText"/> with <paramref name="newText"/> in string cells.
        /// Returns the number of replacements performed.
        /// </summary>
        public int ReplaceAll(string oldText, string newText) {
            if (string.IsNullOrEmpty(oldText)) return 0;
            int count = 0;
            WriteLock(() => {
                MaterializeDeferredDataSetImportIfNeeded();
                var ws = WorksheetRoot;
                var sd = ws.GetFirstChild<SheetData>();
                if (sd == null) return;

                var sharedStringCache = BuildCellTextSharedStringSnapshot();
                var sharedStringMatches = sharedStringCache.FindIndexesContaining(oldText, StringComparison.OrdinalIgnoreCase);
                int replacementCapacity = sharedStringMatches?.Count ?? 0;
                var replacements = replacementCapacity > 0
                    ? new List<(Cell Cell, int TextIndex)>(replacementCapacity)
                    : new List<(Cell Cell, int TextIndex)>();
                var distinctReplacementTexts = replacementCapacity > 0
                    ? new List<string>(replacementCapacity)
                    : new List<string>();
                var distinctReplacementLookup = replacementCapacity > 0
                    ? new Dictionary<string, int>(replacementCapacity, StringComparer.Ordinal)
                    : new Dictionary<string, int>(StringComparer.Ordinal);
                Dictionary<int, int>? sharedStringReplacementIndexes = null;
                if (sharedStringMatches != null) {
                    sharedStringReplacementIndexes = new Dictionary<int, int>(sharedStringMatches.Count);
                    foreach (int sharedStringIndex in sharedStringMatches) {
                        string? current = sharedStringCache.Get(sharedStringIndex);
                        if (string.IsNullOrEmpty(current)) {
                            continue;
                        }

                        string replaced = ReplaceIgnoreCase(current!, oldText, newText);
                        int replacementTextIndex = GetOrAddReplacementTextIndex(
                            replaced,
                            distinctReplacementTexts,
                            distinctReplacementLookup,
                            nameof(newText));
                        sharedStringReplacementIndexes.Add(sharedStringIndex, replacementTextIndex);
                    }
                }

                foreach (var row in sd.Elements<Row>()) {
                    foreach (var cell in row.Elements<Cell>()) {
                        string? current;
                        bool currentContainsOldText;
                        if (TryGetSharedStringCellIndex(cell, out int sharedStringIndex)) {
                            if (sharedStringReplacementIndexes == null
                                || !sharedStringReplacementIndexes.TryGetValue(sharedStringIndex, out int replacementTextIndex)) {
                                continue;
                            }

                            replacements.Add((cell, replacementTextIndex));
                            continue;
                        } else {
                            if (!TryGetReplaceableCellText(cell, out current)) {
                                continue;
                            }

                            currentContainsOldText = !string.IsNullOrEmpty(current)
                                && current!.IndexOf(oldText, StringComparison.OrdinalIgnoreCase) >= 0;
                        }

                        if (string.IsNullOrEmpty(current)) continue;
                        string currentText = current!;
                        if (currentContainsOldText) {
                            var replaced = ReplaceIgnoreCase(currentText, oldText, newText);
                            int replacementTextIndex = GetOrAddReplacementTextIndex(
                                replaced,
                                distinctReplacementTexts,
                                distinctReplacementLookup,
                                nameof(newText));

                            replacements.Add((cell, replacementTextIndex));
                        }
                    }
                }
                if (replacements.Count > 0) {
                    var replacementIndexes = _excelDocument.GetSharedStringIndexArray(distinctReplacementTexts, assumeDistinct: true);
                    foreach (var replacement in replacements) {
                        string replacementText = distinctReplacementTexts[replacement.TextIndex];
                        SetExistingCellSharedStringValue(replacement.Cell, replacementText, replacementIndexes[replacement.TextIndex]);
                    }

                    count = replacements.Count;
                    ClearHeaderCache();
                }
            });
            return count;
        }

        private static int GetOrAddReplacementTextIndex(
            string replaced,
            List<string> distinctReplacementTexts,
            Dictionary<string, int> distinctReplacementLookup,
            string paramName) {
            if (!distinctReplacementLookup.TryGetValue(replaced, out int replacementTextIndex)) {
                CoerceValueHelper.ValidateSharedStringLength(replaced, paramName);
                replacementTextIndex = distinctReplacementTexts.Count;
                distinctReplacementTexts.Add(replaced);
                distinctReplacementLookup.Add(replaced, replacementTextIndex);
            }

            return replacementTextIndex;
        }

        private static bool TryGetSharedStringCellIndex(Cell cell, out int index) {
            index = 0;
            return cell.DataType?.Value == DocumentFormat.OpenXml.Spreadsheet.CellValues.SharedString
                && TryParseCellTextSharedStringIndex(cell.CellValue?.InnerText ?? cell.InnerText, out index);
        }

        private static bool TryGetReplaceableCellText(Cell cell, out string? text) {
            var dataType = cell.DataType?.Value;
            if (dataType == DocumentFormat.OpenXml.Spreadsheet.CellValues.InlineString
                || dataType == DocumentFormat.OpenXml.Spreadsheet.CellValues.String) {
                text = cell.DataType?.Value == DocumentFormat.OpenXml.Spreadsheet.CellValues.InlineString
                    ? ExtractReplaceableInlineString(cell)
                    : cell.CellValue?.InnerText ?? string.Empty;
                return true;
            }

            if (cell.InlineString != null) {
                text = ExtractReplaceableInlineString(cell);
                return true;
            }

            text = null;
            return false;
        }

        private static string ExtractReplaceableInlineString(Cell cell) {
            var inline = cell.InlineString;
            if (inline == null) {
                return string.Empty;
            }

            if (inline.Text != null) {
                return inline.Text.Text ?? string.Empty;
            }

            string? first = null;
            StringBuilder? builder = null;
            foreach (var run in inline.Elements<Run>()) {
                string value = run.Text?.Text ?? string.Empty;
                if (builder != null) {
                    builder.Append(value);
                } else if (first == null) {
                    first = value;
                } else {
                    builder = new StringBuilder(first.Length + value.Length);
                    builder.Append(first);
                    builder.Append(value);
                }
            }

            return builder?.ToString() ?? first ?? string.Empty;
        }

        private static string ReplaceIgnoreCase(string input, string oldValue, string newValue) {
            if (string.IsNullOrEmpty(input) || string.IsNullOrEmpty(oldValue)) return input;
            int prev = 0;
            var sb = new StringBuilder(input.Length);
            while (true) {
                int idx = input.IndexOf(oldValue, prev, StringComparison.OrdinalIgnoreCase);
                if (idx < 0) break;
                sb.Append(input, prev, idx - prev);
                sb.Append(newValue);
                prev = idx + oldValue.Length;
            }
            sb.Append(input, prev, input.Length - prev);
            return sb.ToString();
        }

        internal void ApplyAutoFilterCustomCriteria(
            string range,
            uint columnId,
            bool matchAll,
            IReadOnlyList<(FilterOperatorValues Operator, string Value)> conditions) {
            MaterializeDeferredAutoFilterWorkbook();
            if (string.IsNullOrWhiteSpace(range)) throw new ArgumentNullException(nameof(range));
            if (conditions == null || conditions.Count == 0) throw new ArgumentException("At least one condition is required.", nameof(conditions));
            range = ValidateAutoFilterColumnOffset(range, columnId);

            WriteLock(() => {
                Worksheet worksheet = WorksheetRoot;
                AutoFilter autoFilter = GetOrReplaceAutoFilter(worksheet, range, out Table? tableOwner);

                FilterColumn? existingColumn = autoFilter.Elements<FilterColumn>().FirstOrDefault(fc => fc.ColumnId?.Value == columnId);
                existingColumn?.Remove();

                var filterColumn = new FilterColumn { ColumnId = columnId };
                var customFilters = new CustomFilters {
                    And = matchAll,
                };
                foreach ((FilterOperatorValues filterOperator, string value) in conditions) {
                    customFilters.Append(new CustomFilter {
                        Operator = filterOperator,
                        Val = value,
                    });
                }

                if (customFilters.Elements<CustomFilter>().Any()) {
                    filterColumn.Append(customFilters);
                    autoFilter.Append(filterColumn);
                    DeferAutoFilterOwner(tableOwner);
                }
            });
        }

        internal void ApplyAutoFilterBlankCriteria(string range, uint columnId) {
            MaterializeDeferredAutoFilterWorkbook();
            if (string.IsNullOrWhiteSpace(range)) throw new ArgumentNullException(nameof(range));
            range = ValidateAutoFilterColumnOffset(range, columnId);

            WriteLock(() => {
                Worksheet worksheet = WorksheetRoot;
                AutoFilter autoFilter = GetOrReplaceAutoFilter(worksheet, range, out Table? tableOwner);

                FilterColumn? existingColumn = autoFilter.Elements<FilterColumn>().FirstOrDefault(fc => fc.ColumnId?.Value == columnId);
                existingColumn?.Remove();

                var filterColumn = new FilterColumn { ColumnId = columnId };
                filterColumn.Append(new Filters { Blank = true });
                autoFilter.Append(filterColumn);
                DeferAutoFilterOwner(tableOwner);
            });
        }

        internal void ApplyAutoFilterTop10Criteria(
            string range,
            uint columnId,
            ushort value,
            bool isTop,
            bool isPercent) {
            MaterializeDeferredAutoFilterWorkbook();
            if (string.IsNullOrWhiteSpace(range)) throw new ArgumentNullException(nameof(range));
            ushort maximumValue = isPercent ? (ushort)100 : (ushort)500;
            if (value < 1 || value > maximumValue) {
                throw new ArgumentOutOfRangeException(
                    nameof(value),
                    $"Top10 AutoFilter values must be between 1 and {maximumValue}.");
            }
            range = ValidateAutoFilterColumnOffset(range, columnId);

            WriteLock(() => {
                Worksheet worksheet = WorksheetRoot;
                AutoFilter autoFilter = GetOrReplaceAutoFilter(worksheet, range, out Table? tableOwner);

                FilterColumn? existingColumn = autoFilter.Elements<FilterColumn>().FirstOrDefault(fc => fc.ColumnId?.Value == columnId);
                existingColumn?.Remove();

                var filterColumn = new FilterColumn { ColumnId = columnId };
                filterColumn.Append(new Top10 {
                    Top = isTop,
                    Percent = isPercent,
                    Val = (double)value,
                });
                autoFilter.Append(filterColumn);
                DeferAutoFilterOwner(tableOwner);
            });
        }

        private AutoFilter GetOrReplaceAutoFilter(Worksheet worksheet, string range, out Table? tableOwner) {
            Table[] matchingTables = _worksheetPart.TableDefinitionParts
                .Select(part => part.Table)
                .Where(table => table != null
                    && TryGetTableAutoFilterRange(table, out string tableFilterRange)
                    && (AutoFilterRangesMatch(table.Reference?.Value, range)
                        || AutoFilterRangesMatch(tableFilterRange, range)))
                .Cast<Table>()
                .ToArray();
            if (matchingTables.Length > 1) {
                throw new InvalidDataException("Multiple table AutoFilters own the requested range.");
            }
            if (matchingTables.Length == 1) {
                tableOwner = matchingTables[0];
                TryGetTableAutoFilterRange(tableOwner, out string tableFilterRange);
                AutoFilter? worksheetFilter = worksheet.GetFirstChild<AutoFilter>();
                if (worksheetFilter != null
                    && AutoFilterRangesOverlap(worksheetFilter.Reference?.Value, tableOwner.Reference?.Value)) {
                    worksheetFilter.Remove();
                }
                AutoFilter? tableFilter = tableOwner.GetFirstChild<AutoFilter>();
                if (tableFilter == null) {
                    tableFilter = new AutoFilter { Reference = tableFilterRange };
                    TableColumns? tableColumns = tableOwner.GetFirstChild<TableColumns>();
                    if (tableColumns == null) tableOwner.Append(tableFilter);
                    else tableOwner.InsertBefore(tableFilter, tableColumns);
                } else {
                    tableFilter.Reference = tableFilterRange;
                }
                return tableFilter;
            }
            tableOwner = null;
            return GetOrReplaceWorksheetAutoFilter(worksheet, range);
        }

        private void DeferAutoFilterOwner(Table? tableOwner) {
            if (tableOwner != null) {
                DeferPartSave(tableOwner.TableDefinitionPart
                    ?? throw new InvalidOperationException("The AutoFilter table is not attached to a package part."));
            }
        }

        private static AutoFilter GetOrReplaceWorksheetAutoFilter(Worksheet worksheet, string range) {
            AutoFilter? autoFilter = worksheet.GetFirstChild<AutoFilter>();
            if (autoFilter != null && AutoFilterRangesMatch(autoFilter.Reference?.Value, range)) {
                return autoFilter;
            }
            autoFilter?.Remove();
            autoFilter = new AutoFilter { Reference = range };
            SheetData? sheetData = worksheet.GetFirstChild<SheetData>();
            if (sheetData == null) {
                worksheet.Append(autoFilter);
                return autoFilter;
            }
            ConditionalFormatting? conditionalFormatting = worksheet.GetFirstChild<ConditionalFormatting>();
            if (conditionalFormatting != null) worksheet.InsertBefore(autoFilter, conditionalFormatting);
            else worksheet.InsertAfter(autoFilter, sheetData);
            return autoFilter;
        }

        private static bool AutoFilterRangesMatch(string? existingRange, string requestedRange) {
            if (!ExcelReference.TryParse(existingRange, out ExcelReference? existing)
                || !ExcelReference.TryParse(requestedRange, out ExcelReference? requested)
                || existing!.IsQualified
                || requested!.IsQualified) {
                return false;
            }
            existing.GetBounds(out int existingR1, out int existingC1, out int existingR2, out int existingC2);
            requested.GetBounds(out int requestedR1, out int requestedC1, out int requestedR2, out int requestedC2);
            return existingR1 == requestedR1
                && existingC1 == requestedC1
                && existingR2 == requestedR2
                && existingC2 == requestedC2;
        }

        private static bool AutoFilterRangesOverlap(string? firstRange, string? secondRange) {
            if (!ExcelReference.TryParse(firstRange, out ExcelReference? first)
                || !ExcelReference.TryParse(secondRange, out ExcelReference? second)
                || first!.IsQualified
                || second!.IsQualified) {
                return false;
            }
            first.GetBounds(out int firstR1, out int firstC1, out int firstR2, out int firstC2);
            second.GetBounds(out int secondR1, out int secondC1, out int secondR2, out int secondC2);
            return firstR1 <= secondR2
                && firstR2 >= secondR1
                && firstC1 <= secondC2
                && firstC2 >= secondC1;
        }

        private static bool TryGetTableAutoFilterRange(Table table, out string range) {
            range = string.Empty;
            if (table.HeaderRowCount?.Value == 0U
                || !ExcelA1.TryParseRange(table.Reference?.Value ?? string.Empty, out int r1, out int c1, out int r2, out int c2)) {
                return false;
            }
            int totalsRows = table.TotalsRowShown?.Value == true
                ? Math.Max(1, checked((int)(table.TotalsRowCount?.Value ?? 1U)))
                : 0;
            int filterLastRow = r2 - totalsRows;
            if (filterLastRow < r1) return false;
            range = ExcelA1.CellReference(r1, c1) + ":" + ExcelA1.CellReference(filterLastRow, c2);
            return true;
        }

        private static string ValidateAutoFilterColumnOffset(string range, uint columnId) {
            ExcelReference reference = ExcelReference.Parse(range);
            if (reference.Kind != ExcelReferenceKind.Cell && reference.Kind != ExcelReferenceKind.Range) {
                throw new ArgumentException("AutoFilter ranges require a cell or rectangular range reference.", nameof(range));
            }
            if (reference.IsQualified) {
                throw new ArgumentException("Worksheet AutoFilter ranges must be local, unqualified references.", nameof(range));
            }
            reference.GetBounds(out int firstRow, out int firstColumn, out int lastRow, out int lastColumn);
            uint width = checked((uint)(lastColumn - firstColumn + 1));
            if (columnId >= width) throw new ArgumentOutOfRangeException(nameof(columnId), "AutoFilter column offset must be inside the filter range.");
            string firstCell = ExcelA1.CellReference(firstRow, firstColumn);
            return reference.Kind == ExcelReferenceKind.Cell
                ? firstCell
                : firstCell + ":" + ExcelA1.CellReference(lastRow, lastColumn);
        }

    }
}
