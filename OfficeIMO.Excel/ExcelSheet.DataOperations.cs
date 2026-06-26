using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;
using System.Text;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Data operations: autofilter, sorting (values-only), find/replace, and validation.
    /// Kept small and focused; advanced formatting lives in dedicated files.
    /// </summary>
    public partial class ExcelSheet {
        // -------- AutoFilter --------

        /// <summary>
        /// Attaches an AutoFilter to the given A1 range (e.g., "A1:C200").
        /// </summary>
        public void AutoFilterAdd(string a1Range) {
            if (string.IsNullOrWhiteSpace(a1Range)) throw new ArgumentNullException(nameof(a1Range));
            WriteLock(() => {
                var ws = WorksheetRoot;
                // Remove existing AutoFilter, if any
                var existing = ws.Elements<AutoFilter>().FirstOrDefault();
                existing?.Remove();

                ws.InsertAfter(new AutoFilter { Reference = a1Range }, ws.GetFirstChild<SheetData>());
                ws.Save();
            });
        }

        /// <summary>
        /// Clears any AutoFilter from the worksheet.
        /// </summary>
        public void AutoFilterClear() {
            WriteLock(() => {
                var ws = WorksheetRoot;
                var existing = ws.Elements<AutoFilter>().FirstOrDefault();
                existing?.Remove();
                ws.Save();
            });
        }

        /// <summary>
        /// Applies an AutoFilter equals filter to a column resolved by header within the current AutoFilter range.
        /// Ensures an AutoFilter exists over the sheet's UsedRange when none is present.
        /// When the header is missing the operation is skipped.
        /// </summary>
        public void AutoFilterByHeaderEquals(string header, IEnumerable<string> values) {
            if (string.IsNullOrWhiteSpace(header)) throw new ArgumentNullException(nameof(header));
            if (values == null) throw new ArgumentNullException(nameof(values));

            WriteLock(() => {
                if (!TryGetColumnIndexByHeader(header, out var colIndex))
                    return;
                var ws = WorksheetRoot;
                var af = ws.Elements<AutoFilter>().FirstOrDefault();
                if (af == null) {
                    var used = GetUsedRangeA1();
                    af = new AutoFilter { Reference = used };
                    ws.InsertAfter(af, ws.GetFirstChild<SheetData>());
                }

                var (r1, c1, r2, c2) = A1.ParseRange(af.Reference!);
                if (colIndex < c1 || colIndex > c2)
                    throw new ArgumentOutOfRangeException(nameof(header), $"Header '{header}' is outside the AutoFilter range {af.Reference}.");

                // ColumnId is zero-based within the AutoFilter range
                uint columnId = (uint)(colIndex - c1);

                // Remove existing filter for this ColumnId
                var existingColumn = af.Elements<FilterColumn>().FirstOrDefault(fc => fc.ColumnId?.Value == columnId);
                existingColumn?.Remove();

                var fcNew = new FilterColumn { ColumnId = columnId };
                var filters = new Filters();
                foreach (var v in values.Distinct(StringComparer.OrdinalIgnoreCase)) {
                    if (v == null) continue;
                    filters.Append(new Filter { Val = v });
                }
                fcNew.Append(filters);
                af.Append(fcNew);
                ws.Save();
            });
        }

        /// <summary>
        /// Applies equals filters for multiple headers at once (AND semantics across columns, OR semantics within a column).
        /// Headers that cannot be resolved are ignored.
        /// </summary>
        public void AutoFilterByHeadersEquals(params (string Header, IEnumerable<string> Values)[] filters) {
            if (filters == null || filters.Length == 0) throw new ArgumentException("At least one filter must be provided.", nameof(filters));
            WriteLock(() => {
                var toApply = new List<(int ColumnIndex, IEnumerable<string> Values)>();
                var headerMap = GetHeaderMapCached(DefaultHeaderReadOptions);
                foreach (var (header, values) in filters) {
                    if (string.IsNullOrWhiteSpace(header)) continue;
                    if (!headerMap.TryGetValue(header, out var colIndex)) continue;
                    toApply.Add((colIndex, values ?? Array.Empty<string>()));
                }
                if (toApply.Count == 0) return;

                var ws = WorksheetRoot;
                var af = ws.Elements<AutoFilter>().FirstOrDefault();
                if (af == null) {
                    var used = GetUsedRangeA1();
                    af = new AutoFilter { Reference = used };
                    ws.InsertAfter(af, ws.GetFirstChild<SheetData>());
                }

                var (r1, c1, r2, c2) = A1.ParseRange(af.Reference!);
                var existingColumns = new Dictionary<uint, FilterColumn>();
                foreach (var existing in af.Elements<FilterColumn>()) {
                    if (existing.ColumnId?.Value is uint existingColumnId) {
                        existingColumns[existingColumnId] = existing;
                    }
                }

                bool changed = false;
                foreach (var (colIndex, values) in toApply) {
                    if (colIndex < c1 || colIndex > c2) continue;
                    uint columnId = (uint)(colIndex - c1);
                    var filterValues = BuildDistinctFilterValues(values);

                    if (existingColumns.TryGetValue(columnId, out var existingColumn)) {
                        if (FilterColumnMatchesValues(existingColumn, filterValues)) {
                            continue;
                        }

                        existingColumn.Remove();
                    }

                    var fcNew = new FilterColumn { ColumnId = columnId };
                    var filtersNode = new Filters();
                    foreach (var v in filterValues) {
                        filtersNode.Append(new Filter { Val = v });
                    }
                    fcNew.Append(filtersNode);
                    af.Append(fcNew);
                    existingColumns[columnId] = fcNew;
                    changed = true;
                }
                if (changed) {
                    ws.Save();
                }
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

        /// <summary>
        /// Applies an AutoFilter text contains filter to a column resolved by header within the current AutoFilter range.
        /// Uses wildcard pattern matching ("*text*") via CustomFilters with Equal operator.
        /// When the header is missing the operation is skipped.
        /// </summary>
        public void AutoFilterByHeaderContains(string header, string containsText) {
            if (string.IsNullOrWhiteSpace(header)) throw new ArgumentNullException(nameof(header));
            if (containsText is null) throw new ArgumentNullException(nameof(containsText));

            ApplyCustomAutoFilterByHeader(header, FilterOperatorValues.Equal, "*" + containsText + "*");
        }

        /// <summary>
        /// Applies an AutoFilter text does-not-contain filter to a column resolved by header within the current AutoFilter range.
        /// Uses wildcard pattern matching ("*text*") via CustomFilters with NotEqual operator.
        /// </summary>
        public void AutoFilterByHeaderDoesNotContain(string header, string containsText) {
            if (string.IsNullOrWhiteSpace(header)) throw new ArgumentNullException(nameof(header));
            if (containsText is null) throw new ArgumentNullException(nameof(containsText));

            ApplyCustomAutoFilterByHeader(header, FilterOperatorValues.NotEqual, "*" + containsText + "*");
        }

        /// <summary>
        /// Applies an AutoFilter text starts-with filter to a column resolved by header within the current AutoFilter range.
        /// Uses wildcard pattern matching ("text*") via CustomFilters with Equal operator.
        /// </summary>
        public void AutoFilterByHeaderStartsWith(string header, string startsWithText) {
            if (string.IsNullOrWhiteSpace(header)) throw new ArgumentNullException(nameof(header));
            if (startsWithText is null) throw new ArgumentNullException(nameof(startsWithText));

            ApplyCustomAutoFilterByHeader(header, FilterOperatorValues.Equal, startsWithText + "*");
        }

        /// <summary>
        /// Applies an AutoFilter text ends-with filter to a column resolved by header within the current AutoFilter range.
        /// Uses wildcard pattern matching ("*text") via CustomFilters with Equal operator.
        /// </summary>
        public void AutoFilterByHeaderEndsWith(string header, string endsWithText) {
            if (string.IsNullOrWhiteSpace(header)) throw new ArgumentNullException(nameof(header));
            if (endsWithText is null) throw new ArgumentNullException(nameof(endsWithText));

            ApplyCustomAutoFilterByHeader(header, FilterOperatorValues.Equal, "*" + endsWithText);
        }

        /// <summary>
        /// Applies an AutoFilter numeric greater-than-or-equal filter to a column resolved by header within the current AutoFilter range.
        /// </summary>
        public void AutoFilterByHeaderGreaterThanOrEqual(string header, double value) {
            if (string.IsNullOrWhiteSpace(header)) throw new ArgumentNullException(nameof(header));

            ApplyCustomAutoFilterByHeader(header, FilterOperatorValues.GreaterThanOrEqual, value.ToString(CultureInfo.InvariantCulture));
        }

        /// <summary>
        /// Applies an AutoFilter numeric not-equal filter to a column resolved by header within the current AutoFilter range.
        /// </summary>
        public void AutoFilterByHeaderNotEqual(string header, double value) {
            if (string.IsNullOrWhiteSpace(header)) throw new ArgumentNullException(nameof(header));

            ApplyCustomAutoFilterByHeader(header, FilterOperatorValues.NotEqual, value.ToString(CultureInfo.InvariantCulture));
        }

        /// <summary>
        /// Applies an AutoFilter numeric less-than-or-equal filter to a column resolved by header within the current AutoFilter range.
        /// </summary>
        public void AutoFilterByHeaderLessThanOrEqual(string header, double value) {
            if (string.IsNullOrWhiteSpace(header)) throw new ArgumentNullException(nameof(header));

            ApplyCustomAutoFilterByHeader(header, FilterOperatorValues.LessThanOrEqual, value.ToString(CultureInfo.InvariantCulture));
        }

        /// <summary>
        /// Applies an AutoFilter numeric inclusive range filter to a column resolved by header within the current AutoFilter range.
        /// </summary>
        public void AutoFilterByHeaderBetween(string header, double minimumInclusive, double maximumInclusive) {
            if (string.IsNullOrWhiteSpace(header)) throw new ArgumentNullException(nameof(header));
            if (maximumInclusive < minimumInclusive) throw new ArgumentOutOfRangeException(nameof(maximumInclusive), "The maximum value must be greater than or equal to the minimum value.");

            ApplyDualCustomAutoFilterByHeader(
                header,
                matchAll: true,
                firstOperator: FilterOperatorValues.GreaterThanOrEqual,
                firstValue: minimumInclusive.ToString(CultureInfo.InvariantCulture),
                secondOperator: FilterOperatorValues.LessThanOrEqual,
                secondValue: maximumInclusive.ToString(CultureInfo.InvariantCulture));
        }

        /// <summary>
        /// Applies an AutoFilter numeric outside-range filter to a column resolved by header within the current AutoFilter range.
        /// Values lower than <paramref name="minimumExclusive"/> or higher than <paramref name="maximumExclusive"/> remain visible.
        /// </summary>
        public void AutoFilterByHeaderNotBetween(string header, double minimumExclusive, double maximumExclusive) {
            if (string.IsNullOrWhiteSpace(header)) throw new ArgumentNullException(nameof(header));
            if (maximumExclusive < minimumExclusive) throw new ArgumentOutOfRangeException(nameof(maximumExclusive), "The maximum value must be greater than or equal to the minimum value.");

            ApplyDualCustomAutoFilterByHeader(
                header,
                matchAll: false,
                firstOperator: FilterOperatorValues.LessThan,
                firstValue: minimumExclusive.ToString(CultureInfo.InvariantCulture),
                secondOperator: FilterOperatorValues.GreaterThan,
                secondValue: maximumExclusive.ToString(CultureInfo.InvariantCulture));
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
                    ws.Save();
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

        private void ApplyCustomAutoFilterByHeader(string header, FilterOperatorValues filterOperator, string value) {
            WriteLock(() => {
                if (!TryGetColumnIndexByHeader(header, out var colIndex))
                    return;
                var ws = WorksheetRoot;
                var af = ws.Elements<AutoFilter>().FirstOrDefault();
                if (af == null) {
                    var used = GetUsedRangeA1();
                    af = new AutoFilter { Reference = used };
                    ws.InsertAfter(af, ws.GetFirstChild<SheetData>());
                }

                var (r1, c1, r2, c2) = A1.ParseRange(af.Reference!);
                if (colIndex < c1 || colIndex > c2)
                    throw new ArgumentOutOfRangeException(nameof(header), $"Header '{header}' is outside the AutoFilter range {af.Reference}.");

                uint columnId = (uint)(colIndex - c1);
                var existingColumn = af.Elements<FilterColumn>().FirstOrDefault(fc => fc.ColumnId?.Value == columnId);
                existingColumn?.Remove();

                var filterColumn = new FilterColumn { ColumnId = columnId };
                var customFilters = new CustomFilters();
                customFilters.Append(new CustomFilter {
                    Operator = filterOperator,
                    Val = value,
                });
                filterColumn.Append(customFilters);
                af.Append(filterColumn);
                ws.Save();
            });
        }

        private void ApplyDualCustomAutoFilterByHeader(
            string header,
            bool matchAll,
            FilterOperatorValues firstOperator,
            string firstValue,
            FilterOperatorValues secondOperator,
            string secondValue) {
            WriteLock(() => {
                if (!TryGetColumnIndexByHeader(header, out var colIndex))
                    return;
                var ws = WorksheetRoot;
                var af = ws.Elements<AutoFilter>().FirstOrDefault();
                if (af == null) {
                    var used = GetUsedRangeA1();
                    af = new AutoFilter { Reference = used };
                    ws.InsertAfter(af, ws.GetFirstChild<SheetData>());
                }

                var (r1, c1, r2, c2) = A1.ParseRange(af.Reference!);
                if (colIndex < c1 || colIndex > c2)
                    throw new ArgumentOutOfRangeException(nameof(header), $"Header '{header}' is outside the AutoFilter range {af.Reference}.");

                uint columnId = (uint)(colIndex - c1);
                var existingColumn = af.Elements<FilterColumn>().FirstOrDefault(fc => fc.ColumnId?.Value == columnId);
                existingColumn?.Remove();

                var filterColumn = new FilterColumn { ColumnId = columnId };
                var customFilters = new CustomFilters {
                    And = matchAll,
                };
                customFilters.Append(new CustomFilter {
                    Operator = firstOperator,
                    Val = firstValue,
                });
                customFilters.Append(new CustomFilter {
                    Operator = secondOperator,
                    Val = secondValue,
                });

                filterColumn.Append(customFilters);
                af.Append(filterColumn);
                ws.Save();
            });
        }

        internal void ApplyAutoFilterCustomCriteria(
            string range,
            uint columnId,
            bool matchAll,
            IReadOnlyList<(FilterOperatorValues Operator, string Value)> conditions) {
            if (string.IsNullOrWhiteSpace(range)) throw new ArgumentNullException(nameof(range));
            if (conditions == null || conditions.Count == 0) throw new ArgumentException("At least one condition is required.", nameof(conditions));

            WriteLock(() => {
                Worksheet worksheet = WorksheetRoot;
                AutoFilter? autoFilter = worksheet.GetFirstChild<AutoFilter>();
                if (autoFilter == null || !string.Equals(autoFilter.Reference?.Value, range, StringComparison.OrdinalIgnoreCase)) {
                    autoFilter?.Remove();
                    autoFilter = new AutoFilter { Reference = range };
                    var sheetData = worksheet.GetFirstChild<SheetData>();
                    if (sheetData != null) {
                        var conditionalFormatting = worksheet.GetFirstChild<ConditionalFormatting>();
                        if (conditionalFormatting != null) {
                            worksheet.InsertBefore(autoFilter, conditionalFormatting);
                        } else {
                            worksheet.InsertAfter(autoFilter, sheetData);
                        }
                    } else {
                        worksheet.Append(autoFilter);
                    }
                }

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
                    worksheet.Save();
                }
            });
        }

        internal void ApplyAutoFilterBlankCriteria(string range, uint columnId) {
            if (string.IsNullOrWhiteSpace(range)) throw new ArgumentNullException(nameof(range));

            WriteLock(() => {
                Worksheet worksheet = WorksheetRoot;
                AutoFilter? autoFilter = worksheet.GetFirstChild<AutoFilter>();
                if (autoFilter == null || !string.Equals(autoFilter.Reference?.Value, range, StringComparison.OrdinalIgnoreCase)) {
                    autoFilter?.Remove();
                    autoFilter = new AutoFilter { Reference = range };
                    var sheetData = worksheet.GetFirstChild<SheetData>();
                    if (sheetData != null) {
                        var conditionalFormatting = worksheet.GetFirstChild<ConditionalFormatting>();
                        if (conditionalFormatting != null) {
                            worksheet.InsertBefore(autoFilter, conditionalFormatting);
                        } else {
                            worksheet.InsertAfter(autoFilter, sheetData);
                        }
                    } else {
                        worksheet.Append(autoFilter);
                    }
                }

                FilterColumn? existingColumn = autoFilter.Elements<FilterColumn>().FirstOrDefault(fc => fc.ColumnId?.Value == columnId);
                existingColumn?.Remove();

                var filterColumn = new FilterColumn { ColumnId = columnId };
                filterColumn.Append(new Filters { Blank = true });
                autoFilter.Append(filterColumn);
                worksheet.Save();
            });
        }

        internal void ApplyAutoFilterTop10Criteria(
            string range,
            uint columnId,
            ushort value,
            bool isTop,
            bool isPercent) {
            if (string.IsNullOrWhiteSpace(range)) throw new ArgumentNullException(nameof(range));
            if (value < 1 || value > 500) throw new ArgumentOutOfRangeException(nameof(value), "Top10 AutoFilter values must be between 1 and 500.");

            WriteLock(() => {
                Worksheet worksheet = WorksheetRoot;
                AutoFilter? autoFilter = worksheet.GetFirstChild<AutoFilter>();
                if (autoFilter == null || !string.Equals(autoFilter.Reference?.Value, range, StringComparison.OrdinalIgnoreCase)) {
                    autoFilter?.Remove();
                    autoFilter = new AutoFilter { Reference = range };
                    var sheetData = worksheet.GetFirstChild<SheetData>();
                    if (sheetData != null) {
                        var conditionalFormatting = worksheet.GetFirstChild<ConditionalFormatting>();
                        if (conditionalFormatting != null) {
                            worksheet.InsertBefore(autoFilter, conditionalFormatting);
                        } else {
                            worksheet.InsertAfter(autoFilter, sheetData);
                        }
                    } else {
                        worksheet.Append(autoFilter);
                    }
                }

                FilterColumn? existingColumn = autoFilter.Elements<FilterColumn>().FirstOrDefault(fc => fc.ColumnId?.Value == columnId);
                existingColumn?.Remove();

                var filterColumn = new FilterColumn { ColumnId = columnId };
                filterColumn.Append(new Top10 {
                    Top = isTop,
                    Percent = isPercent,
                    Val = (double)value,
                });
                autoFilter.Append(filterColumn);
                worksheet.Save();
            });
        }

    }
}
