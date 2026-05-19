using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Enables a totals row for the table covering <paramref name="range"/> and assigns per-column functions by header name.
        /// Supported functions are those in TotalsRowFunctionValues (Sum, Average, Count, Min, Max, etc.).
        /// </summary>
        /// <param name="range">Address of the table range (for example, "A1:D10") whose totals row should be displayed.</param>
        /// <param name="byHeader">Mapping of table header names to the totals function that should be applied for each column.</param>
        public void SetTableTotals(string range, System.Collections.Generic.Dictionary<string, DocumentFormat.OpenXml.Spreadsheet.TotalsRowFunctionValues> byHeader) {
            if (string.IsNullOrWhiteSpace(range)) throw new System.ArgumentNullException(nameof(range));
            if (byHeader == null) throw new System.ArgumentNullException(nameof(byHeader));

            SetTableTotalsCore(range, byHeader, throwIfMissing: false);
        }

        /// <summary>
        /// Enables a totals row for the named table and assigns per-column functions by header name.
        /// </summary>
        /// <param name="tableName">Table name or display name.</param>
        /// <param name="byHeader">Mapping of table header names to totals functions.</param>
        public void SetTableTotalsByName(string tableName, System.Collections.Generic.IDictionary<string, DocumentFormat.OpenXml.Spreadsheet.TotalsRowFunctionValues> byHeader) {
            if (string.IsNullOrWhiteSpace(tableName)) throw new System.ArgumentNullException(nameof(tableName));
            if (byHeader == null) throw new System.ArgumentNullException(nameof(byHeader));

            SetTableTotalsCore(tableName, byHeader, throwIfMissing: true);
        }

        /// <summary>
        /// Clears totals-row settings for the table identified by range, name, or display name.
        /// </summary>
        /// <param name="tableOrRange">Table range, name, or display name.</param>
        public void ClearTableTotals(string tableOrRange) {
            if (string.IsNullOrWhiteSpace(tableOrRange)) throw new System.ArgumentNullException(nameof(tableOrRange));

            WriteLock(() => {
                var table = FindTableByRangeNameOrDisplayName(tableOrRange);
                if (table == null) {
                    throw new InvalidOperationException($"Table '{tableOrRange}' was not found on worksheet '{Name}'.");
                }

                table.TotalsRowShown = false;
                table.TotalsRowCount = 0U;
                foreach (var tableColumn in table.TableColumns?.Elements<TableColumn>() ?? Enumerable.Empty<TableColumn>()) {
                    tableColumn.TotalsRowFunction = null;
                    tableColumn.TotalsRowFormula = null;
                    tableColumn.TotalsRowLabel = null;
                }

                table.Save();
                WorksheetRoot.Save();
            });
        }

        /// <summary>
        /// Updates the visual style flags for the table identified by range, name, or display name.
        /// </summary>
        /// <param name="tableOrRange">Table range, name, or display name.</param>
        /// <param name="style">Table style to apply.</param>
        /// <param name="showFirstColumn">Optional first-column emphasis flag.</param>
        /// <param name="showLastColumn">Optional last-column emphasis flag.</param>
        /// <param name="showRowStripes">Optional row stripe flag.</param>
        /// <param name="showColumnStripes">Optional column stripe flag.</param>
        public void SetTableStyle(
            string tableOrRange,
            TableStyle style,
            bool? showFirstColumn = null,
            bool? showLastColumn = null,
            bool? showRowStripes = null,
            bool? showColumnStripes = null) {
            if (string.IsNullOrWhiteSpace(tableOrRange)) throw new System.ArgumentNullException(nameof(tableOrRange));

            WriteLock(() => {
                var table = FindTableByRangeNameOrDisplayName(tableOrRange);
                if (table == null) {
                    throw new InvalidOperationException($"Table '{tableOrRange}' was not found on worksheet '{Name}'.");
                }

                var styleInfo = table.TableStyleInfo;
                if (styleInfo == null) {
                    styleInfo = new TableStyleInfo();
                    var extensionList = table.GetFirstChild<TableExtensionList>();
                    if (extensionList == null) {
                        table.Append(styleInfo);
                    } else {
                        table.InsertBefore(styleInfo, extensionList);
                    }
                }

                styleInfo.Name = style.ToString();
                if (showFirstColumn.HasValue) styleInfo.ShowFirstColumn = showFirstColumn.Value;
                if (showLastColumn.HasValue) styleInfo.ShowLastColumn = showLastColumn.Value;
                if (showRowStripes.HasValue) styleInfo.ShowRowStripes = showRowStripes.Value;
                if (showColumnStripes.HasValue) styleInfo.ShowColumnStripes = showColumnStripes.Value;

                table.Save();
                WorksheetRoot.Save();
            });
        }

        private void SetTableTotalsCore(string tableOrRange, System.Collections.Generic.IDictionary<string, DocumentFormat.OpenXml.Spreadsheet.TotalsRowFunctionValues> byHeader, bool throwIfMissing) {
            var totalsByHeader = new System.Collections.Generic.Dictionary<string, DocumentFormat.OpenXml.Spreadsheet.TotalsRowFunctionValues>(byHeader, System.StringComparer.OrdinalIgnoreCase);
            WriteLock(() => {
                var table = FindTableByRangeNameOrDisplayName(tableOrRange);
                if (table == null) {
                    if (throwIfMissing) {
                        throw new InvalidOperationException($"Table '{tableOrRange}' was not found on worksheet '{Name}'.");
                    }

                    return;
                }

                table.TotalsRowShown = true;
                table.TotalsRowCount = 1U;
                var tableColumns = table.TableColumns ?? throw new InvalidOperationException("Table columns are missing.");
                foreach (var tc in tableColumns.Elements<TableColumn>()) {
                    var name = tc.Name?.Value ?? string.Empty;
                    if (totalsByHeader.TryGetValue(name, out var fn)) {
                        tc.TotalsRowFunction = fn;
                    }
                }

                table.Save();
                WorksheetRoot.Save();
            });
        }

        private Table? FindTableByRangeNameOrDisplayName(string tableOrRange) {
            return _worksheetPart.TableDefinitionParts
                .Select(part => part.Table)
                .FirstOrDefault(table => table != null && (
                    string.Equals(table.Reference?.Value, tableOrRange, StringComparison.OrdinalIgnoreCase)
                    || string.Equals(table.Name?.Value, tableOrRange, StringComparison.OrdinalIgnoreCase)
                    || string.Equals(table.DisplayName?.Value, tableOrRange, StringComparison.OrdinalIgnoreCase)));
        }
        /// <summary>
        /// Adds an AutoFilter to the worksheet or table.
        /// </summary>
        /// <param name="range">The cell range to apply the filter to.</param>
        /// <param name="filterCriteria">Optional filter criteria to apply.</param>
        /// <remarks>
        /// Smart conflict resolution:
        /// - If a table exists on the same range, the AutoFilter is added to the table instead of the worksheet
        /// - If a worksheet-level AutoFilter exists, it's removed before adding to the table
        /// - This ensures Excel won't complain about conflicting filters
        /// Order doesn't matter - whether you add Table then AutoFilter or AutoFilter then Table, 
        /// the final result will be a table with AutoFilter enabled.
        /// </remarks>
        public void AddAutoFilter(string range, Dictionary<uint, IEnumerable<string>>? filterCriteria = null) {
            if (string.IsNullOrEmpty(range)) {
                throw new ArgumentNullException(nameof(range));
            }

            WriteLock(() => {
                // SMART DETECTION: Check if there's a table on this range
                // If there is, we'll add the AutoFilter to the table instead of the worksheet
                foreach (var tableDefinitionPart in _worksheetPart.TableDefinitionParts) {
                    var table = tableDefinitionPart.Table;
                    if (table is not null && table.Reference?.Value == range) {
                        // Found a table on the same range - add/update its AutoFilter

                        // First, remove any worksheet-level AutoFilter to avoid conflicts
                        var worksheetAutoFilter = WorksheetRoot.Elements<AutoFilter>().FirstOrDefault();
                        if (worksheetAutoFilter != null && worksheetAutoFilter.Reference?.Value == range) {
                            WorksheetRoot.RemoveChild(worksheetAutoFilter);
                        }

                        // Now handle the table's AutoFilter
                        var tableAutoFilter = table.Elements<AutoFilter>().FirstOrDefault();
                        if (tableAutoFilter != null) {
                            // Remove existing to replace with new one
                            table.RemoveChild(tableAutoFilter);
                        }

                        // Create new AutoFilter
                        var newAutoFilter = new AutoFilter { Reference = range };

                        // Apply filter criteria if provided
                        if (filterCriteria != null) {
                            foreach (KeyValuePair<uint, IEnumerable<string>> criteria in filterCriteria) {
                                FilterColumn filterColumn = new FilterColumn { ColumnId = criteria.Key };
                                Filters filters = new Filters();
                                foreach (string value in criteria.Value) {
                                    filters.Append(new Filter { Val = value });
                                }
                                filterColumn.Append(filters);
                                newAutoFilter.Append(filterColumn);
                            }
                        }

                        // Add AutoFilter to the table (must be before tableColumns)
                        var tableColumns = table.Elements<TableColumns>().FirstOrDefault();
                        if (tableColumns != null) {
                            table.InsertBefore(newAutoFilter, tableColumns);
                        } else {
                            table.InsertAt(newAutoFilter, 0);
                        }

                        table.Save();
                        return; // Exit early - we've handled the AutoFilter via the table
                    }
                }

                // No table found, add worksheet-level AutoFilter
                Worksheet worksheet = WorksheetRoot;

                // Remove any existing worksheet AutoFilter
                AutoFilter? existing = worksheet.Elements<AutoFilter>().FirstOrDefault();
                if (existing != null) {
                    worksheet.RemoveChild(existing);
                }

                // Create new AutoFilter
                AutoFilter autoFilter = new AutoFilter { Reference = range };

                if (filterCriteria != null) {
                    foreach (KeyValuePair<uint, IEnumerable<string>> criteria in filterCriteria) {
                        FilterColumn filterColumn = new FilterColumn { ColumnId = criteria.Key };
                        Filters filters = new Filters();
                        foreach (string value in criteria.Value) {
                            filters.Append(new Filter { Val = value });
                        }

                        filterColumn.Append(filters);
                        autoFilter.Append(filterColumn);
                    }
                }

                // Insert AutoFilter after SheetData but before ConditionalFormatting
                var sheetData = worksheet.GetFirstChild<SheetData>();
                if (sheetData != null) {
                    // Find the right position to insert AutoFilter
                    var conditionalFormatting = worksheet.Elements<ConditionalFormatting>().FirstOrDefault();
                    if (conditionalFormatting != null) {
                        worksheet.InsertBefore(autoFilter, conditionalFormatting);
                    } else {
                        worksheet.InsertAfter(autoFilter, sheetData);
                    }
                } else {
                    worksheet.Append(autoFilter);
                }

                worksheet.Save();
            });
        }

        /// <summary>
        /// Adds an Excel table to the worksheet over the specified range.
        /// </summary>
        /// <param name="range">Cell range (e.g. "A1:B3") defining the table area.</param>
        /// <param name="hasHeader">Indicates whether the first row is a header row.</param>
        /// <param name="name">Name of the table. If empty, a default name is used.</param>
        /// <param name="style">Table style to apply.</param>
        /// <remarks>
        /// All cells within <paramref name="range"/> must exist. Missing cells are automatically created with empty values.
        /// </remarks>
        /// <exception cref="ArgumentNullException">Thrown when <paramref name="range"/> is null or empty.</exception>
        /// <exception cref="ArgumentException">Thrown when <paramref name="range"/> is not in a valid format.</exception>
        /// <exception cref="InvalidOperationException">Thrown when the specified range overlaps with an existing table.</exception>
        public void AddTable(string range, bool hasHeader, string name, TableStyle style) {
            AddTableCore(range, hasHeader, name, style, includeAutoFilter: true, ensureRangeCellsExist: true);
        }

        /// <summary>
        /// Gets the A1 range covered by a table on this worksheet.
        /// </summary>
        /// <param name="tableName">Table name or display name.</param>
        /// <returns>The table reference, or <c>null</c> when no matching table exists.</returns>
        public string? GetTableRange(string tableName) {
            if (string.IsNullOrWhiteSpace(tableName)) {
                return null;
            }

            return _worksheetPart.TableDefinitionParts
                .Select(part => part.Table)
                .FirstOrDefault(table => string.Equals(table?.Name?.Value ?? table?.DisplayName?.Value, tableName, StringComparison.OrdinalIgnoreCase))
                ?.Reference?.Value;
        }

        /// <summary>
        /// Adds an Excel table to the worksheet over the specified range with optional AutoFilter and name validation behavior.
        /// </summary>
        /// <param name="range">Cell range (e.g. "A1:B3") defining the table area.</param>
        /// <param name="hasHeader">Indicates whether the first row is a header row.</param>
        /// <param name="name">Name of the table. If empty, a default name is used.
        /// Examples: "My Table" becomes "My_Table"; "123Report" becomes "_123Report"; spaces and invalid characters are replaced with underscores.
        /// If a name already exists in this workbook, a numeric suffix is appended (e.g., "Table", "Table2").</param>
        /// <param name="style">Table style to apply.</param>
        /// <param name="includeAutoFilter">Whether to include AutoFilter dropdowns in the table headers.</param>
        /// <param name="validationMode">Controls how invalid table names are handled:
        /// <see cref="TableNameValidationMode.Sanitize"/> (default) replaces invalid characters and adjusts the name;
        /// <see cref="TableNameValidationMode.Strict"/> throws descriptive exceptions for invalid names.</param>
        /// <remarks>
        /// Smart AutoFilter handling:
        /// - If includeAutoFilter is true and a worksheet-level AutoFilter exists on the same range, it's moved to the table (preserving any filter criteria).
        /// - If includeAutoFilter is false and a worksheet-level AutoFilter exists, it's removed (Excel doesn't allow both worksheet AutoFilter and a table on the same range).
        /// - Order doesn't matter — the final state will be consistent regardless of operation order.
        /// </remarks>
        public void AddTable(string range, bool hasHeader, string name, TableStyle style, bool includeAutoFilter, TableNameValidationMode validationMode = TableNameValidationMode.Sanitize) {
            AddTableCore(range, hasHeader, name, style, includeAutoFilter, validationMode, ensureRangeCellsExist: true);
        }

        internal string AddTableAndGetName(string range, bool hasHeader, string name, TableStyle style, bool includeAutoFilter, TableNameValidationMode validationMode = TableNameValidationMode.Sanitize, bool ensureRangeCellsExist = true, IReadOnlyList<string>? headerNames = null, bool deferPartSave = false, bool skipExistingTableScan = false) {
            return AddTableCore(range, hasHeader, name, style, includeAutoFilter, validationMode, ensureRangeCellsExist, headerNames, deferPartSave, skipExistingTableScan);
        }

        private string AddTableCore(string range, bool hasHeader, string name, TableStyle style, bool includeAutoFilter, TableNameValidationMode validationMode = TableNameValidationMode.Sanitize, bool ensureRangeCellsExist = true, IReadOnlyList<string>? headerNames = null, bool deferPartSave = false, bool skipExistingTableScan = false) {
            if (string.IsNullOrEmpty(range)) {
                throw new ArgumentNullException(nameof(range));
            }

            string resolvedName = string.Empty;
            WriteLock(() => {
                var cells = range.Split(':');
                if (cells.Length != 2) {
                    throw new ArgumentException("Invalid range format", nameof(range));
                }

                string startRef = cells[0];
                string endRef = cells[1];

                int startColumnIndex = GetColumnIndex(startRef);
                int endColumnIndex = GetColumnIndex(endRef);
                int startRowIndex = GetRowIndex(startRef);
                int endRowIndex = GetRowIndex(endRef);

                if (startColumnIndex > endColumnIndex || startRowIndex > endRowIndex) {
                    throw new ArgumentException($"Invalid range '{range}'. The start cell must be the top-left cell and the end cell must be the bottom-right cell.", nameof(range));
                }

                uint columnsCount = (uint)(endColumnIndex - startColumnIndex + 1);

                if (!skipExistingTableScan) {
                    foreach (var existingPart in _worksheetPart.TableDefinitionParts) {
                        var existingRange = existingPart.Table?.Reference?.Value;
                        if (string.IsNullOrEmpty(existingRange)) continue;
                        var existingCells = existingRange!.Split(':');
                        if (existingCells.Length != 2) continue;
                        string existingStartRef = existingCells[0];
                        string existingEndRef = existingCells[1];

                        int existingStartColumn = GetColumnIndex(existingStartRef);
                        int existingEndColumn = GetColumnIndex(existingEndRef);
                        int existingStartRow = GetRowIndex(existingStartRef);
                        int existingEndRow = GetRowIndex(existingEndRef);

                        bool overlaps = startColumnIndex <= existingEndColumn &&
                                        endColumnIndex >= existingStartColumn &&
                                        startRowIndex <= existingEndRow &&
                                        endRowIndex >= existingStartRow;
                        if (overlaps) {
                            throw new InvalidOperationException("The specified range overlaps with an existing table.");
                        }
                    }
                }

                if (ensureRangeCellsExist) {
                    EnsureRangeCellsExist(startRowIndex, endRowIndex, startColumnIndex, endColumnIndex);
                }

                var swScan = System.Diagnostics.Stopwatch.StartNew();
                uint tableId = _excelDocument.AllocateTableId();
                swScan.Stop();
                EffectiveExecution.ReportTiming("AddTable.ScanExistingIds", swScan.Elapsed);

                // Create the table part with a conventional relationship id (rIdN) to avoid Excel rewriting
                string MakeRelId() {
                    // Find an unused rId# for this worksheet
                    int n = 1;
                    var existing = new System.Collections.Generic.HashSet<string>(
                        _worksheetPart.Parts.Select(p => p.RelationshipId ?? string.Empty),
                        System.StringComparer.Ordinal);
                    string id;
                    do { id = "rId" + n++; } while (existing.Contains(id));
                    return id;
                }
                string relIdNew = MakeRelId();
                var tableDefinitionPart = _worksheetPart.AddNewPart<TableDefinitionPart>(relIdNew);

                if (string.IsNullOrWhiteSpace(name)) {
                    name = $"Table{tableId}";
                }
                name = EnsureValidUniqueTableName(name, validationMode);
                if (string.IsNullOrWhiteSpace(name)) {
                    throw new InvalidOperationException("Table name cannot be empty after validation.");
                }
                resolvedName = name;
                // Reserve the final name in the workbook-level cache for fast uniqueness checks
                _excelDocument.ReserveTableName(name);

                var table = new Table {
                    Id = tableId,
                    Name = name,
                    DisplayName = name,
                    Reference = range,
                    HeaderRowCount = hasHeader ? (uint)1 : (uint)0,
                    TotalsRowShown = false
                };

                var tableColumns = new TableColumns { Count = columnsCount };
                var usedHeaders = new System.Collections.Generic.HashSet<string>(System.StringComparer.OrdinalIgnoreCase);
                for (uint i = 0; i < columnsCount; i++) {
                    string baseName = $"Column{i + 1}";
                    string headerValue = string.Empty;
                    bool headerCellIsSharedString = false;
                    bool headerValueProvided = false;
                    // If the table has headers, try to get the actual header value
                    if (hasHeader && headerNames != null && i < headerNames.Count) {
                        headerValueProvided = true;
                        headerValue = headerNames[(int)i] ?? string.Empty;
                        if (!string.IsNullOrWhiteSpace(headerValue)) {
                            baseName = headerValue;
                        }
                    } else if (hasHeader && startRowIndex > 0) {
                        var headerCell = GetCell(startRowIndex, startColumnIndex + (int)i);
                        if (headerCell != null) {
                            headerCellIsSharedString = headerCell.DataType?.Value == DocumentFormat.OpenXml.Spreadsheet.CellValues.SharedString;
                            headerValue = GetCellText(headerCell);
                            if (!string.IsNullOrWhiteSpace(headerValue)) {
                                baseName = headerValue;
                            }
                        }
                    }
                    string candidate = baseName;
                    int suffix = 2;
                    while (!usedHeaders.Add(candidate)) {
                        candidate = $"{baseName} ({suffix++})";
                    }
                    tableColumns.Append(new TableColumn { Id = i + 1, Name = candidate });

                    bool shouldRewriteHeader = headerValueProvided
                        ? !string.Equals(headerValue, candidate, System.StringComparison.Ordinal)
                        : (!headerCellIsSharedString || !string.Equals(headerValue, candidate, System.StringComparison.Ordinal));
                    if (hasHeader && shouldRewriteHeader) {
                        CellValueCore(startRowIndex, startColumnIndex + (int)i, candidate);
                    }
                }

                // SMART AUTOFILTER HANDLING
                // Check if there's already a worksheet-level AutoFilter on this range
                AutoFilter? existingWorksheetAutoFilter = WorksheetRoot.Elements<AutoFilter>().FirstOrDefault();
                bool hasExistingFilter = existingWorksheetAutoFilter?.Reference?.Value == range;
                bool shouldIncludeAutoFilter = includeAutoFilter && hasHeader;

                if (shouldIncludeAutoFilter) {
                    // User wants AutoFilter on the table
                    if (hasExistingFilter && existingWorksheetAutoFilter != null) {
                        // MIGRATE: Move the existing worksheet AutoFilter to the table (preserving filter criteria)
                        var tableAutoFilter = (AutoFilter)existingWorksheetAutoFilter.CloneNode(true);

                        // Remove from worksheet to avoid conflicts
                        WorksheetRoot.RemoveChild(existingWorksheetAutoFilter);

                        // Add to table
                        table.Append(tableAutoFilter);
                    } else {
                        // No existing filter, just add a new one to the table
                        table.Append(new AutoFilter { Reference = range });
                    }
                } else {
                    // User doesn't want AutoFilter on the table
                    if (hasExistingFilter && existingWorksheetAutoFilter != null) {
                        // REMOVE: Excel doesn't allow worksheet AutoFilter and table on same range
                        // We must remove the worksheet AutoFilter to avoid conflicts
                        WorksheetRoot.RemoveChild(existingWorksheetAutoFilter);
                        // Note: User explicitly set includeAutoFilter=false, so we honor that
                    }
                    // Don't add AutoFilter to the table
                }

                table.Append(tableColumns);

                table.Append(new TableStyleInfo {
                    Name = style.ToString(),
                    ShowFirstColumn = false,
                    ShowLastColumn = false,
                    ShowRowStripes = true,
                    ShowColumnStripes = false
                });

                tableDefinitionPart.Table = table;
                if (deferPartSave) {
                    DeferTableDefinitionPartSave(tableDefinitionPart);
                } else {
                    tableDefinitionPart.Table.Save();
                }

                var tableParts = WorksheetRoot.Elements<TableParts>().FirstOrDefault();
                if (tableParts == null) {
                    tableParts = new TableParts();
                    WorksheetRoot.Append(tableParts);
                }
                var relId = _worksheetPart.GetIdOfPart(tableDefinitionPart);
                // Avoid duplicate TablePart entries
                bool hasPart = tableParts.Elements<TablePart>().Any(tp => tp.Id?.Value == relId);
                if (!hasPart)
                    tableParts.Append(new TablePart { Id = relId });
                tableParts.Count = (uint)tableParts.Elements<TablePart>().Count();

                bool promotedDirectSaveCandidate = _excelDocument.TryPromoteDirectTabularSaveCandidateToTable(
                    this,
                    range,
                    resolvedName,
                    hasHeader,
                    style,
                    includeAutoFilter);

                if (deferPartSave) {
                    if (promotedDirectSaveCandidate) {
                        _excelDocument.PreserveDirectDataSetSaveCandidateForNextDirtyMark();
                    }

                    MarkRequiresSavePreparation();
                } else {
                    WorksheetRoot.Save();
                }
            });

            return resolvedName;
        }

        /// <summary>
        /// Ensures a valid and unique table name according to OfficeIMO rules.
        /// Rules:
        /// - Allowed characters: letters, digits, underscore; spaces become underscores.
        /// - Names cannot start with a digit; an underscore is prefixed if necessary.
        /// - Names are scoped per workbook and checked case-insensitively.
        /// - When <paramref name="mode"/> is <see cref="TableNameValidationMode.Strict"/>, throws for invalid input.
        /// Examples:
        /// - "My Table" ⇒ "My_Table"
        /// - "Sales#2025" ⇒ "Sales_2025"
        /// - "123Report" ⇒ "_123Report"
        /// - If "Table" already exists, next becomes "Table2" ("Table3", ...)
        /// </summary>
        private string EnsureValidUniqueTableName(string name, TableNameValidationMode mode) {
            const int MaxLen = 255; // Excel UI limit; conservative

            if (string.IsNullOrWhiteSpace(name)) {
                if (mode == TableNameValidationMode.Strict)
                    throw new ArgumentException("Table name cannot be null, empty, or whitespace.", nameof(name));
                name = "Table";
            }

            // Sanitize characters
            bool changed = false;
            var sanitized = new System.Text.StringBuilder(name.Length);
            foreach (char ch in name) {
                if (char.IsLetterOrDigit(ch) || ch == '_') {
                    sanitized.Append(ch);
                } else if (char.IsWhiteSpace(ch)) {
                    sanitized.Append('_');
                    changed = true;
                } else {
                    sanitized.Append('_');
                    changed = true;
                }
            }
            if (sanitized.Length == 0) {
                if (mode == TableNameValidationMode.Strict)
                    throw new ArgumentException("Table name must contain at least one letter, digit or underscore.", nameof(name));
                sanitized.Append("Table");
                changed = true;
            }
            if (char.IsDigit(sanitized[0])) {
                if (mode == TableNameValidationMode.Strict)
                    throw new ArgumentException("Table name cannot start with a digit.", nameof(name));
                sanitized.Insert(0, '_');
                changed = true;
            }

            // If any character required sanitation, and in strict mode, throw
            if (mode == TableNameValidationMode.Strict && changed)
                throw new ArgumentException("Table name contains invalid characters or spaces. Allowed: letters, digits, and underscore.", nameof(name));

            // Trim to max length
            if (sanitized.Length > MaxLen) {
                if (mode == TableNameValidationMode.Strict)
                    throw new ArgumentException($"Table name exceeds maximum length of {MaxLen} characters.", nameof(name));
                sanitized.Length = MaxLen;
                changed = true;
            }

            string baseName = sanitized.ToString();

            // Use workbook-level cache for efficiency
            var used = _excelDocument.GetOrInitTableNameCache();
            if (!used.Contains(baseName)) return baseName;

            // Add numeric suffix until unique; trim if needed to fit
            int i = 2;
            while (true) {
                string suffix = i.ToString(System.Globalization.CultureInfo.InvariantCulture);
                int maxBase = Math.Max(1, MaxLen - suffix.Length);
                string trimmedBase = baseName.Length > maxBase ? baseName.Substring(0, maxBase) : baseName;
                string candidate = trimmedBase + suffix;
                if (!used.Contains(candidate)) return candidate;
                i++;
            }
        }

        private void EnsureRangeCellsExist(int startRowIndex, int endRowIndex, int startColumnIndex, int endColumnIndex) {
            var sheetData = GetOrCreateSheetData();

            if (RangeCellsAlreadyExistAsContiguousRows(sheetData, startRowIndex, endRowIndex, startColumnIndex, endColumnIndex)) {
                return;
            }

            var rows = sheetData.Elements<Row>()
                .Where(r => r.RowIndex != null)
                .ToDictionary(r => (int)r.RowIndex!.Value);

            int columnCount = endColumnIndex - startColumnIndex + 1;
            bool useBitMask = columnCount <= 64;
            ulong fullMask = columnCount == 64 ? ulong.MaxValue : (1UL << columnCount) - 1UL;

            for (int rowIndex = startRowIndex; rowIndex <= endRowIndex; rowIndex++) {
                if (!rows.TryGetValue(rowIndex, out Row? row)) {
                    row = GetOrCreateRowElement(sheetData, rowIndex);
                    rows[rowIndex] = row;
                }

                if (useBitMask) {
                    ulong existingMask = 0UL;
                    foreach (var cell in row.Elements<Cell>()) {
                        var cellReference = cell.CellReference?.Value;
                        if (string.IsNullOrEmpty(cellReference)) {
                            continue;
                        }

                        int columnIndex = GetColumnIndex(cellReference!);
                        if (columnIndex < startColumnIndex || columnIndex > endColumnIndex) {
                            continue;
                        }

                        existingMask |= 1UL << (columnIndex - startColumnIndex);
                        if (existingMask == fullMask) {
                            break;
                        }
                    }

                    if (existingMask == fullMask) {
                        continue;
                    }

                    for (int offset = 0; offset < columnCount; offset++) {
                        if ((existingMask & (1UL << offset)) == 0UL) {
                            GetCell(rowIndex, startColumnIndex + offset);
                        }
                    }

                    continue;
                }

                var existingColumns = new HashSet<int>();
                foreach (var cell in row.Elements<Cell>()) {
                    var cellReference = cell.CellReference?.Value;
                    if (!string.IsNullOrEmpty(cellReference)) {
                        existingColumns.Add(GetColumnIndex(cellReference!));
                    }
                }

                for (int columnIndex = startColumnIndex; columnIndex <= endColumnIndex; columnIndex++) {
                    if (!existingColumns.Contains(columnIndex)) {
                        GetCell(rowIndex, columnIndex);
                    }
                }
            }
        }

        private static bool RangeCellsAlreadyExistAsContiguousRows(SheetData sheetData, int startRowIndex, int endRowIndex, int startColumnIndex, int endColumnIndex) {
            int expectedRow = startRowIndex;
            int expectedColumnCount = endColumnIndex - startColumnIndex + 1;

            foreach (var row in sheetData.Elements<Row>()) {
                if (row.RowIndex == null) {
                    continue;
                }

                int rowIndex = (int)row.RowIndex.Value;
                if (rowIndex < startRowIndex) {
                    continue;
                }

                if (rowIndex > endRowIndex) {
                    break;
                }

                if (rowIndex != expectedRow || !RowHasExactContiguousCells(row, startColumnIndex, endColumnIndex, expectedColumnCount)) {
                    return false;
                }

                expectedRow++;
            }

            return expectedRow > endRowIndex;
        }

        private static bool RowHasExactContiguousCells(Row row, int startColumnIndex, int endColumnIndex, int expectedColumnCount) {
            int expectedColumn = startColumnIndex;
            int cellCount = 0;

            foreach (var cell in row.Elements<Cell>()) {
                string? cellReference = cell.CellReference?.Value;
                if (string.IsNullOrEmpty(cellReference)) {
                    return false;
                }

                if (GetColumnIndex(cellReference!) != expectedColumn) {
                    return false;
                }

                expectedColumn++;
                cellCount++;
                if (cellCount > expectedColumnCount) {
                    return false;
                }
            }

            return cellCount == expectedColumnCount && expectedColumn == endColumnIndex + 1;
        }

    }
}
