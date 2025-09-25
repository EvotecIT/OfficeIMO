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

            var totalsByHeader = new System.Collections.Generic.Dictionary<string, DocumentFormat.OpenXml.Spreadsheet.TotalsRowFunctionValues>(byHeader, System.StringComparer.OrdinalIgnoreCase);
            WriteLock(() => {
                foreach (var tdp in _worksheetPart.TableDefinitionParts) {
                    var table = tdp.Table;
                    if (table?.Reference?.Value != range) continue;
                    table.TotalsRowShown = true;
                    var headerNames = table.TableColumns?.Elements<TableColumn>().Select(tc => tc.Name?.Value ?? string.Empty).ToList() ?? new System.Collections.Generic.List<string>();
                    int idx = 0;
                    foreach (var tc in table.TableColumns!.Elements<TableColumn>()) {
                        var name = headerNames[idx++];
                        if (totalsByHeader.TryGetValue(name, out var fn)) {
                            tc.TotalsRowFunction = fn;
                        }
                    }
                    tdp.Table.Save();
                    break;
                }
                _worksheetPart.Worksheet.Save();
            });
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
                    if (table?.Reference?.Value == range) {
                        // Found a table on the same range - add/update its AutoFilter

                        // First, remove any worksheet-level AutoFilter to avoid conflicts
                        var worksheetAutoFilter = _worksheetPart.Worksheet.Elements<AutoFilter>().FirstOrDefault();
                        if (worksheetAutoFilter != null && worksheetAutoFilter.Reference?.Value == range) {
                            _worksheetPart.Worksheet.RemoveChild(worksheetAutoFilter);
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

                        tableDefinitionPart.Table.Save();
                        return; // Exit early - we've handled the AutoFilter via the table
                    }
                }

                // No table found, add worksheet-level AutoFilter
                Worksheet worksheet = _worksheetPart.Worksheet;

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
            AddTable(range, hasHeader, name, style, includeAutoFilter: true);
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
            if (string.IsNullOrEmpty(range)) {
                throw new ArgumentNullException(nameof(range));
            }

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

                uint columnsCount = (uint)(endColumnIndex - startColumnIndex + 1);

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

                // Ensure all cells in the table range exist (but don't set empty values)
                for (int row = startRowIndex; row <= endRowIndex; row++) {
                    for (int column = startColumnIndex; column <= endColumnIndex; column++) {
                        var cell = GetCell(row, column);
                        // Just ensure the cell exists, don't set a value if it's empty
                        // Excel will handle empty cells in tables correctly
                    }
                }

                // Generate unique table ID atomically (must be unique across the entire workbook)
                uint tableId;
                var swScan = System.Diagnostics.Stopwatch.StartNew();
                lock (_tableIdLock) {
                    // Get max existing table ID across all sheets to ensure uniqueness when opening existing files
                    uint maxExistingId = 0;
                    var wbPart = _spreadSheetDocument.WorkbookPart;
                    if (wbPart != null) {
                        foreach (var ws in wbPart.WorksheetParts) {
                            foreach (var part in ws.TableDefinitionParts) {
                                var idv = part.Table?.Id?.Value;
                                if (idv != null && idv.Value > maxExistingId)
                                    maxExistingId = idv.Value;
                            }
                        }
                    }
                    // Ensure _nextTableId always advances beyond any seen IDs
                    var next = Math.Max(_nextTableId, (int)(maxExistingId + 1));
                    tableId = (uint)next;
                    _nextTableId = next + 1;
                }
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
                    // If the table has headers, try to get the actual header value
                    if (hasHeader && startRowIndex > 0) {
                        var headerCell = GetCell(startRowIndex, startColumnIndex + (int)i);
                        if (headerCell != null) {
                            string headerValue = GetCellText(headerCell);
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
                }

                // SMART AUTOFILTER HANDLING
                // Check if there's already a worksheet-level AutoFilter on this range
                AutoFilter? existingWorksheetAutoFilter = _worksheetPart.Worksheet.Elements<AutoFilter>().FirstOrDefault();
                bool hasExistingFilter = existingWorksheetAutoFilter?.Reference?.Value == range;

                if (includeAutoFilter) {
                    // User wants AutoFilter on the table
                    if (hasExistingFilter && existingWorksheetAutoFilter != null) {
                        // MIGRATE: Move the existing worksheet AutoFilter to the table (preserving filter criteria)
                        var tableAutoFilter = (AutoFilter)existingWorksheetAutoFilter.CloneNode(true);

                        // Remove from worksheet to avoid conflicts
                        _worksheetPart.Worksheet.RemoveChild(existingWorksheetAutoFilter);

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
                        _worksheetPart.Worksheet.RemoveChild(existingWorksheetAutoFilter);
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
                tableDefinitionPart.Table.Save();

                var tableParts = _worksheetPart.Worksheet.Elements<TableParts>().FirstOrDefault();
                if (tableParts == null) {
                    tableParts = new TableParts();
                    _worksheetPart.Worksheet.Append(tableParts);
                }
                var relId = _worksheetPart.GetIdOfPart(tableDefinitionPart);
                // Avoid duplicate TablePart entries
                bool hasPart = tableParts.Elements<TablePart>().Any(tp => tp.Id?.Value == relId);
                if (!hasPart)
                    tableParts.Append(new TablePart { Id = relId });
                tableParts.Count = (uint)tableParts.Elements<TablePart>().Count();

                _worksheetPart.Worksheet.Save();
            });
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
                    if (ch != '_') changed = true;
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

    }
}
