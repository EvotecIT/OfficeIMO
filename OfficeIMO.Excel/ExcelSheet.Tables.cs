using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using SixLabors.Fonts;
using System;
using System.Collections;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;
using SixLaborsColor = SixLabors.ImageSharp.Color;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
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
                AutoFilter existing = worksheet.Elements<AutoFilter>().FirstOrDefault();
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
            AddTable(range, hasHeader, name, style, true);
        }

        /// <summary>
        /// Adds an Excel table to the worksheet over the specified range with optional AutoFilter.
        /// </summary>
        /// <param name="range">Cell range (e.g. "A1:B3") defining the table area.</param>
        /// <param name="hasHeader">Indicates whether the first row is a header row.</param>
        /// <param name="name">Name of the table. If empty, a default name is used.</param>
        /// <param name="style">Table style to apply.</param>
        /// <param name="includeAutoFilter">Whether to include AutoFilter dropdowns in the table headers.</param>
        /// <remarks>
        /// Smart AutoFilter handling:
        /// - If includeAutoFilter is true and a worksheet-level AutoFilter exists on the same range, 
        ///   it's moved to the table (preserving any filter criteria)
        /// - If includeAutoFilter is false and a worksheet-level AutoFilter exists, it's removed 
        ///   (Excel doesn't allow both worksheet AutoFilter and table on same range)
        /// - Order doesn't matter - the final state will be consistent regardless of operation order
        /// </remarks>
        /// <exception cref="ArgumentNullException">Thrown when <paramref name="range"/> is null or empty.</exception>
        /// <exception cref="ArgumentException">Thrown when <paramref name="range"/> is not in a valid format.</exception>
        /// <exception cref="InvalidOperationException">Thrown when the specified range overlaps with an existing table.</exception>
        public void AddTable(string range, bool hasHeader, string name, TableStyle style, bool includeAutoFilter) {
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
                    var existingCells = existingRange.Split(':');
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

                // Generate unique table ID atomically
                uint tableId;
                lock (_tableIdLock) {
                    // Get max existing table ID to ensure uniqueness
                    uint maxExistingId = 0;
                    foreach (var part in _worksheetPart.TableDefinitionParts) {
                        if (part.Table?.Id?.Value != null && part.Table.Id.Value > maxExistingId) {
                            maxExistingId = part.Table.Id.Value;
                        }
                    }
                    tableId = Math.Max((uint)_nextTableId, maxExistingId + 1);
                    _nextTableId = (int)tableId + 1;
                }

                var tableDefinitionPart = _worksheetPart.AddNewPart<TableDefinitionPart>();

                if (string.IsNullOrEmpty(name)) {
                    name = $"Table{tableId}";
                }

                var table = new Table {
                    Id = tableId,
                    Name = name,
                    DisplayName = name,
                    Reference = range,
                    HeaderRowCount = hasHeader ? (uint)1 : (uint)0,
                    TotalsRowShown = false
                };

                var tableColumns = new TableColumns { Count = columnsCount };
                for (uint i = 0; i < columnsCount; i++) {
                    string columnName = $"Column{i + 1}";

                    // If the table has headers, try to get the actual header value
                    if (hasHeader && startRowIndex > 0) {
                        var headerCell = GetCell(startRowIndex, startColumnIndex + (int)i);
                        if (headerCell != null) {
                            string headerValue = GetCellText(headerCell);
                            if (!string.IsNullOrWhiteSpace(headerValue)) {
                                columnName = headerValue;
                            }
                        }
                    }

                    tableColumns.Append(new TableColumn { Id = i + 1, Name = columnName });
                }
                
                // SMART AUTOFILTER HANDLING
                // Check if there's already a worksheet-level AutoFilter on this range
                AutoFilter existingWorksheetAutoFilter = _worksheetPart.Worksheet.Elements<AutoFilter>().FirstOrDefault();
                bool hasExistingFilter = existingWorksheetAutoFilter != null && existingWorksheetAutoFilter.Reference?.Value == range;
                
                if (includeAutoFilter) {
                    // User wants AutoFilter on the table
                    if (hasExistingFilter) {
                        // MIGRATE: Move the existing worksheet AutoFilter to the table (preserving filter criteria)
                        // First clone the existing AutoFilter to preserve any filter criteria
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
                    if (hasExistingFilter) {
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
                    tableParts = new TableParts { Count = 1 };
                    _worksheetPart.Worksheet.Append(tableParts);
                } else {
                    tableParts.Count = (tableParts.Count ?? 0) + 1;
                }

                var relId = _worksheetPart.GetIdOfPart(tableDefinitionPart);
                tableParts.Append(new TablePart { Id = relId });

                _worksheetPart.Worksheet.Save();
            });
        }

    }
}

