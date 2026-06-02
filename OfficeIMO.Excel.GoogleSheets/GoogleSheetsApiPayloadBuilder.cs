using System.Globalization;

namespace OfficeIMO.Excel.GoogleSheets {
    internal static partial class GoogleSheetsApiPayloadBuilder {
        internal static GoogleSheetsApiCreateSpreadsheetPayload BuildCreateSpreadsheetPayload(GoogleSheetsBatch batch) {
            return BuildCreateSpreadsheetPayload(batch, BuildSheetIdMap(batch));
        }

        internal static GoogleSheetsApiCreateSpreadsheetPayload BuildCreateSpreadsheetPayload(
            GoogleSheetsBatch batch,
            IReadOnlyDictionary<string, int> sheetIds) {
            if (batch == null) throw new ArgumentNullException(nameof(batch));
            if (sheetIds == null) throw new ArgumentNullException(nameof(sheetIds));

            var payload = new GoogleSheetsApiCreateSpreadsheetPayload {
                Properties = new GoogleSheetsApiSpreadsheetPropertiesPayload {
                    Title = batch.Title,
                }
            };

            foreach (var sheet in batch.Requests.OfType<GoogleSheetsAddSheetRequest>()) {
                if (!sheetIds.TryGetValue(sheet.SheetName, out var sheetId)) {
                    continue;
                }

                payload.Sheets.Add(new GoogleSheetsApiSheetPayload {
                    Properties = new GoogleSheetsApiSheetPropertiesPayload {
                        SheetId = sheetId,
                        Title = sheet.SheetName,
                        Index = sheet.SheetIndex,
                        Hidden = sheet.Hidden,
                        RightToLeft = sheet.RightToLeft ? true : (bool?)null,
                        TabColor = BuildColor(sheet.TabColorArgb),
                        GridProperties = new GoogleSheetsApiGridPropertiesPayload {
                            FrozenRowCount = sheet.FrozenRowCount > 0 ? sheet.FrozenRowCount : (int?)null,
                            FrozenColumnCount = sheet.FrozenColumnCount > 0 ? sheet.FrozenColumnCount : (int?)null,
                        }
                    }
                });
            }

            return payload;
        }

        internal static GoogleSheetsApiBatchUpdatePayload BuildBatchUpdatePayload(GoogleSheetsBatch batch) {
            return BuildBatchUpdatePayload(batch, BuildSheetIdMap(batch), null);
        }

        internal static GoogleSheetsApiBatchUpdatePayload BuildBatchUpdatePayload(
            GoogleSheetsBatch batch,
            IReadOnlyDictionary<string, int> sheetIds) {
            return BuildBatchUpdatePayload(batch, sheetIds, null);
        }

        internal static GoogleSheetsApiBatchUpdatePayload BuildBatchUpdatePayload(
            GoogleSheetsBatch batch,
            IReadOnlyDictionary<string, int> sheetIds,
            string? spreadsheetId) {
            if (batch == null) throw new ArgumentNullException(nameof(batch));
            if (sheetIds == null) throw new ArgumentNullException(nameof(sheetIds));
            var payload = new GoogleSheetsApiBatchUpdatePayload();

            foreach (var request in batch.Requests) {
                switch (request) {
                    case GoogleSheetsAddSheetRequest:
                        break;
                    case GoogleSheetsUpdateDimensionPropertiesRequest updateDimension:
                        if (!sheetIds.TryGetValue(updateDimension.SheetName, out var dimensionSheetId)) {
                            continue;
                        }

                        payload.Requests.Add(new GoogleSheetsApiRequestPayload {
                            UpdateDimensionProperties = new GoogleSheetsApiUpdateDimensionPropertiesRequestPayload {
                                Range = new GoogleSheetsApiDimensionRangePayload {
                                    SheetId = dimensionSheetId,
                                    Dimension = updateDimension.DimensionKind == GoogleSheetsDimensionKind.Columns ? "COLUMNS" : "ROWS",
                                    StartIndex = updateDimension.StartIndex,
                                    EndIndex = updateDimension.EndIndexExclusive,
                                },
                                Properties = new GoogleSheetsApiDimensionPropertiesPayload {
                                    PixelSize = updateDimension.PixelSize,
                                    HiddenByUser = updateDimension.Hidden ? true : (bool?)null,
                                },
                                Fields = BuildDimensionFields(updateDimension),
                            }
                        });
                        break;
                    case GoogleSheetsUpdateCellsRequest updateCells:
                        if (!sheetIds.TryGetValue(updateCells.SheetName, out var updateSheetId)) {
                            continue;
                        }

                        AppendUpdateCellsRequests(batch, payload, updateSheetId, updateCells, sheetIds, spreadsheetId);
                        break;
                    case GoogleSheetsMergeCellsRequest merge:
                        if (!sheetIds.TryGetValue(merge.SheetName, out var mergeSheetId)) {
                            continue;
                        }

                        payload.Requests.Add(new GoogleSheetsApiRequestPayload {
                            MergeCells = new GoogleSheetsApiMergeCellsRequestPayload {
                                Range = new GoogleSheetsApiGridRangePayload {
                                    SheetId = mergeSheetId,
                                    StartRowIndex = merge.StartRowIndex,
                                    EndRowIndex = merge.EndRowIndexExclusive,
                                    StartColumnIndex = merge.StartColumnIndex,
                                    EndColumnIndex = merge.EndColumnIndexExclusive,
                                },
                                MergeType = "MERGE_ALL",
                            }
                        });
                        break;
                    case GoogleSheetsAddNamedRangeRequest namedRange:
                        if (!TryBuildNamedRangePayload(sheetIds, namedRange, out var namedRangePayload)) {
                            batch.Report.Add(
                                OfficeIMO.GoogleWorkspace.TranslationSeverity.Warning,
                                "NamedRanges",
                                $"Named range '{namedRange.Name}' could not be converted into a Google Sheets GridRange payload.");
                            continue;
                        }

                        payload.Requests.Add(new GoogleSheetsApiRequestPayload {
                            AddNamedRange = new GoogleSheetsApiAddNamedRangeRequestPayload {
                                NamedRange = namedRangePayload,
                            }
                        });
                        break;
                    case GoogleSheetsAddProtectedRangeRequest protectedRange:
                        if (!sheetIds.TryGetValue(protectedRange.SheetName, out var protectedSheetId)) {
                            continue;
                        }

                        payload.Requests.Add(new GoogleSheetsApiRequestPayload {
                            AddProtectedRange = new GoogleSheetsApiAddProtectedRangeRequestPayload {
                                ProtectedRange = new GoogleSheetsApiProtectedRangePayload {
                                    Range = new GoogleSheetsApiGridRangePayload {
                                        SheetId = protectedSheetId,
                                    },
                                    Description = protectedRange.Description,
                                    WarningOnly = protectedRange.WarningOnly,
                                }
                            }
                        });
                        break;
                    case GoogleSheetsSetBasicFilterRequest basicFilter:
                        if (!TryBuildGridRange(sheetIds, basicFilter.SheetName, basicFilter.A1Range, out var basicFilterRange)) {
                            batch.Report.Add(
                                OfficeIMO.GoogleWorkspace.TranslationSeverity.Warning,
                                "Filters",
                                $"Basic filter range '{basicFilter.A1Range}' on sheet '{basicFilter.SheetName}' could not be converted into a Google Sheets GridRange payload.");
                            continue;
                        }

                        payload.Requests.Add(new GoogleSheetsApiRequestPayload {
                            SetBasicFilter = new GoogleSheetsApiSetBasicFilterRequestPayload {
                                Filter = new GoogleSheetsApiBasicFilterPayload {
                                    Range = basicFilterRange,
                                    Criteria = BuildFilterCriteriaMap(basicFilter.Criteria),
                                }
                            }
                        });
                        break;
                    case GoogleSheetsAddFilterViewRequest filterView:
                        if (!TryBuildGridRange(sheetIds, filterView.SheetName, filterView.A1Range, out var filterViewRange)) {
                            batch.Report.Add(
                                OfficeIMO.GoogleWorkspace.TranslationSeverity.Warning,
                                "Filters",
                                $"Filter view range '{filterView.A1Range}' on sheet '{filterView.SheetName}' could not be converted into a Google Sheets GridRange payload.");
                            continue;
                        }

                        payload.Requests.Add(new GoogleSheetsApiRequestPayload {
                            AddFilterView = new GoogleSheetsApiAddFilterViewRequestPayload {
                                Filter = new GoogleSheetsApiFilterViewPayload {
                                    Title = filterView.Title,
                                    Range = filterViewRange,
                                    Criteria = BuildFilterCriteriaMap(filterView.Criteria),
                                }
                            }
                        });
                        break;
                    case GoogleSheetsAddTableRequest table:
                        if (!TryBuildGridRange(sheetIds, table.SheetName, table.A1Range, out var tableRange)) {
                            batch.Report.Add(
                                OfficeIMO.GoogleWorkspace.TranslationSeverity.Warning,
                                "Tables",
                                $"Table range '{table.A1Range}' on sheet '{table.SheetName}' could not be converted into a Google Sheets GridRange payload.");
                            continue;
                        }

                        payload.Requests.Add(new GoogleSheetsApiRequestPayload {
                            AddTable = new GoogleSheetsApiAddTableRequestPayload {
                                Table = new GoogleSheetsApiTablePayload {
                                    Name = table.TableName,
                                    Range = tableRange,
                                    RowsProperties = BuildTableRowsProperties(table),
                                    ColumnProperties = table.Columns.Select(column => new GoogleSheetsApiTableColumnPropertiesPayload {
                                        Name = column.Name,
                                        ColumnType = column.ColumnType,
                                        DataValidationRule = BuildDataValidationRule(column.DataValidationRule),
                                    }).ToList(),
                                }
                            }
                        });
                        break;
                }
            }

            return payload;
        }

        internal static IReadOnlyDictionary<string, int> BuildSheetIdMap(
            GoogleSheetsBatch batch,
            IEnumerable<int>? reservedSheetIds = null) {
            if (batch == null) throw new ArgumentNullException(nameof(batch));

            var reserved = new HashSet<int>(reservedSheetIds ?? Array.Empty<int>());
            var nextId = reserved.Count == 0 ? 1 : reserved.Max() + 1;
            var map = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

            foreach (var sheet in batch.Requests.OfType<GoogleSheetsAddSheetRequest>().OrderBy(s => s.SheetIndex)) {
                while (reserved.Contains(nextId)) {
                    nextId++;
                }

                map[sheet.SheetName] = nextId;
                reserved.Add(nextId);
                nextId++;
            }

            return map;
        }

        internal static GoogleSheetsApiBatchUpdatePayload BuildReplaceSpreadsheetPayload(
            GoogleSheetsBatch batch,
            IReadOnlyCollection<int> existingSheetIds,
            IReadOnlyDictionary<string, int> desiredSheetIds) {
            if (batch == null) throw new ArgumentNullException(nameof(batch));
            if (existingSheetIds == null) throw new ArgumentNullException(nameof(existingSheetIds));
            if (desiredSheetIds == null) throw new ArgumentNullException(nameof(desiredSheetIds));

            var payload = new GoogleSheetsApiBatchUpdatePayload();
            payload.Requests.Add(new GoogleSheetsApiRequestPayload {
                UpdateSpreadsheetProperties = new GoogleSheetsApiUpdateSpreadsheetPropertiesRequestPayload {
                    Properties = new GoogleSheetsApiSpreadsheetPropertiesPayload {
                        Title = batch.Title,
                    },
                    Fields = "title",
                }
            });

            foreach (var sheet in batch.Requests.OfType<GoogleSheetsAddSheetRequest>().OrderBy(s => s.SheetIndex)) {
                if (!desiredSheetIds.TryGetValue(sheet.SheetName, out var desiredSheetId)) {
                    continue;
                }

                payload.Requests.Add(new GoogleSheetsApiRequestPayload {
                    AddSheet = new GoogleSheetsApiAddSheetRequestPayload {
                        Properties = new GoogleSheetsApiSheetPropertiesPayload {
                            SheetId = desiredSheetId,
                            Title = sheet.SheetName,
                            Index = sheet.SheetIndex,
                            Hidden = sheet.Hidden,
                            RightToLeft = sheet.RightToLeft ? true : (bool?)null,
                            TabColor = BuildColor(sheet.TabColorArgb),
                            GridProperties = new GoogleSheetsApiGridPropertiesPayload {
                                FrozenRowCount = sheet.FrozenRowCount > 0 ? sheet.FrozenRowCount : (int?)null,
                                FrozenColumnCount = sheet.FrozenColumnCount > 0 ? sheet.FrozenColumnCount : (int?)null,
                            }
                        }
                    }
                });
            }

            foreach (var existingSheetId in existingSheetIds) {
                payload.Requests.Add(new GoogleSheetsApiRequestPayload {
                    DeleteSheet = new GoogleSheetsApiDeleteSheetRequestPayload {
                        SheetId = existingSheetId,
                    }
                });
            }

            return payload;
        }

        private static void AppendUpdateCellsRequests(
            GoogleSheetsBatch batch,
            GoogleSheetsApiBatchUpdatePayload payload,
            int sheetId,
            GoogleSheetsUpdateCellsRequest request,
            IReadOnlyDictionary<string, int> sheetIds,
            string? spreadsheetId) {
            var groupedRows = request.Cells
                .OrderBy(cell => cell.RowIndex)
                .ThenBy(cell => cell.ColumnIndex)
                .GroupBy(cell => cell.RowIndex);

            foreach (var rowGroup in groupedRows) {
                var currentSegment = new List<GoogleSheetsCellData>();
                int expectedColumn = -1;

                foreach (var cell in rowGroup) {
                    if (currentSegment.Count == 0) {
                        currentSegment.Add(cell);
                        expectedColumn = cell.ColumnIndex + 1;
                        continue;
                    }

                    if (cell.ColumnIndex == expectedColumn) {
                        currentSegment.Add(cell);
                        expectedColumn = cell.ColumnIndex + 1;
                        continue;
                    }

                    AddCellSegment(batch, payload, sheetId, rowGroup.Key, currentSegment, request.SheetName, sheetIds, spreadsheetId);
                    currentSegment = new List<GoogleSheetsCellData> { cell };
                    expectedColumn = cell.ColumnIndex + 1;
                }

                if (currentSegment.Count > 0) {
                    AddCellSegment(batch, payload, sheetId, rowGroup.Key, currentSegment, request.SheetName, sheetIds, spreadsheetId);
                }
            }
        }

        private static void AddCellSegment(
            GoogleSheetsBatch batch,
            GoogleSheetsApiBatchUpdatePayload payload,
            int sheetId,
            int rowIndex,
            IReadOnlyList<GoogleSheetsCellData> cells,
            string sourceSheetName,
            IReadOnlyDictionary<string, int> sheetIds,
            string? spreadsheetId) {
            var rowData = new GoogleSheetsApiRowDataPayload();
            bool includeFormat = false;
            bool includeNote = false;
            bool includeValidation = false;

            foreach (var cell in cells) {
                var apiCell = BuildCellData(batch, cell, sourceSheetName, sheetIds, spreadsheetId, out var hasFormat, out var hasNote, out var hasValidation);
                rowData.Values.Add(apiCell);
                includeFormat |= hasFormat;
                includeNote |= hasNote;
                includeValidation |= hasValidation;
            }

            payload.Requests.Add(new GoogleSheetsApiRequestPayload {
                UpdateCells = new GoogleSheetsApiUpdateCellsRequestPayload {
                    Start = new GoogleSheetsApiGridCoordinatePayload {
                        SheetId = sheetId,
                        RowIndex = rowIndex,
                        ColumnIndex = cells[0].ColumnIndex,
                    },
                    Rows = new List<GoogleSheetsApiRowDataPayload> { rowData },
                    Fields = BuildUpdateCellsFields(includeFormat, includeNote, includeValidation),
                }
            });
        }

    }
}
