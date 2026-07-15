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
            ApplySpreadsheetProperties(payload.Properties, batch.Requests.OfType<GoogleSheetsUpdateSpreadsheetPropertiesRequest>().LastOrDefault());

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
                            HideGridlines = sheet.HideGridlines,
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
            string? spreadsheetId,
            bool includeCellValues = true) {
            if (batch == null) throw new ArgumentNullException(nameof(batch));
            if (sheetIds == null) throw new ArgumentNullException(nameof(sheetIds));
            var payload = new GoogleSheetsApiBatchUpdatePayload();

            foreach (var request in batch.Requests) {
                switch (request) {
                    case GoogleSheetsAddSheetRequest:
                        break;
                    case GoogleSheetsUpdateSpreadsheetPropertiesRequest:
                        // Create and replace payloads apply these properties atomically with the title.
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

                        AppendUpdateCellsRequests(batch, payload, updateSheetId, updateCells, sheetIds, spreadsheetId, includeCellValues);
                        break;
                    case GoogleSheetsSetDataValidationRequest validation:
                        if (!sheetIds.TryGetValue(validation.SheetName, out var validationSheetId)) {
                            continue;
                        }

                        payload.Requests.Add(new GoogleSheetsApiRequestPayload {
                            SetDataValidation = new GoogleSheetsApiSetDataValidationRequestPayload {
                                Range = new GoogleSheetsApiGridRangePayload {
                                    SheetId = validationSheetId,
                                    StartRowIndex = validation.StartRowIndex,
                                    EndRowIndex = validation.EndRowIndexExclusive,
                                    StartColumnIndex = validation.StartColumnIndex,
                                    EndColumnIndex = validation.EndColumnIndexExclusive,
                                },
                                Rule = BuildDataValidationRule(validation.Rule)!,
                            }
                        });
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
                                    Editors = protectedRange.EditorEmailAddresses.Count > 0 || protectedRange.DomainUsersCanEdit
                                        ? new GoogleSheetsApiEditorsPayload {
                                            Users = protectedRange.EditorEmailAddresses.Count > 0 ? protectedRange.EditorEmailAddresses.ToList() : null,
                                            DomainUsersCanEdit = protectedRange.DomainUsersCanEdit,
                                        }
                                        : null,
                                    UnprotectedRanges = BuildUnprotectedRanges(sheetIds, protectedRange),
                                }
                            }
                        });
                        break;
                    case GoogleSheetsAddDimensionGroupRequest group:
                        if (!sheetIds.TryGetValue(group.SheetName, out var groupSheetId)) continue;
                        payload.Requests.Add(new GoogleSheetsApiRequestPayload {
                            AddDimensionGroup = new GoogleSheetsApiAddDimensionGroupRequestPayload {
                                Range = new GoogleSheetsApiDimensionRangePayload {
                                    SheetId = groupSheetId,
                                    Dimension = group.DimensionKind == GoogleSheetsDimensionKind.Columns ? "COLUMNS" : "ROWS",
                                    StartIndex = group.StartIndex,
                                    EndIndex = group.EndIndexExclusive,
                                }
                            }
                        });
                        break;
                    case GoogleSheetsAddConditionalFormatRuleRequest conditional:
                        if (!TryBuildGridRange(sheetIds, conditional.SheetName, conditional.A1Range, out var conditionalRange)) continue;
                        var conditionalPayload = new GoogleSheetsApiConditionalFormatRulePayload {
                            BooleanRule = new GoogleSheetsApiBooleanRulePayload {
                                Condition = new GoogleSheetsApiBooleanConditionPayload {
                                    Type = conditional.ConditionType,
                                    Values = conditional.Values.Count == 0 ? null : conditional.Values.Select(value => new GoogleSheetsApiConditionValuePayload { UserEnteredValue = value }).ToList(),
                                },
                                Format = BuildCellFormat(conditional.Format) ?? new GoogleSheetsApiCellFormatPayload(),
                            }
                        };
                        conditionalPayload.Ranges.Add(conditionalRange);
                        payload.Requests.Add(new GoogleSheetsApiRequestPayload {
                            AddConditionalFormatRule = new GoogleSheetsApiAddConditionalFormatRuleRequestPayload { Rule = conditionalPayload, Index = conditional.Index }
                        });
                        break;
                    case GoogleSheetsAddChartRequest chart:
                        if (!TryBuildChartPayload(sheetIds, chart, out var chartPayload)) continue;
                        payload.Requests.Add(new GoogleSheetsApiRequestPayload { AddChart = chartPayload });
                        break;
                    case GoogleSheetsAddPivotTableRequest pivot:
                        if (!TryBuildPivotTablePayload(sheetIds, pivot, out var pivotUpdate)) continue;
                        payload.Requests.Add(pivotUpdate);
                        break;
                    case GoogleSheetsAddDeveloperMetadataRequest metadata:
                        payload.Requests.Add(new GoogleSheetsApiRequestPayload {
                            AddDeveloperMetadata = new GoogleSheetsApiAddDeveloperMetadataRequestPayload {
                                DeveloperMetadata = new GoogleSheetsApiDeveloperMetadataPayload {
                                    MetadataKey = metadata.Key,
                                    MetadataValue = metadata.Value,
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
                                        ColumnIndex = column.ColumnIndex,
                                        ColumnName = column.Name,
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

        internal static GoogleSheetsApiBatchUpdateValuesPayload BuildValuesBatchUpdatePayload(
            GoogleSheetsBatch batch,
            IReadOnlyDictionary<string, int> sheetIds,
            string? spreadsheetId) {
            if (batch == null) throw new ArgumentNullException(nameof(batch));
            if (sheetIds == null) throw new ArgumentNullException(nameof(sheetIds));
            var payload = new GoogleSheetsApiBatchUpdateValuesPayload();
            foreach (GoogleSheetsUpdateCellsRequest update in batch.Requests.OfType<GoogleSheetsUpdateCellsRequest>()) {
                if (!sheetIds.ContainsKey(update.SheetName)) continue;
                foreach (IGrouping<int, GoogleSheetsCellData> row in update.Cells
                    .OrderBy(cell => cell.RowIndex)
                    .ThenBy(cell => cell.ColumnIndex)
                    .GroupBy(cell => cell.RowIndex)) {
                    var segment = new List<GoogleSheetsCellData>();
                    int expected = -1;
                    foreach (GoogleSheetsCellData cell in row) {
                        if (segment.Count == 0 || cell.ColumnIndex == expected) {
                            segment.Add(cell);
                        } else {
                            AddValueRange(payload, batch, update.SheetName, row.Key, segment, sheetIds, spreadsheetId);
                            segment = new List<GoogleSheetsCellData> { cell };
                        }
                        expected = cell.ColumnIndex + 1;
                    }
                    if (segment.Count > 0) AddValueRange(payload, batch, update.SheetName, row.Key, segment, sheetIds, spreadsheetId);
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
            IReadOnlyDictionary<int, string> existingSheets,
            IReadOnlyDictionary<string, int> desiredSheetIds) {
            if (batch == null) throw new ArgumentNullException(nameof(batch));
            if (existingSheets == null) throw new ArgumentNullException(nameof(existingSheets));
            if (desiredSheetIds == null) throw new ArgumentNullException(nameof(desiredSheetIds));

            var payload = new GoogleSheetsApiBatchUpdatePayload();
            var desiredSheets = batch.Requests
                .OfType<GoogleSheetsAddSheetRequest>()
                .OrderBy(sheet => sheet.SheetIndex)
                .ToList();
            if (desiredSheets.Count == 0) {
                throw new InvalidOperationException("Spreadsheet replacement requires at least one desired sheet.");
            }

            int keeperSheetId = existingSheets.Keys
                .Concat(desiredSheetIds.Values)
                .DefaultIfEmpty(0)
                .Max() + 1;
            string keeperTitle = BuildUniqueReplacementKeeperTitle(
                existingSheets.Values.Concat(desiredSheets.Select(sheet => sheet.SheetName)));

            var replaceProperties = new GoogleSheetsApiSpreadsheetPropertiesPayload {
                Title = batch.Title,
            };
            GoogleSheetsUpdateSpreadsheetPropertiesRequest? requestedProperties = batch.Requests.OfType<GoogleSheetsUpdateSpreadsheetPropertiesRequest>().LastOrDefault();
            ApplySpreadsheetProperties(replaceProperties, requestedProperties);
            payload.Requests.Add(new GoogleSheetsApiRequestPayload {
                UpdateSpreadsheetProperties = new GoogleSheetsApiUpdateSpreadsheetPropertiesRequestPayload {
                    Properties = replaceProperties,
                    Fields = "title," + BuildSpreadsheetPropertyFields(requestedProperties),
                }
            });

            payload.Requests.Add(new GoogleSheetsApiRequestPayload {
                AddSheet = new GoogleSheetsApiAddSheetRequestPayload {
                    Properties = new GoogleSheetsApiSheetPropertiesPayload {
                        SheetId = keeperSheetId,
                        Title = keeperTitle,
                        Index = 0,
                    }
                }
            });

            foreach (var existingSheetId in existingSheets.Keys) {
                payload.Requests.Add(new GoogleSheetsApiRequestPayload {
                    DeleteSheet = new GoogleSheetsApiDeleteSheetRequestPayload {
                        SheetId = existingSheetId,
                    }
                });
            }

            foreach (var sheet in desiredSheets) {
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
                                    HideGridlines = sheet.HideGridlines,
                            }
                        }
                    }
                });
            }

            payload.Requests.Add(new GoogleSheetsApiRequestPayload {
                DeleteSheet = new GoogleSheetsApiDeleteSheetRequestPayload {
                    SheetId = keeperSheetId,
                }
            });

            return payload;
        }

        private static string BuildUniqueReplacementKeeperTitle(IEnumerable<string> reservedTitles) {
            const string baseTitle = "__OfficeIMO_Replacement_Keeper__";
            var reserved = new HashSet<string>(reservedTitles, StringComparer.OrdinalIgnoreCase);
            if (!reserved.Contains(baseTitle)) return baseTitle;

            for (int suffix = 2; ; suffix++) {
                string candidate = baseTitle + "_" + suffix.ToString(System.Globalization.CultureInfo.InvariantCulture);
                if (!reserved.Contains(candidate)) return candidate;
            }
        }

        private static void AppendUpdateCellsRequests(
            GoogleSheetsBatch batch,
            GoogleSheetsApiBatchUpdatePayload payload,
            int sheetId,
            GoogleSheetsUpdateCellsRequest request,
            IReadOnlyDictionary<string, int> sheetIds,
            string? spreadsheetId,
            bool includeCellValues) {
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

                    AddCellSegment(batch, payload, sheetId, rowGroup.Key, currentSegment, request.SheetName, sheetIds, spreadsheetId, includeCellValues);
                    currentSegment = new List<GoogleSheetsCellData> { cell };
                    expectedColumn = cell.ColumnIndex + 1;
                }

                if (currentSegment.Count > 0) {
                    AddCellSegment(batch, payload, sheetId, rowGroup.Key, currentSegment, request.SheetName, sheetIds, spreadsheetId, includeCellValues);
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
            string? spreadsheetId,
            bool includeCellValues) {
            var rowData = new GoogleSheetsApiRowDataPayload();
            bool includeFormat = false;
            bool includeNote = false;
            bool includeValidation = false;

            foreach (var cell in cells) {
                var apiCell = BuildCellData(batch, cell, sourceSheetName, sheetIds, spreadsheetId, includeCellValues, out var hasFormat, out var hasNote, out var hasValidation);
                rowData.Values.Add(apiCell);
                includeFormat |= hasFormat;
                includeNote |= hasNote;
                includeValidation |= hasValidation;
            }

            if (!includeCellValues && !includeFormat && !includeNote && !includeValidation) {
                return;
            }

            payload.Requests.Add(new GoogleSheetsApiRequestPayload {
                UpdateCells = new GoogleSheetsApiUpdateCellsRequestPayload {
                    Start = new GoogleSheetsApiGridCoordinatePayload {
                        SheetId = sheetId,
                        RowIndex = rowIndex,
                        ColumnIndex = cells[0].ColumnIndex,
                    },
                    Rows = new List<GoogleSheetsApiRowDataPayload> { rowData },
                    Fields = BuildUpdateCellsFields(includeCellValues, includeFormat, includeNote, includeValidation),
                }
            });
        }

        private static void AddValueRange(
            GoogleSheetsApiBatchUpdateValuesPayload payload,
            GoogleSheetsBatch batch,
            string sheetName,
            int rowIndex,
            IReadOnlyList<GoogleSheetsCellData> cells,
            IReadOnlyDictionary<string, int> sheetIds,
            string? spreadsheetId) {
            var range = new GoogleSheetsApiValueRangePayload {
                Range = $"'{sheetName.Replace("'", "''")}'!{ToColumnName(cells[0].ColumnIndex + 1)}{rowIndex + 1}:{ToColumnName(cells[cells.Count - 1].ColumnIndex + 1)}{rowIndex + 1}",
            };
            var values = new List<object?>();
            foreach (GoogleSheetsCellData cell in cells) {
                GoogleSheetsApiExtendedValuePayload? value = BuildExtendedValue(cell, batch, sheetName, sheetIds, spreadsheetId, out _);
                values.Add(ToValueInput(value));
            }
            range.Values.Add(values);
            payload.Data.Add(range);
        }

        private static object? ToValueInput(GoogleSheetsApiExtendedValuePayload? value) {
            if (value == null) return null;
            if (value.FormulaValue != null) return value.FormulaValue;
            if (value.StringValue != null) return value.StringValue.Length == 0 ? string.Empty : "'" + value.StringValue;
            if (value.NumberValue.HasValue) return value.NumberValue.Value;
            if (value.BoolValue.HasValue) return value.BoolValue.Value;
            return null;
        }

        private static string ToColumnName(int column) {
            var characters = new Stack<char>();
            while (column > 0) {
                column--;
                characters.Push((char)('A' + column % 26));
                column /= 26;
            }
            return new string(characters.ToArray());
        }

    }
}
