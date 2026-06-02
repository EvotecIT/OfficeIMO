using System.Globalization;

namespace OfficeIMO.Excel.GoogleSheets {
    internal static class GoogleSheetsApiPayloadBuilder {
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

        private static GoogleSheetsApiCellDataPayload BuildCellData(
            GoogleSheetsBatch batch,
            GoogleSheetsCellData cell,
            string sourceSheetName,
            IReadOnlyDictionary<string, int> sheetIds,
            string? spreadsheetId,
            out bool hasFormat,
            out bool hasNote,
            out bool hasValidation) {
            hasFormat = cell.Style != null;
            hasNote = false;
            hasValidation = cell.DataValidationRule != null;
            var valuePayload = BuildExtendedValue(cell, batch, sourceSheetName, sheetIds, spreadsheetId, out var hyperlinkNote);
            var payload = new GoogleSheetsApiCellDataPayload {
                UserEnteredValue = valuePayload,
                UserEnteredFormat = BuildCellFormat(cell.Style),
                DataValidationRule = BuildDataValidationRule(cell.DataValidationRule),
                Note = ComposeNote(cell.Comment, hyperlinkNote),
            };
            hasNote = !string.IsNullOrWhiteSpace(payload.Note);
            return payload;
        }

        private static GoogleSheetsApiExtendedValuePayload? BuildExtendedValue(
            GoogleSheetsCellData cell,
            GoogleSheetsBatch batch,
            string sourceSheetName,
            IReadOnlyDictionary<string, int> sheetIds,
            string? spreadsheetId,
            out string? note) {
            note = null;

            if (cell.Hyperlink != null && cell.Hyperlink.IsExternal && cell.Value.Kind is GoogleSheetsCellValueKind.String or GoogleSheetsCellValueKind.Blank) {
                var display = cell.Value.Kind == GoogleSheetsCellValueKind.String
                    ? Convert.ToString(cell.Value.Value, CultureInfo.InvariantCulture) ?? string.Empty
                    : cell.Hyperlink.Target;

                return new GoogleSheetsApiExtendedValuePayload {
                    FormulaValue = $"=HYPERLINK(\"{EscapeFormulaString(cell.Hyperlink.Target)}\",\"{EscapeFormulaString(display)}\")"
                };
            }

            if (cell.Hyperlink != null && !cell.Hyperlink.IsExternal) {
                if (TryBuildInternalHyperlinkFormula(cell, batch, sourceSheetName, sheetIds, spreadsheetId, out var hyperlinkFormula, out var hyperlinkNote)) {
                    note = hyperlinkNote;
                    AddReportNoticeOnce(
                        batch.Report,
                        OfficeIMO.GoogleWorkspace.TranslationSeverity.Info,
                        "InternalHyperlinks",
                        "Internal workbook hyperlinks are exported as Google Sheets hyperlinks to the target sheet while preserving the exact Excel target as a note.");

                    return new GoogleSheetsApiExtendedValuePayload {
                        FormulaValue = hyperlinkFormula,
                    };
                }

                note = "OfficeIMO internal link target: " + cell.Hyperlink.Target;
                AddReportNoticeOnce(
                    batch.Report,
                    OfficeIMO.GoogleWorkspace.TranslationSeverity.Info,
                    "InternalHyperlinks",
                    "Internal workbook hyperlinks are currently exported as Google Sheets cell notes.");
            }

            return cell.Value.Kind switch {
                GoogleSheetsCellValueKind.Blank => new GoogleSheetsApiExtendedValuePayload { StringValue = string.Empty },
                GoogleSheetsCellValueKind.String => new GoogleSheetsApiExtendedValuePayload { StringValue = Convert.ToString(cell.Value.Value, CultureInfo.InvariantCulture) ?? string.Empty },
                GoogleSheetsCellValueKind.Number => new GoogleSheetsApiExtendedValuePayload { NumberValue = Convert.ToDouble(cell.Value.Value, CultureInfo.InvariantCulture) },
                GoogleSheetsCellValueKind.Boolean => new GoogleSheetsApiExtendedValuePayload { BoolValue = Convert.ToBoolean(cell.Value.Value, CultureInfo.InvariantCulture) },
                GoogleSheetsCellValueKind.DateTime => new GoogleSheetsApiExtendedValuePayload { NumberValue = ConvertToSerialDate(cell.Value.Value) },
                GoogleSheetsCellValueKind.Formula => new GoogleSheetsApiExtendedValuePayload { FormulaValue = Convert.ToString(cell.Value.Value, CultureInfo.InvariantCulture) ?? "=" },
                _ => new GoogleSheetsApiExtendedValuePayload { StringValue = Convert.ToString(cell.Value.Value, CultureInfo.InvariantCulture) ?? string.Empty },
            };
        }

        private static bool TryBuildInternalHyperlinkFormula(
            GoogleSheetsCellData cell,
            GoogleSheetsBatch batch,
            string sourceSheetName,
            IReadOnlyDictionary<string, int> sheetIds,
            string? spreadsheetId,
            out string formula,
            out string note) {
            formula = string.Empty;
            note = string.Empty;

            if (cell.Hyperlink == null || cell.Hyperlink.IsExternal || string.IsNullOrWhiteSpace(spreadsheetId)) {
                return false;
            }

            string? targetSheetName;
            string? targetRangeText;
            bool resolvedFromNamedRange;

            if (!TryResolveInternalHyperlinkTarget(batch, sourceSheetName, cell.Hyperlink.Target, out targetSheetName, out targetRangeText, out resolvedFromNamedRange)) {
                return false;
            }

            if (string.IsNullOrWhiteSpace(targetSheetName) || !sheetIds.TryGetValue(targetSheetName!, out var targetSheetId)) {
                return false;
            }

            var display = cell.Value.Kind == GoogleSheetsCellValueKind.String
                ? Convert.ToString(cell.Value.Value, CultureInfo.InvariantCulture) ?? string.Empty
                : cell.Hyperlink.Target;
            var hyperlinkTarget = $"https://docs.google.com/spreadsheets/d/{spreadsheetId}/edit#gid={targetSheetId}";

            formula = $"=HYPERLINK(\"{EscapeFormulaString(hyperlinkTarget)}\",\"{EscapeFormulaString(display)}\")";
            if (string.IsNullOrWhiteSpace(targetRangeText)) {
                note = $"OfficeIMO internal link target: {cell.Hyperlink.Target}";
            } else if (resolvedFromNamedRange) {
                note = $"OfficeIMO internal link target: {cell.Hyperlink.Target} -> {targetSheetName}!{targetRangeText}";
            } else {
                note = $"OfficeIMO internal link target: {targetSheetName}!{targetRangeText}";
            }
            return true;
        }

        private static string? ComposeNote(GoogleSheetsComment? comment, string? hyperlinkNote) {
            string? commentNote = null;
            if (comment != null && !string.IsNullOrWhiteSpace(comment.Text)) {
                commentNote = string.IsNullOrWhiteSpace(comment.Author)
                    ? comment.Text
                    : comment.Author + ": " + comment.Text;
            }

            if (string.IsNullOrWhiteSpace(commentNote)) {
                return string.IsNullOrWhiteSpace(hyperlinkNote) ? null : hyperlinkNote;
            }

            if (string.IsNullOrWhiteSpace(hyperlinkNote)) {
                return commentNote;
            }

            return commentNote + Environment.NewLine + Environment.NewLine + hyperlinkNote;
        }

        private static bool TryResolveInternalHyperlinkTarget(
            GoogleSheetsBatch batch,
            string sourceSheetName,
            string hyperlinkTarget,
            out string? targetSheetName,
            out string? targetRangeText,
            out bool resolvedFromNamedRange) {
            targetSheetName = null;
            targetRangeText = null;
            resolvedFromNamedRange = false;

            if (TrySplitSheetQualifiedRange(hyperlinkTarget, out var explicitSheetName, out var explicitRangeText)) {
                targetSheetName = explicitSheetName;
                targetRangeText = explicitRangeText;
                return !string.IsNullOrWhiteSpace(targetSheetName);
            }

            var namedRange = ResolveNamedRangeTarget(batch, sourceSheetName, hyperlinkTarget);
            if (namedRange == null) {
                return false;
            }

            resolvedFromNamedRange = true;
            if (TrySplitSheetQualifiedRange(namedRange.A1Range, out var namedRangeSheetName, out var namedRangeRangeText)) {
                targetSheetName = namedRangeSheetName;
                targetRangeText = namedRangeRangeText;
                return !string.IsNullOrWhiteSpace(targetSheetName);
            }

            targetSheetName = namedRange.SheetName;
            targetRangeText = namedRange.A1Range.Replace("$", string.Empty);
            return !string.IsNullOrWhiteSpace(targetSheetName);
        }

        private static GoogleSheetsAddNamedRangeRequest? ResolveNamedRangeTarget(
            GoogleSheetsBatch batch,
            string sourceSheetName,
            string hyperlinkTarget) {
            var namedRanges = batch.Requests
                .OfType<GoogleSheetsAddNamedRangeRequest>()
                .Where(request => string.Equals(request.Name, hyperlinkTarget, StringComparison.OrdinalIgnoreCase))
                .ToList();

            if (namedRanges.Count == 0) {
                return null;
            }

            return namedRanges.FirstOrDefault(request => string.Equals(request.SheetName, sourceSheetName, StringComparison.OrdinalIgnoreCase))
                ?? namedRanges.FirstOrDefault(request => string.IsNullOrWhiteSpace(request.SheetName))
                ?? namedRanges[0];
        }

        private static void AddReportNoticeOnce(
            OfficeIMO.GoogleWorkspace.TranslationReport report,
            OfficeIMO.GoogleWorkspace.TranslationSeverity severity,
            string feature,
            string message) {
            if (!report.Notices.Any(notice =>
                    notice.Severity == severity
                    && string.Equals(notice.Feature, feature, StringComparison.Ordinal)
                    && string.Equals(notice.Message, message, StringComparison.Ordinal))) {
                report.Add(severity, feature, message);
            }
        }

        private static double ConvertToSerialDate(object? value) {
            if (value is DateTimeOffset dto) {
                return dto.UtcDateTime.ToOADate();
            }

            if (value is DateTime dateTime) {
                return dateTime.ToOADate();
            }

            return 0;
        }

        private static GoogleSheetsApiCellFormatPayload? BuildCellFormat(GoogleSheetsCellStyle? style) {
            if (style == null) {
                return null;
            }

            var payload = new GoogleSheetsApiCellFormatPayload {
                NumberFormat = BuildNumberFormat(style),
                BackgroundColor = BuildColor(style.FillColorArgb),
                Borders = BuildBorders(style.Borders),
                HorizontalAlignment = NormalizeHorizontalAlignment(style.HorizontalAlignment),
                VerticalAlignment = NormalizeVerticalAlignment(style.VerticalAlignment),
                WrapStrategy = style.WrapText ? "WRAP" : null,
            };

            if (style.Bold || style.Italic || style.Underline || !string.IsNullOrWhiteSpace(style.FontColorArgb)) {
                payload.TextFormat = new GoogleSheetsApiTextFormatPayload {
                    Bold = style.Bold ? true : (bool?)null,
                    Italic = style.Italic ? true : (bool?)null,
                    Underline = style.Underline ? true : (bool?)null,
                    ForegroundColor = BuildColor(style.FontColorArgb),
                };
            }

            return payload;
        }

        private static GoogleSheetsApiBordersPayload? BuildBorders(GoogleSheetsCellBorders? borders) {
            if (borders == null) {
                return null;
            }

            var payload = new GoogleSheetsApiBordersPayload {
                Left = BuildBorderSide(borders.Left),
                Right = BuildBorderSide(borders.Right),
                Top = BuildBorderSide(borders.Top),
                Bottom = BuildBorderSide(borders.Bottom),
            };

            if (payload.Left == null && payload.Right == null && payload.Top == null && payload.Bottom == null) {
                return null;
            }

            return payload;
        }

        private static GoogleSheetsApiBorderPayload? BuildBorderSide(GoogleSheetsBorderSide? side) {
            if (side == null) {
                return null;
            }

            var style = NormalizeBorderStyle(side.Style);
            var color = BuildColor(side.ColorArgb);
            if (style == null && color == null) {
                return null;
            }

            return new GoogleSheetsApiBorderPayload {
                Style = style ?? "SOLID",
                Color = color,
            };
        }

        private static GoogleSheetsApiNumberFormatPayload? BuildNumberFormat(GoogleSheetsCellStyle style) {
            if (string.IsNullOrWhiteSpace(style.NumberFormatCode) && !style.IsDateLike) {
                return null;
            }

            return new GoogleSheetsApiNumberFormatPayload {
                Type = ResolveNumberFormatType(style),
                Pattern = style.NumberFormatCode,
            };
        }

        private static string ResolveNumberFormatType(GoogleSheetsCellStyle style) {
            if (style.IsDateLike) {
                return "DATE_TIME";
            }

            var pattern = style.NumberFormatCode ?? string.Empty;
            if (pattern.IndexOf('%') >= 0) {
                return "PERCENT";
            }

            if (pattern.IndexOf('$') >= 0 || pattern.IndexOf("z", StringComparison.OrdinalIgnoreCase) >= 0) {
                return "CURRENCY";
            }

            return "NUMBER";
        }

        private static string? NormalizeBorderStyle(string? value) {
            if (string.IsNullOrWhiteSpace(value)) {
                return null;
            }

            var normalized = value == null ? string.Empty : value.Trim().ToLowerInvariant();
            return normalized switch {
                "thin" => "SOLID",
                "medium" => "SOLID_MEDIUM",
                "thick" => "SOLID_THICK",
                "double" => "DOUBLE",
                "dashed" => "DASHED",
                "mediumdashed" => "DASHED",
                "dashdot" => "DASHED",
                "mediumdashdot" => "DASHED",
                "dashdotdot" => "DOTTED",
                "mediumdashdotdot" => "DOTTED",
                "dotted" => "DOTTED",
                "hair" => "DOTTED",
                "slantdashdot" => "DASHED",
                _ => "SOLID",
            };
        }

        private static GoogleSheetsApiColorPayload? BuildColor(string? argb) {
            if (string.IsNullOrWhiteSpace(argb) || argb!.Length != 8) {
                return null;
            }

            var red = int.Parse(argb.Substring(2, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture) / 255d;
            var green = int.Parse(argb.Substring(4, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture) / 255d;
            var blue = int.Parse(argb.Substring(6, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture) / 255d;

            return new GoogleSheetsApiColorPayload {
                Red = red,
                Green = green,
                Blue = blue,
            };
        }

        private static GoogleSheetsApiTableRowsPropertiesPayload? BuildTableRowsProperties(GoogleSheetsAddTableRequest table) {
            if (table == null) throw new ArgumentNullException(nameof(table));
            var headerColorStyle = BuildColorStyle(table.HeaderColorArgb);
            var firstBandColorStyle = BuildColorStyle(table.FirstBandColorArgb);
            var secondBandColorStyle = BuildColorStyle(table.SecondBandColorArgb);
            var footerColorStyle = BuildColorStyle(table.FooterColorArgb);

            if (headerColorStyle == null
                && firstBandColorStyle == null
                && secondBandColorStyle == null
                && footerColorStyle == null) {
                return null;
            }

            return new GoogleSheetsApiTableRowsPropertiesPayload {
                HeaderColorStyle = headerColorStyle,
                FirstBandColorStyle = firstBandColorStyle,
                SecondBandColorStyle = secondBandColorStyle,
                FooterColorStyle = footerColorStyle,
            };
        }

        private static GoogleSheetsApiColorStylePayload? BuildColorStyle(string? argb) {
            var color = BuildColor(argb);
            if (color == null) {
                return null;
            }

            return new GoogleSheetsApiColorStylePayload {
                RgbColor = color,
            };
        }

        private static GoogleSheetsApiDataValidationRulePayload? BuildDataValidationRule(GoogleSheetsDataValidationRule? rule) {
            if (rule == null || string.IsNullOrWhiteSpace(rule.ConditionType)) {
                return null;
            }

            return new GoogleSheetsApiDataValidationRulePayload {
                Condition = new GoogleSheetsApiBooleanConditionPayload {
                    Type = rule.ConditionType,
                    Values = rule.Values.Count == 0
                        ? null
                        : rule.Values.Select(value => new GoogleSheetsApiConditionValuePayload {
                            UserEnteredValue = value,
                        }).ToList(),
                },
                Strict = rule.Strict,
                ShowCustomUi = rule.ShowCustomUi,
            };
        }

        private static string? NormalizeHorizontalAlignment(string? value) {
            return value switch {
                null => null,
                "" => null,
                "left" => "LEFT",
                "center" => "CENTER",
                "right" => "RIGHT",
                "fill" => "LEFT",
                "justify" => "CENTER",
                _ => value.ToUpperInvariant(),
            };
        }

        private static string? NormalizeVerticalAlignment(string? value) {
            return value switch {
                null => null,
                "" => null,
                "top" => "TOP",
                "center" => "MIDDLE",
                "bottom" => "BOTTOM",
                _ => value.ToUpperInvariant(),
            };
        }

        private static bool TryBuildNamedRangePayload(
            IReadOnlyDictionary<string, int> sheetIds,
            GoogleSheetsAddNamedRangeRequest request,
            out GoogleSheetsApiNamedRangePayload payload) {
            payload = new GoogleSheetsApiNamedRangePayload {
                Name = request.Name,
                Range = new GoogleSheetsApiGridRangePayload(),
            };

            string? sheetName = request.SheetName;
            string rangeText = request.A1Range;

            if (TrySplitSheetQualifiedRange(rangeText, out var explicitSheet, out var unqualifiedRange)) {
                sheetName = explicitSheet;
                rangeText = unqualifiedRange;
            }

            if (string.IsNullOrWhiteSpace(sheetName)) {
                return false;
            }

            if (!sheetIds.TryGetValue(sheetName!, out var sheetId)) {
                return false;
            }

            int rowStart;
            int columnStart;
            int rowEnd;
            int columnEnd;
            if (!A1.TryParseRange(rangeText, out rowStart, out columnStart, out rowEnd, out columnEnd)) {
                var (row, column) = A1.ParseCellRef(rangeText);
                if (row <= 0 || column <= 0) {
                    return false;
                }

                rowStart = rowEnd = row;
                columnStart = columnEnd = column;
            }

            payload.Range = new GoogleSheetsApiGridRangePayload {
                SheetId = sheetId,
                StartRowIndex = rowStart - 1,
                EndRowIndex = rowEnd,
                StartColumnIndex = columnStart - 1,
                EndColumnIndex = columnEnd,
            };
            return true;
        }

        private static bool TryBuildGridRange(
            IReadOnlyDictionary<string, int> sheetIds,
            string sheetName,
            string rangeText,
            out GoogleSheetsApiGridRangePayload payload) {
            payload = new GoogleSheetsApiGridRangePayload();
            if (string.IsNullOrWhiteSpace(sheetName) || string.IsNullOrWhiteSpace(rangeText)) {
                return false;
            }

            if (!sheetIds.TryGetValue(sheetName, out var sheetId)) {
                return false;
            }

            string unqualifiedRange = rangeText;
            if (TrySplitSheetQualifiedRange(rangeText, out _, out var explicitRange)) {
                unqualifiedRange = explicitRange;
            }

            int rowStart;
            int columnStart;
            int rowEnd;
            int columnEnd;
            if (!A1.TryParseRange(unqualifiedRange, out rowStart, out columnStart, out rowEnd, out columnEnd)) {
                var (row, column) = A1.ParseCellRef(unqualifiedRange);
                if (row <= 0 || column <= 0) {
                    return false;
                }

                rowStart = rowEnd = row;
                columnStart = columnEnd = column;
            }

            payload = new GoogleSheetsApiGridRangePayload {
                SheetId = sheetId,
                StartRowIndex = rowStart - 1,
                EndRowIndex = rowEnd,
                StartColumnIndex = columnStart - 1,
                EndColumnIndex = columnEnd,
            };
            return true;
        }

        private static Dictionary<string, GoogleSheetsApiFilterCriteriaPayload> BuildFilterCriteriaMap(
            IReadOnlyList<GoogleSheetsFilterColumnCriteria> criteria) {
            var map = new Dictionary<string, GoogleSheetsApiFilterCriteriaPayload>();
            foreach (var criterion in criteria) {
                if ((criterion.HiddenValues == null || criterion.HiddenValues.Count == 0) && criterion.Condition == null) {
                    continue;
                }

                map[criterion.ColumnId.ToString(CultureInfo.InvariantCulture)] = new GoogleSheetsApiFilterCriteriaPayload {
                    HiddenValues = criterion.HiddenValues == null || criterion.HiddenValues.Count == 0 ? null : criterion.HiddenValues.ToList(),
                    Condition = criterion.Condition == null ? null : new GoogleSheetsApiBooleanConditionPayload {
                        Type = criterion.Condition.Type,
                        Values = criterion.Condition.Values.Select(value => new GoogleSheetsApiConditionValuePayload {
                            UserEnteredValue = value,
                        }).ToList(),
                    },
                };
            }

            return map;
        }

        private static bool TrySplitSheetQualifiedRange(string value, out string? sheetName, out string unqualifiedRange) {
            sheetName = null;
            unqualifiedRange = value;
            if (string.IsNullOrWhiteSpace(value)) {
                return false;
            }

            var bangIndex = value.LastIndexOf('!');
            if (bangIndex <= 0 || bangIndex >= value.Length - 1) {
                return false;
            }

            var sheetPart = value.Substring(0, bangIndex).Trim();
            var rangePart = value.Substring(bangIndex + 1).Trim();
            if (sheetPart.Length >= 2 && sheetPart[0] == '\'' && sheetPart[sheetPart.Length - 1] == '\'') {
                sheetPart = sheetPart.Substring(1, sheetPart.Length - 2).Replace("''", "'");
            }

            sheetName = sheetPart;
            unqualifiedRange = rangePart.Replace("$", string.Empty);
            return true;
        }

        private static string BuildUpdateCellsFields(bool includeFormat, bool includeNote, bool includeValidation) {
            var fields = new List<string> { "userEnteredValue" };
            if (includeFormat) {
                fields.Add("userEnteredFormat");
            }
            if (includeValidation) {
                fields.Add("dataValidationRule");
            }
            if (includeNote) {
                fields.Add("note");
            }
            return string.Join(",", fields);
        }

        private static string BuildDimensionFields(GoogleSheetsUpdateDimensionPropertiesRequest request) {
            var fields = new List<string>();
            if (request.PixelSize.HasValue) {
                fields.Add("pixelSize");
            }
            if (request.Hidden) {
                fields.Add("hiddenByUser");
            }
            return fields.Count > 0 ? string.Join(",", fields) : "hiddenByUser";
        }

        private static string EscapeFormulaString(string value) {
            return (value ?? string.Empty).Replace("\"", "\"\"");
        }

    }
}
