namespace OfficeIMO.Excel.GoogleSheets {
    internal static partial class GoogleSheetsApiPayloadBuilder {
        private static void ApplySpreadsheetProperties(
            GoogleSheetsApiSpreadsheetPropertiesPayload payload,
            GoogleSheetsUpdateSpreadsheetPropertiesRequest? request) {
            if (request == null) return;
            payload.Locale = string.IsNullOrWhiteSpace(request.Locale) ? null : request.Locale;
            payload.TimeZone = string.IsNullOrWhiteSpace(request.TimeZone) ? null : request.TimeZone;
            payload.AutoRecalc = request.RecalculationInterval switch {
                GoogleSheetsRecalculationInterval.Minute => "MINUTE",
                GoogleSheetsRecalculationInterval.Hour => "HOUR",
                _ => "ON_CHANGE",
            };
        }

        private static string BuildSpreadsheetPropertyFields(GoogleSheetsUpdateSpreadsheetPropertiesRequest? request) {
            if (request == null) return "autoRecalc";
            var fields = new List<string> { "autoRecalc" };
            if (!string.IsNullOrWhiteSpace(request.Locale)) fields.Add("locale");
            if (!string.IsNullOrWhiteSpace(request.TimeZone)) fields.Add("timeZone");
            return string.Join(",", fields);
        }

        private static List<GoogleSheetsApiGridRangePayload>? BuildUnprotectedRanges(
            IReadOnlyDictionary<string, int> sheetIds,
            GoogleSheetsAddProtectedRangeRequest request) {
            var ranges = new List<GoogleSheetsApiGridRangePayload>();
            foreach (string a1 in request.UnprotectedA1Ranges) {
                if (TryBuildGridRange(sheetIds, request.SheetName, a1, out var range)) ranges.Add(range);
            }
            return ranges.Count == 0 ? null : ranges;
        }

        private static bool TryBuildChartPayload(
            IReadOnlyDictionary<string, int> sheetIds,
            GoogleSheetsAddChartRequest request,
            out GoogleSheetsApiAddChartRequestPayload payload) {
            payload = new GoogleSheetsApiAddChartRequestPayload();
            if (!sheetIds.TryGetValue(request.SheetName, out int targetSheetId)
                || !sheetIds.TryGetValue(request.DataSheetName, out int dataSheetId)
                || request.DataRowCount < 2 || request.SeriesCount < 1) return false;

            var chart = new GoogleSheetsApiBasicChartSpecPayload {
                ChartType = request.ChartType,
            };
            chart.Domains.Add(new GoogleSheetsApiBasicChartDomainPayload {
                Domain = BuildChartData(dataSheetId, request.DataStartRowIndex, request.DataRowCount, 0),
            });
            for (int i = 0; i < request.SeriesCount; i++) {
                chart.Series.Add(new GoogleSheetsApiBasicChartSeriesPayload {
                    Series = BuildChartData(dataSheetId, request.DataStartRowIndex, request.DataRowCount, i + 1),
                });
            }

            payload.Chart = new GoogleSheetsApiEmbeddedChartPayload {
                Spec = new GoogleSheetsApiChartSpecPayload { Title = request.Title, BasicChart = chart },
                Position = new GoogleSheetsApiEmbeddedObjectPositionPayload {
                    OverlayPosition = new GoogleSheetsApiOverlayPositionPayload {
                        AnchorCell = new GoogleSheetsApiGridCoordinatePayload {
                            SheetId = targetSheetId,
                            RowIndex = request.AnchorRowIndex,
                            ColumnIndex = request.AnchorColumnIndex,
                        }
                    }
                }
            };
            return true;
        }

        private static GoogleSheetsApiChartDataPayload BuildChartData(int sheetId, int startRow, int rowCount, int column) {
            var result = new GoogleSheetsApiChartDataPayload();
            result.SourceRange.Sources.Add(new GoogleSheetsApiGridRangePayload {
                SheetId = sheetId,
                StartRowIndex = startRow,
                EndRowIndex = startRow + rowCount,
                StartColumnIndex = column,
                EndColumnIndex = column + 1,
            });
            return result;
        }

        private static bool TryBuildPivotTablePayload(
            IReadOnlyDictionary<string, int> sheetIds,
            GoogleSheetsAddPivotTableRequest request,
            out GoogleSheetsApiRequestPayload payload) {
            payload = new GoogleSheetsApiRequestPayload();
            if (!sheetIds.TryGetValue(request.SheetName, out int destinationSheetId)
                || !TryBuildGridRange(sheetIds, request.SourceSheetName, request.SourceA1Range, out var sourceRange)) return false;

            var pivot = new GoogleSheetsApiPivotTablePayload { Source = sourceRange };
            foreach (GoogleSheetsPivotGroup row in request.Rows) {
                pivot.Rows.Add(new GoogleSheetsApiPivotGroupPayload { SourceColumnOffset = row.SourceColumnOffset, ShowTotals = row.ShowTotals, SortOrder = row.SortOrder });
            }
            foreach (GoogleSheetsPivotGroup column in request.Columns) {
                pivot.Columns.Add(new GoogleSheetsApiPivotGroupPayload { SourceColumnOffset = column.SourceColumnOffset, ShowTotals = column.ShowTotals, SortOrder = column.SortOrder });
            }
            foreach (GoogleSheetsPivotValue value in request.Values) {
                pivot.Values.Add(new GoogleSheetsApiPivotValuePayload { SourceColumnOffset = value.SourceColumnOffset, SummarizeFunction = value.SummarizeFunction, Name = value.Name });
            }

            var rowData = new GoogleSheetsApiRowDataPayload();
            rowData.Values.Add(new GoogleSheetsApiCellDataPayload { PivotTable = pivot });
            payload.UpdateCells = new GoogleSheetsApiUpdateCellsRequestPayload {
                Start = new GoogleSheetsApiGridCoordinatePayload {
                    SheetId = destinationSheetId,
                    RowIndex = request.DestinationRowIndex,
                    ColumnIndex = request.DestinationColumnIndex,
                },
                Rows = new List<GoogleSheetsApiRowDataPayload> { rowData },
                Fields = "pivotTable",
            };
            return true;
        }
    }
}
