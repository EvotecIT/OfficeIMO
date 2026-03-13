using OfficeIMO.Excel;
using OfficeIMO.Excel.GoogleSheets;
using OfficeIMO.GoogleWorkspace;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_ExcelInspectionSnapshot_ExposesOfficeIMOWorkbookModel() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelInspectionSnapshot.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var summary = document.AddWorkSheet("Summary");
                    var hidden = document.AddWorkSheet("Hidden");

                    summary.CellValue(1, 1, "Name");
                    summary.CellValue(2, 1, "Alpha");
                    summary.CellValue(2, 3, "Accent");
                    summary.CellFormula(2, 2, "SUM(1,2)");
                    summary.SetHyperlink(1, 1, "https://example.org", display: "Name");
                    summary.SetComment(2, 3, "Accent comment", author: "Tester", initials: "TT");
                    summary.FormatCell(2, 1, "0.00%");
                    summary.CellBackground(2, 1, "#00FF00");
                    summary.CellBold(2, 1, true);
                    summary.CellAlign(2, 1, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center);
                    summary.CellFontColor(2, 3, "#112233");
                    summary.SetColumnWidth(1, 20);
                    summary.CellValue(3, 1, "First line\nSecond line");
                    summary.WrapCells(3, 3, 1, 20);
                    summary.AutoFitRow(3);
                    summary.CellValue(1, 4, "Status");
                    summary.CellValue(1, 5, "Region");
                    summary.CellValue(1, 6, "Score");
                    summary.CellValue(1, 7, "Budget");
                    summary.CellValue(2, 4, "Open");
                    summary.CellValue(2, 5, "North");
                    summary.CellValue(2, 6, 10d);
                    summary.CellValue(2, 7, 8d);
                    summary.CellValue(3, 4, "Closed");
                    summary.CellValue(3, 5, "South");
                    summary.CellValue(3, 6, 20d);
                    summary.CellValue(3, 7, 18d);
                    summary.CellValue(4, 4, "Open");
                    summary.CellValue(4, 5, "East");
                    summary.CellValue(4, 6, 30d);
                    summary.CellValue(4, 7, 28d);
                    summary.AddTable("A1:B2", hasHeader: true, name: "SummaryTable", style: OfficeIMO.Excel.TableStyle.TableStyleMedium2, includeAutoFilter: true);
                    summary.SetTableTotals("A1:B2", new Dictionary<string, DocumentFormat.OpenXml.Spreadsheet.TotalsRowFunctionValues> {
                        ["Name"] = DocumentFormat.OpenXml.Spreadsheet.TotalsRowFunctionValues.Count,
                        ["Column2"] = DocumentFormat.OpenXml.Spreadsheet.TotalsRowFunctionValues.Sum,
                    });
                    summary.AddAutoFilter("D1:G4", new Dictionary<uint, IEnumerable<string>> {
                        { 0, new[] { "Open" } }
                    });
                    summary.AutoFilterByHeaderContains("Region", "or");
                    summary.AutoFilterByHeaderGreaterThanOrEqual("Score", 15d);
                    summary.AutoFilterByHeaderBetween("Budget", 10d, 20d);
                    summary.Freeze(topRows: 1, leftCols: 1);
                    summary.Protect(new ExcelSheetProtectionOptions {
                        AllowSelectLockedCells = false,
                        AllowSelectUnlockedCells = false,
                        AllowSort = true,
                        AllowAutoFilter = true,
                        AllowInsertRows = true,
                    });
                    hidden.SetHidden(true);
                    document.SetNamedRange("GlobalData", "'Summary'!A1:B2", save: false);
                    document.Save();
                }

                ApplyBorderToCell(filePath, "Summary", "A2");
                ApplySheetDisplaySettings(filePath, "Summary", "FF336699", rightToLeft: true);

                using var reloadedDocument = ExcelDocument.Load(filePath);
                var snapshot = reloadedDocument.CreateInspectionSnapshot();

                Assert.Equal(2, snapshot.Worksheets.Count);

                var summarySheet = Assert.Single(snapshot.Worksheets, w => w.Name == "Summary");
                Assert.Equal(0, summarySheet.Index);
                Assert.False(summarySheet.Hidden);
                Assert.True(summarySheet.RightToLeft);
                Assert.Equal("FF336699", summarySheet.TabColorArgb);
                Assert.Equal(1, summarySheet.FrozenRowCount);
                Assert.Equal(1, summarySheet.FrozenColumnCount);
                Assert.NotNull(summarySheet.Protection);
                Assert.False(summarySheet.Protection!.AllowSelectLockedCells);
                Assert.False(summarySheet.Protection.AllowSelectUnlockedCells);
                Assert.True(summarySheet.Protection.AllowSort);
                Assert.True(summarySheet.Protection.AllowAutoFilter);
                Assert.True(summarySheet.Protection.AllowInsertRows);
                Assert.NotNull(summarySheet.AutoFilter);
                Assert.Equal("D1:G4", summarySheet.AutoFilter!.A1Range);
                var worksheetFilterColumn = Assert.Single(summarySheet.AutoFilter.Columns, column => column.ColumnId == 0);
                Assert.Equal(new[] { "Open" }, worksheetFilterColumn.Values);
                var table = Assert.Single(summarySheet.Tables);
                Assert.Equal("SummaryTable", table.Name);
                Assert.Equal("A1:B2", table.A1Range);
                Assert.Equal("TableStyleMedium2", table.StyleName);
                Assert.True(table.HasHeaderRow);
                Assert.True(table.TotalsRowShown);
                Assert.NotNull(table.AutoFilter);
                Assert.Equal("A1:B2", table.AutoFilter!.A1Range);
                Assert.Equal(new[] { "Name", "Column2" }, table.Columns.Select(c => c.Name).ToArray());
                Assert.Equal("count", table.Columns[0].TotalsRowFunction);
                Assert.Equal("sum", table.Columns[1].TotalsRowFunction);
                var containsFilter = Assert.Single(summarySheet.AutoFilter.Columns, column => column.ColumnId == 1);
                Assert.NotNull(containsFilter.CustomFilters);
                Assert.False(containsFilter.CustomFilters!.MatchAll);
                var containsCondition = Assert.Single(containsFilter.CustomFilters.Conditions);
                Assert.Equal("equal", containsCondition.Operator, StringComparer.OrdinalIgnoreCase);
                Assert.Equal("*or*", containsCondition.Value);
                var scoreFilter = Assert.Single(summarySheet.AutoFilter.Columns, column => column.ColumnId == 2);
                Assert.NotNull(scoreFilter.CustomFilters);
                var scoreCondition = Assert.Single(scoreFilter.CustomFilters!.Conditions);
                Assert.Equal("greaterThanOrEqual", scoreCondition.Operator, StringComparer.OrdinalIgnoreCase);
                Assert.Equal("15", scoreCondition.Value);
                var budgetFilter = Assert.Single(summarySheet.AutoFilter.Columns, column => column.ColumnId == 3);
                Assert.NotNull(budgetFilter.CustomFilters);
                Assert.True(budgetFilter.CustomFilters!.MatchAll);
                Assert.Equal(2, budgetFilter.CustomFilters.Conditions.Count);
                Assert.Contains(budgetFilter.CustomFilters.Conditions, condition => string.Equals(condition.Operator, "greaterThanOrEqual", StringComparison.OrdinalIgnoreCase) && condition.Value == "10");
                Assert.Contains(budgetFilter.CustomFilters.Conditions, condition => string.Equals(condition.Operator, "lessThanOrEqual", StringComparison.OrdinalIgnoreCase) && condition.Value == "20");
                Assert.Contains(summarySheet.Cells, c => c.Row == 2 && c.Column == 2 && c.Formula == "SUM(1,2)");
                var linkedCell = Assert.Single(summarySheet.Cells, c => c.Row == 1 && c.Column == 1);
                Assert.NotNull(linkedCell.Hyperlink);
                Assert.True(linkedCell.Hyperlink!.IsExternal);
                Assert.Equal("https://example.org", linkedCell.Hyperlink.Target);

                var styledCell = Assert.Single(summarySheet.Cells, c => c.Row == 2 && c.Column == 1);
                Assert.NotNull(styledCell.Style);
                var style = styledCell.Style!;
                Assert.Equal("0.00%", style.NumberFormatCode);
                Assert.True(style.Bold);
                Assert.Equal("FF00FF00", style.FillColorArgb);
                Assert.Equal("center", style.HorizontalAlignment);
                Assert.NotNull(style.Border);
                Assert.Equal("medium", style.Border!.Left!.Style);
                Assert.Equal("FFFF0000", style.Border.Left.ColorArgb);
                Assert.Equal("dashed", style.Border.Top!.Style);
                Assert.Equal("FF0000FF", style.Border.Top.ColorArgb);

                var fontColorCell = Assert.Single(summarySheet.Cells, c => c.Row == 2 && c.Column == 3);
                Assert.NotNull(fontColorCell.Style);
                Assert.Equal("FF112233", fontColorCell.Style!.FontColorArgb);
                Assert.NotNull(fontColorCell.Comment);
                Assert.Equal("Tester (TT)", fontColorCell.Comment!.Author);
                Assert.Equal("Accent comment", fontColorCell.Comment.Text);

                var column = Assert.Single(summarySheet.Columns, c => c.StartIndex == 1 && c.EndIndex == 1);
                Assert.Equal(20, column.Width);

                var row = Assert.Single(summarySheet.Rows, r => r.Index == 3);
                Assert.True(row.Height > 15);

                var hiddenSheet = Assert.Single(snapshot.Worksheets, w => w.Name == "Hidden");
                Assert.True(hiddenSheet.Hidden);

                Assert.Contains(snapshot.NamedRanges, n => n.Name == "GlobalData" && n.ReferenceA1 == "'Summary'!$A$1:$B$2" && !n.IsBuiltIn);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Test_GoogleSheetsBatchCompiler_EmitsWorkbookStructureAndCellRequests() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleSheetsBatchCompiler.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var summary = document.AddWorkSheet("Summary");
                    var hidden = document.AddWorkSheet("Hidden");

                    summary.CellValue(1, 1, "Name");
                    summary.CellValue(1, 2, "Count");
                    summary.CellValue(2, 1, "Alpha");
                    summary.CellValue(2, 2, 12);
                    summary.CellValue(2, 3, true);
                    summary.CellValue(2, 4, new DateTime(2024, 12, 24, 10, 30, 0, DateTimeKind.Utc));
                    summary.CellValue(2, 6, "Accent");
                    summary.CellFormula(2, 5, "SUM(B2:B2)");
                    summary.SetHyperlink(2, 1, "https://alpha.example/", display: "Alpha");
                    summary.SetComment(2, 6, "Accent comment", author: "Tester", initials: "TT");
                    summary.FormatCell(2, 2, "0.00%");
                    summary.CellBackground(2, 2, "#00FF00");
                    summary.CellBold(2, 2, true);
                    summary.CellAlign(2, 2, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center);
                    summary.CellFontColor(2, 6, "#112233");
                    summary.SetColumnWidth(2, 20);
                    summary.CellValue(3, 2, "Wrapped\nRow");
                    summary.WrapCells(3, 3, 2, 20);
                    summary.AutoFitRow(3);
                    summary.CellValue(1, 7, "Status");
                    summary.CellValue(1, 8, "Region");
                    summary.CellValue(1, 9, "Score");
                    summary.CellValue(1, 10, "Budget");
                    summary.CellValue(2, 7, "Open");
                    summary.CellValue(2, 8, "North");
                    summary.CellValue(2, 9, 10d);
                    summary.CellValue(2, 10, 8d);
                    summary.CellValue(3, 7, "Closed");
                    summary.CellValue(3, 8, "South");
                    summary.CellValue(3, 9, 20d);
                    summary.CellValue(3, 10, 18d);
                    summary.CellValue(4, 7, "Open");
                    summary.CellValue(4, 8, "East");
                    summary.CellValue(4, 9, 30d);
                    summary.CellValue(4, 10, 28d);
                    summary.AddTable("A1:B3", hasHeader: true, name: "SummaryTable", style: OfficeIMO.Excel.TableStyle.TableStyleMedium2, includeAutoFilter: true);
                    summary.SetTableTotals("A1:B3", new Dictionary<string, DocumentFormat.OpenXml.Spreadsheet.TotalsRowFunctionValues> {
                        ["Name"] = DocumentFormat.OpenXml.Spreadsheet.TotalsRowFunctionValues.Count,
                        ["Count"] = DocumentFormat.OpenXml.Spreadsheet.TotalsRowFunctionValues.Sum,
                    });
                    summary.CellBackground(3, 1, "#D9EAD3");
                    summary.CellBackground(3, 2, "#D9EAD3");
                    summary.AddAutoFilter("G1:J4", new Dictionary<uint, IEnumerable<string>> {
                        { 0, new[] { "Open" } }
                    });
                    summary.AutoFilterByHeaderContains("Region", "or");
                    summary.AutoFilterByHeaderGreaterThanOrEqual("Score", 15d);
                    summary.AutoFilterByHeaderBetween("Budget", 10d, 20d);
                    summary.Freeze(topRows: 1, leftCols: 1);
                    summary.Protect(new ExcelSheetProtectionOptions {
                        AllowSelectLockedCells = false,
                        AllowSelectUnlockedCells = false,
                        AllowSort = true,
                        AllowAutoFilter = true,
                        AllowInsertRows = true,
                    });
                    hidden.SetHidden(true);
                    summary.SetInternalLink(5, 1, hidden, "B5", display: "Go hidden");
                    summary.SetInternalLink(6, 1, "LocalData", display: "Go local");
                    summary.SetComment(6, 1, "Jump note", author: "Tester", initials: "TT");

                    document.SetNamedRange("GlobalData", "'Summary'!A1:B2", save: false);
                    summary.SetNamedRange("LocalData", "A2:B2", save: false);
                    document.Save();
                }

                ApplyBorderToCell(filePath, "Summary", "B2");
                ApplySheetDisplaySettings(filePath, "Summary", "FF336699", rightToLeft: true);

                using var reloadedDocument = ExcelDocument.Load(filePath);
                var batch = reloadedDocument.CreateGoogleSheetsBatch(new GoogleSheetsSaveOptions {
                    Title = "OfficeIMO Export"
                });

                Assert.Equal("OfficeIMO Export", batch.Title);

                var addSheetRequests = batch.Requests.OfType<GoogleSheetsAddSheetRequest>().ToList();
                Assert.Equal(2, addSheetRequests.Count);

                var summaryRequest = Assert.Single(addSheetRequests, r => r.SheetName == "Summary");
                Assert.Equal(0, summaryRequest.SheetIndex);
                Assert.False(summaryRequest.Hidden);
                Assert.True(summaryRequest.RightToLeft);
                Assert.Equal("FF336699", summaryRequest.TabColorArgb);
                Assert.Equal(1, summaryRequest.FrozenRowCount);
                Assert.Equal(1, summaryRequest.FrozenColumnCount);

                var hiddenRequest = Assert.Single(addSheetRequests, r => r.SheetName == "Hidden");
                Assert.True(hiddenRequest.Hidden);

                var updateRequest = Assert.Single(batch.Requests.OfType<GoogleSheetsUpdateCellsRequest>(), r => r.SheetName == "Summary");
                Assert.Contains(updateRequest.Cells, c => c.RowIndex == 1 && c.ColumnIndex == 0 && Equals(c.Value.Value, "Alpha"));
                Assert.Contains(updateRequest.Cells, c => c.RowIndex == 1 && c.ColumnIndex == 1 && c.Value.Kind == GoogleSheetsCellValueKind.Number && Equals(c.Value.Value, 12d));
                Assert.Contains(updateRequest.Cells, c => c.RowIndex == 1 && c.ColumnIndex == 2 && c.Value.Kind == GoogleSheetsCellValueKind.Boolean && Equals(c.Value.Value, true));
                Assert.Contains(updateRequest.Cells, c => c.RowIndex == 1 && c.ColumnIndex == 3 && c.Value.Kind == GoogleSheetsCellValueKind.DateTime);
                Assert.Contains(updateRequest.Cells, c => c.RowIndex == 1 && c.ColumnIndex == 4 && c.Value.Kind == GoogleSheetsCellValueKind.Formula && Equals(c.Value.Value, "=SUM(B2:B2)"));
                var hyperlinkCell = Assert.Single(updateRequest.Cells, c => c.RowIndex == 1 && c.ColumnIndex == 0);
                Assert.NotNull(hyperlinkCell.Hyperlink);
                Assert.True(hyperlinkCell.Hyperlink!.IsExternal);
                Assert.Equal("https://alpha.example/", hyperlinkCell.Hyperlink.Target);
                var internalHyperlinkCell = Assert.Single(updateRequest.Cells, c => c.RowIndex == 4 && c.ColumnIndex == 0);
                Assert.NotNull(internalHyperlinkCell.Hyperlink);
                Assert.False(internalHyperlinkCell.Hyperlink!.IsExternal);
                Assert.Equal("'Hidden'!B5", internalHyperlinkCell.Hyperlink.Target);
                var localNamedRangeHyperlinkCell = Assert.Single(updateRequest.Cells, c => c.RowIndex == 5 && c.ColumnIndex == 0);
                Assert.NotNull(localNamedRangeHyperlinkCell.Hyperlink);
                Assert.False(localNamedRangeHyperlinkCell.Hyperlink!.IsExternal);
                Assert.Equal("LocalData", localNamedRangeHyperlinkCell.Hyperlink.Target);
                Assert.NotNull(localNamedRangeHyperlinkCell.Comment);
                Assert.Equal("Tester (TT)", localNamedRangeHyperlinkCell.Comment!.Author);
                Assert.Equal("Jump note", localNamedRangeHyperlinkCell.Comment.Text);

                var styledCell = Assert.Single(updateRequest.Cells, c => c.RowIndex == 1 && c.ColumnIndex == 1);
                Assert.NotNull(styledCell.Style);
                Assert.Equal("0.00%", styledCell.Style!.NumberFormatCode);
                Assert.True(styledCell.Style.Bold);
                Assert.Equal("FF00FF00", styledCell.Style.FillColorArgb);
                Assert.Equal("center", styledCell.Style.HorizontalAlignment);
                Assert.NotNull(styledCell.Style.Borders);
                Assert.Equal("medium", styledCell.Style.Borders!.Left!.Style);
                Assert.Equal("FFFF0000", styledCell.Style.Borders.Left.ColorArgb);
                Assert.Equal("dashed", styledCell.Style.Borders.Top!.Style);
                Assert.Equal("FF0000FF", styledCell.Style.Borders.Top.ColorArgb);

                var fontColorCell = Assert.Single(updateRequest.Cells, c => c.RowIndex == 1 && c.ColumnIndex == 5);
                Assert.NotNull(fontColorCell.Style);
                Assert.Equal("FF112233", fontColorCell.Style!.FontColorArgb);
                Assert.NotNull(fontColorCell.Comment);
                Assert.Equal("Tester (TT)", fontColorCell.Comment!.Author);
                Assert.Equal("Accent comment", fontColorCell.Comment.Text);

                var dimensionRequests = batch.Requests.OfType<GoogleSheetsUpdateDimensionPropertiesRequest>().ToList();
                Assert.Contains(dimensionRequests, r => r.SheetName == "Summary" && r.DimensionKind == GoogleSheetsDimensionKind.Columns && r.StartIndex == 1 && r.EndIndexExclusive == 2 && r.PixelSize.HasValue && r.PixelSize.Value > 0);
                Assert.Contains(dimensionRequests, r => r.SheetName == "Summary" && r.DimensionKind == GoogleSheetsDimensionKind.Rows && r.StartIndex == 2 && r.EndIndexExclusive == 3 && r.PixelSize.HasValue && r.PixelSize.Value > 20);

                var tableRequest = Assert.Single(batch.Requests.OfType<GoogleSheetsAddTableRequest>(), r => r.SheetName == "Summary");
                Assert.Equal("SummaryTable", tableRequest.TableName);
                Assert.Equal("A1:B3", tableRequest.A1Range);
                Assert.True(tableRequest.TotalsRowShown);
                Assert.Equal("FFD9EAD3", tableRequest.FooterColorArgb);
                Assert.Equal(new[] { "Name", "Count" }, tableRequest.Columns.Select(c => c.Name).ToArray());
                Assert.Equal("TEXT", tableRequest.Columns[0].ColumnType);
                Assert.Equal("PERCENT", tableRequest.Columns[1].ColumnType);
                Assert.Equal("count", tableRequest.Columns[0].TotalsRowFunction);
                Assert.Equal("sum", tableRequest.Columns[1].TotalsRowFunction);

                var basicFilter = Assert.Single(batch.Requests.OfType<GoogleSheetsSetBasicFilterRequest>(), r => r.SheetName == "Summary");
                Assert.Equal("G1:J4", basicFilter.A1Range);
                Assert.Equal(4, basicFilter.Criteria.Count);
                var basicFilterValueCriteria = Assert.Single(basicFilter.Criteria, c => c.ColumnId == 0);
                Assert.Equal(new[] { "Closed" }, basicFilterValueCriteria.HiddenValues);
                var basicFilterContainsCriteria = Assert.Single(basicFilter.Criteria, c => c.ColumnId == 1);
                Assert.NotNull(basicFilterContainsCriteria.Condition);
                Assert.Equal("TEXT_CONTAINS", basicFilterContainsCriteria.Condition!.Type);
                Assert.Equal(new[] { "or" }, basicFilterContainsCriteria.Condition.Values);
                var basicFilterNumericCriteria = Assert.Single(basicFilter.Criteria, c => c.ColumnId == 2);
                Assert.NotNull(basicFilterNumericCriteria.Condition);
                Assert.Equal("NUMBER_GREATER_THAN_EQ", basicFilterNumericCriteria.Condition!.Type);
                Assert.Equal(new[] { "15" }, basicFilterNumericCriteria.Condition.Values);
                var basicFilterBetweenCriteria = Assert.Single(basicFilter.Criteria, c => c.ColumnId == 3);
                Assert.NotNull(basicFilterBetweenCriteria.Condition);
                Assert.Equal("NUMBER_BETWEEN", basicFilterBetweenCriteria.Condition!.Type);
                Assert.Equal(new[] { "10", "20" }, basicFilterBetweenCriteria.Condition.Values);

                var filterView = Assert.Single(batch.Requests.OfType<GoogleSheetsAddFilterViewRequest>(), r => r.SheetName == "Summary");
                Assert.Equal("SummaryTable Filter", filterView.Title);
                Assert.Equal("A1:B3", filterView.A1Range);

                var protectedRange = Assert.Single(batch.Requests.OfType<GoogleSheetsAddProtectedRangeRequest>(), r => r.SheetName == "Summary");
                Assert.False(protectedRange.WarningOnly);
                Assert.Contains("sort", protectedRange.Description, StringComparison.OrdinalIgnoreCase);
                Assert.Contains("autofilter", protectedRange.Description, StringComparison.OrdinalIgnoreCase);
                Assert.Contains("insert rows", protectedRange.Description, StringComparison.OrdinalIgnoreCase);

                var namedRanges = batch.Requests.OfType<GoogleSheetsAddNamedRangeRequest>().ToList();
                Assert.Equal(2, namedRanges.Count);
                Assert.Contains(namedRanges, r => r.Name == "GlobalData" && r.SheetName == null && r.A1Range == "'Summary'!$A$1:$B$2");
                Assert.Contains(namedRanges, r => r.Name == "LocalData" && r.SheetName == "Summary" && r.A1Range == "'Summary'!$A$2:$B$2");
                Assert.Contains(batch.Report.Notices, n => n.Feature == "SheetProtection");
                Assert.Contains(batch.Report.Notices, n => n.Feature == "SheetProtectionPermissions");
                Assert.Contains(batch.Report.Notices, n => n.Feature == "TableTotals");
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Test_GoogleSheetsBatchCompiler_TreatsBuiltInNamesAsDiagnostics() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleSheetsBatchCompilerBuiltInNames.xlsx");

            try {
                using var document = ExcelDocument.Create(filePath);
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, "Value");
                document.SetPrintArea(sheet, "A1:A5", save: false);

                var batch = document.CreateGoogleSheetsBatch();

                Assert.Empty(batch.Requests.OfType<GoogleSheetsAddNamedRangeRequest>());
                Assert.Contains(batch.Report.Notices, n => n.Feature == "BuiltInNames");
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Test_GoogleSheetsBatchCompiler_AndApiPayloadBuilder_MapAdvancedAutoFilterConditions() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleSheetsAdvancedAutoFilters.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var data = document.AddWorkSheet("Data");

                    data.CellValue(1, 1, "Prefix");
                    data.CellValue(1, 2, "Suffix");
                    data.CellValue(1, 3, "Variance");
                    data.CellValue(1, 4, "Notes");
                    data.CellValue(1, 5, "Delta");
                    data.CellValue(2, 1, "Ops-West");
                    data.CellValue(2, 2, "NorthOps");
                    data.CellValue(2, 3, 5d);
                    data.CellValue(2, 4, "keep");
                    data.CellValue(2, 5, 5d);
                    data.CellValue(3, 1, "Sales-East");
                    data.CellValue(3, 2, "WestTeam");
                    data.CellValue(3, 3, 15d);
                    data.CellValue(3, 4, "review later");
                    data.CellValue(3, 5, 10d);
                    data.CellValue(4, 1, "Ops-Central");
                    data.CellValue(4, 2, "FieldOps");
                    data.CellValue(4, 3, 25d);
                    data.CellValue(4, 4, "done");
                    data.CellValue(4, 5, 15d);

                    data.AddAutoFilter("A1:E4");
                    data.AutoFilterByHeaderStartsWith("Prefix", "Op");
                    data.AutoFilterByHeaderEndsWith("Suffix", "Ops");
                    data.AutoFilterByHeaderNotBetween("Variance", 10d, 20d);
                    data.AutoFilterByHeaderDoesNotContain("Notes", "view");
                    data.AutoFilterByHeaderNotEqual("Delta", 10d);

                    document.Save();
                }

                using var reloadedDocument = ExcelDocument.Load(filePath);
                var batch = reloadedDocument.CreateGoogleSheetsBatch(new GoogleSheetsSaveOptions {
                    Title = "Advanced Auto Filters"
                });

                var basicFilter = Assert.Single(batch.Requests.OfType<GoogleSheetsSetBasicFilterRequest>(), r => r.SheetName == "Data");
                Assert.Equal("A1:E4", basicFilter.A1Range);
                Assert.Equal(5, basicFilter.Criteria.Count);

                var startsWithCriteria = Assert.Single(basicFilter.Criteria, c => c.ColumnId == 0);
                Assert.NotNull(startsWithCriteria.Condition);
                Assert.Equal("TEXT_STARTS_WITH", startsWithCriteria.Condition!.Type);
                Assert.Equal(new[] { "Op" }, startsWithCriteria.Condition.Values);

                var endsWithCriteria = Assert.Single(basicFilter.Criteria, c => c.ColumnId == 1);
                Assert.NotNull(endsWithCriteria.Condition);
                Assert.Equal("TEXT_ENDS_WITH", endsWithCriteria.Condition!.Type);
                Assert.Equal(new[] { "Ops" }, endsWithCriteria.Condition.Values);

                var notBetweenCriteria = Assert.Single(basicFilter.Criteria, c => c.ColumnId == 2);
                Assert.NotNull(notBetweenCriteria.Condition);
                Assert.Equal("NUMBER_NOT_BETWEEN", notBetweenCriteria.Condition!.Type);
                Assert.Equal(new[] { "10", "20" }, notBetweenCriteria.Condition.Values);

                var notContainsCriteria = Assert.Single(basicFilter.Criteria, c => c.ColumnId == 3);
                Assert.NotNull(notContainsCriteria.Condition);
                Assert.Equal("TEXT_NOT_CONTAINS", notContainsCriteria.Condition!.Type);
                Assert.Equal(new[] { "view" }, notContainsCriteria.Condition.Values);

                var notEqualCriteria = Assert.Single(basicFilter.Criteria, c => c.ColumnId == 4);
                Assert.NotNull(notEqualCriteria.Condition);
                Assert.Equal("NUMBER_NOT_EQ", notEqualCriteria.Condition!.Type);
                Assert.Equal(new[] { "10" }, notEqualCriteria.Condition.Values);

                var payload = GoogleSheetsApiPayloadBuilder.BuildBatchUpdatePayload(batch, GoogleSheetsApiPayloadBuilder.BuildSheetIdMap(batch));
                var basicFilterRequest = Assert.Single(payload.Requests, r => r.SetBasicFilter != null);

                Assert.Equal("TEXT_STARTS_WITH", basicFilterRequest.SetBasicFilter!.Filter.Criteria!["0"].Condition!.Type);
                Assert.Equal("Op", Assert.Single(basicFilterRequest.SetBasicFilter.Filter.Criteria["0"].Condition!.Values!).UserEnteredValue);
                Assert.Equal("TEXT_ENDS_WITH", basicFilterRequest.SetBasicFilter.Filter.Criteria["1"].Condition!.Type);
                Assert.Equal("Ops", Assert.Single(basicFilterRequest.SetBasicFilter.Filter.Criteria["1"].Condition!.Values!).UserEnteredValue);
                Assert.Equal("NUMBER_NOT_BETWEEN", basicFilterRequest.SetBasicFilter.Filter.Criteria["2"].Condition!.Type);
                Assert.Equal(new[] { "10", "20" }, basicFilterRequest.SetBasicFilter.Filter.Criteria["2"].Condition!.Values!.Select(v => v.UserEnteredValue).ToArray());
                Assert.Equal("TEXT_NOT_CONTAINS", basicFilterRequest.SetBasicFilter.Filter.Criteria["3"].Condition!.Type);
                Assert.Equal("view", Assert.Single(basicFilterRequest.SetBasicFilter.Filter.Criteria["3"].Condition!.Values!).UserEnteredValue);
                Assert.Equal("NUMBER_NOT_EQ", basicFilterRequest.SetBasicFilter.Filter.Criteria["4"].Condition!.Type);
                Assert.Equal("10", Assert.Single(basicFilterRequest.SetBasicFilter.Filter.Criteria["4"].Condition!.Values!).UserEnteredValue);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Test_GoogleSheetsBatchCompiler_AndApiPayloadBuilder_MapNativeTableRowStyles() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleSheetsTableRowStyles.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var summary = document.AddWorkSheet("Summary");

                    summary.CellValue(1, 1, "Name");
                    summary.CellValue(1, 2, "Amount");
                    summary.CellValue(2, 1, "Alpha");
                    summary.CellValue(2, 2, 10d);
                    summary.CellValue(3, 1, "Beta");
                    summary.CellValue(3, 2, 20d);
                    summary.CellValue(4, 1, "Total");
                    summary.CellValue(4, 2, 30d);

                    summary.AddTable("A1:B4", hasHeader: true, name: "StyledTable", style: OfficeIMO.Excel.TableStyle.TableStyleMedium2, includeAutoFilter: true);
                    summary.SetTableTotals("A1:B4", new Dictionary<string, DocumentFormat.OpenXml.Spreadsheet.TotalsRowFunctionValues> {
                        ["Name"] = DocumentFormat.OpenXml.Spreadsheet.TotalsRowFunctionValues.Count,
                        ["Amount"] = DocumentFormat.OpenXml.Spreadsheet.TotalsRowFunctionValues.Sum,
                    });
                    summary.CellBackground(1, 1, "#B6D7A8");
                    summary.CellBackground(1, 2, "#B6D7A8");
                    summary.CellBackground(2, 1, "#E2F0D9");
                    summary.CellBackground(2, 2, "#E2F0D9");
                    summary.CellBackground(3, 1, "#FFF2CC");
                    summary.CellBackground(3, 2, "#FFF2CC");
                    summary.CellBackground(4, 1, "#D9EAD3");
                    summary.CellBackground(4, 2, "#D9EAD3");

                    document.Save();
                }

                using var reloadedDocument = ExcelDocument.Load(filePath);
                var batch = reloadedDocument.CreateGoogleSheetsBatch(new GoogleSheetsSaveOptions {
                    Title = "Styled Table Export"
                });

                var tableRequest = Assert.Single(batch.Requests.OfType<GoogleSheetsAddTableRequest>(), r => r.SheetName == "Summary");
                Assert.Equal("StyledTable", tableRequest.TableName);
                Assert.Equal("FFB6D7A8", tableRequest.HeaderColorArgb);
                Assert.Equal("FFE2F0D9", tableRequest.FirstBandColorArgb);
                Assert.Equal("FFFFF2CC", tableRequest.SecondBandColorArgb);
                Assert.Equal("FFD9EAD3", tableRequest.FooterColorArgb);

                var payload = GoogleSheetsApiPayloadBuilder.BuildBatchUpdatePayload(batch, GoogleSheetsApiPayloadBuilder.BuildSheetIdMap(batch));
                var apiTableRequest = Assert.Single(payload.Requests, r => r.AddTable != null);
                var rowsProperties = Assert.IsType<GoogleSheetsApiTableRowsPropertiesPayload>(apiTableRequest.AddTable!.Table.RowsProperties);
                var headerColor = Assert.IsType<GoogleSheetsApiColorPayload>(Assert.IsType<GoogleSheetsApiColorStylePayload>(rowsProperties.HeaderColorStyle).RgbColor);
                var firstBandColor = Assert.IsType<GoogleSheetsApiColorPayload>(Assert.IsType<GoogleSheetsApiColorStylePayload>(rowsProperties.FirstBandColorStyle).RgbColor);
                var secondBandColor = Assert.IsType<GoogleSheetsApiColorPayload>(Assert.IsType<GoogleSheetsApiColorStylePayload>(rowsProperties.SecondBandColorStyle).RgbColor);
                var footerColor = Assert.IsType<GoogleSheetsApiColorPayload>(Assert.IsType<GoogleSheetsApiColorStylePayload>(rowsProperties.FooterColorStyle).RgbColor);

                Assert.True(headerColor.Red > 0.71d);
                Assert.True(headerColor.Green > 0.84d);
                Assert.True(headerColor.Blue > 0.65d);

                Assert.True(firstBandColor.Red > 0.88d);
                Assert.True(firstBandColor.Green > 0.93d);
                Assert.True(firstBandColor.Blue > 0.84d);

                Assert.True(secondBandColor.Red > 0.99d);
                Assert.True(secondBandColor.Green > 0.94d);
                Assert.True(secondBandColor.Blue > 0.79d);

                Assert.True(footerColor.Red > 0.84d);
                Assert.True(footerColor.Green > 0.91d);
                Assert.True(footerColor.Blue > 0.82d);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Test_GoogleSheetsBatchCompiler_AndApiPayloadBuilder_MapTableDropdownValidation() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleSheetsTableDropdownValidation.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var summary = document.AddWorkSheet("Summary");

                    summary.CellValue(1, 1, "Name");
                    summary.CellValue(1, 2, "Status");
                    summary.CellValue(2, 1, "Alpha");
                    summary.CellValue(2, 2, "Open");
                    summary.CellValue(3, 1, "Beta");
                    summary.CellValue(3, 2, "Closed");

                    summary.AddTable("A1:B3", hasHeader: true, name: "StatusTable", style: OfficeIMO.Excel.TableStyle.TableStyleMedium2, includeAutoFilter: true);
                    summary.ValidationList("B2:B3", new[] { "Open", "Closed", "Pending" });

                    document.Save();
                }

                using var reloadedDocument = ExcelDocument.Load(filePath);
                var snapshot = reloadedDocument.CreateInspectionSnapshot();
                var summarySheet = Assert.Single(snapshot.Worksheets, worksheet => worksheet.Name == "Summary");
                var validation = Assert.Single(summarySheet.Validations);
                Assert.Equal("list", validation.Type);
                Assert.Equal("B2:B3", Assert.Single(validation.A1Ranges));
                Assert.Equal("\"Open,Closed,Pending\"", validation.Formula1);

                var batch = reloadedDocument.CreateGoogleSheetsBatch(new GoogleSheetsSaveOptions {
                    Title = "Dropdown Export"
                });

                var tableRequest = Assert.Single(batch.Requests.OfType<GoogleSheetsAddTableRequest>(), r => r.SheetName == "Summary");
                var statusColumn = Assert.Single(tableRequest.Columns, column => column.Name == "Status");
                Assert.Equal("DROPDOWN", statusColumn.ColumnType);
                Assert.NotNull(statusColumn.DataValidationRule);
                Assert.Equal("ONE_OF_LIST", statusColumn.DataValidationRule!.ConditionType);
                Assert.Equal(new[] { "Open", "Closed", "Pending" }, statusColumn.DataValidationRule.Values);
                Assert.True(statusColumn.DataValidationRule.Strict);
                Assert.True(statusColumn.DataValidationRule.ShowCustomUi);

                var payload = GoogleSheetsApiPayloadBuilder.BuildBatchUpdatePayload(batch, GoogleSheetsApiPayloadBuilder.BuildSheetIdMap(batch));
                var apiTableRequest = Assert.Single(payload.Requests, r => r.AddTable != null);
                var apiStatusColumn = Assert.Single(apiTableRequest.AddTable!.Table.ColumnProperties!, column => column.Name == "Status");
                Assert.Equal("DROPDOWN", apiStatusColumn.ColumnType);
                Assert.NotNull(apiStatusColumn.DataValidationRule);
                Assert.Equal("ONE_OF_LIST", apiStatusColumn.DataValidationRule!.Condition.Type);
                Assert.True(apiStatusColumn.DataValidationRule.Strict);
                Assert.True(apiStatusColumn.DataValidationRule.ShowCustomUi);
                Assert.Equal(new[] { "Open", "Closed", "Pending" }, apiStatusColumn.DataValidationRule.Condition.Values!.Select(value => value.UserEnteredValue).ToArray());
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Test_GoogleSheetsBatchCompiler_AndApiPayloadBuilder_MapNamedRangeBackedTableDropdownValidation() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleSheetsTableNamedRangeDropdownValidation.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var options = document.AddWorkSheet("Options");
                    var summary = document.AddWorkSheet("Summary");

                    options.CellValue(1, 1, "Open");
                    options.CellValue(2, 1, "Closed");
                    options.CellValue(3, 1, "Pending");
                    document.SetNamedRange("StatusOptions", "'Options'!A1:A3", save: false);

                    summary.CellValue(1, 1, "Name");
                    summary.CellValue(1, 2, "Status");
                    summary.CellValue(2, 1, "Alpha");
                    summary.CellValue(2, 2, "Open");
                    summary.CellValue(3, 1, "Beta");
                    summary.CellValue(3, 2, "Closed");

                    summary.AddTable("A1:B3", hasHeader: true, name: "StatusTable", style: OfficeIMO.Excel.TableStyle.TableStyleMedium2, includeAutoFilter: true);
                    summary.ValidationListNamedRange("B2:B3", "StatusOptions");

                    document.Save();
                }

                using var reloadedDocument = ExcelDocument.Load(filePath);
                var snapshot = reloadedDocument.CreateInspectionSnapshot();
                var summarySheet = Assert.Single(snapshot.Worksheets, worksheet => worksheet.Name == "Summary");
                var validation = Assert.Single(summarySheet.Validations);
                Assert.Equal("list", validation.Type);
                Assert.Equal("=StatusOptions", validation.Formula1);

                var batch = reloadedDocument.CreateGoogleSheetsBatch(new GoogleSheetsSaveOptions {
                    Title = "Named Dropdown Export"
                });

                var tableRequest = Assert.Single(batch.Requests.OfType<GoogleSheetsAddTableRequest>(), r => r.SheetName == "Summary");
                var statusColumn = Assert.Single(tableRequest.Columns, column => column.Name == "Status");
                Assert.Equal("DROPDOWN", statusColumn.ColumnType);
                Assert.NotNull(statusColumn.DataValidationRule);
                Assert.Equal("ONE_OF_LIST", statusColumn.DataValidationRule!.ConditionType);
                Assert.Equal(new[] { "Open", "Closed", "Pending" }, statusColumn.DataValidationRule.Values);

                var payload = GoogleSheetsApiPayloadBuilder.BuildBatchUpdatePayload(batch, GoogleSheetsApiPayloadBuilder.BuildSheetIdMap(batch));
                var apiTableRequest = Assert.Single(payload.Requests, r => r.AddTable != null);
                var apiStatusColumn = Assert.Single(apiTableRequest.AddTable!.Table.ColumnProperties!, column => column.Name == "Status");
                Assert.Equal("DROPDOWN", apiStatusColumn.ColumnType);
                Assert.NotNull(apiStatusColumn.DataValidationRule);
                Assert.Equal("ONE_OF_LIST", apiStatusColumn.DataValidationRule!.Condition.Type);
                Assert.Equal(new[] { "Open", "Closed", "Pending" }, apiStatusColumn.DataValidationRule.Condition.Values!.Select(value => value.UserEnteredValue).ToArray());
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Test_GoogleSheetsBatchCompiler_AndApiPayloadBuilder_MapRangeBackedTableDropdownValidation() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleSheetsTableRangeDropdownValidation.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var options = document.AddWorkSheet("Options");
                    var summary = document.AddWorkSheet("Summary");

                    options.CellValue(1, 1, "Open");
                    options.CellValue(2, 1, "Closed");
                    options.CellValue(3, 1, "Pending");

                    summary.CellValue(1, 1, "Name");
                    summary.CellValue(1, 2, "Status");
                    summary.CellValue(2, 1, "Alpha");
                    summary.CellValue(2, 2, "Open");
                    summary.CellValue(3, 1, "Beta");
                    summary.CellValue(3, 2, "Closed");

                    summary.AddTable("A1:B3", hasHeader: true, name: "StatusTable", style: OfficeIMO.Excel.TableStyle.TableStyleMedium2, includeAutoFilter: true);
                    summary.ValidationListRange("B2:B3", "A1:A3", "Options");

                    document.Save();
                }

                using var reloadedDocument = ExcelDocument.Load(filePath);
                var snapshot = reloadedDocument.CreateInspectionSnapshot();
                var summarySheet = Assert.Single(snapshot.Worksheets, worksheet => worksheet.Name == "Summary");
                var validation = Assert.Single(summarySheet.Validations);
                Assert.Equal("list", validation.Type);
                Assert.Equal("='Options'!A1:A3", validation.Formula1);

                var batch = reloadedDocument.CreateGoogleSheetsBatch(new GoogleSheetsSaveOptions {
                    Title = "Range Dropdown Export"
                });

                var tableRequest = Assert.Single(batch.Requests.OfType<GoogleSheetsAddTableRequest>(), r => r.SheetName == "Summary");
                var statusColumn = Assert.Single(tableRequest.Columns, column => column.Name == "Status");
                Assert.Equal("DROPDOWN", statusColumn.ColumnType);
                Assert.NotNull(statusColumn.DataValidationRule);
                Assert.Equal("ONE_OF_LIST", statusColumn.DataValidationRule!.ConditionType);
                Assert.Equal(new[] { "Open", "Closed", "Pending" }, statusColumn.DataValidationRule.Values);

                var payload = GoogleSheetsApiPayloadBuilder.BuildBatchUpdatePayload(batch, GoogleSheetsApiPayloadBuilder.BuildSheetIdMap(batch));
                var apiTableRequest = Assert.Single(payload.Requests, r => r.AddTable != null);
                var apiStatusColumn = Assert.Single(apiTableRequest.AddTable!.Table.ColumnProperties!, column => column.Name == "Status");
                Assert.Equal("DROPDOWN", apiStatusColumn.ColumnType);
                Assert.NotNull(apiStatusColumn.DataValidationRule);
                Assert.Equal("ONE_OF_LIST", apiStatusColumn.DataValidationRule!.Condition.Type);
                Assert.Equal(new[] { "Open", "Closed", "Pending" }, apiStatusColumn.DataValidationRule.Condition.Values!.Select(value => value.UserEnteredValue).ToArray());
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Test_GoogleSheetsBatchCompiler_AndApiPayloadBuilder_MapLocalRangeBackedTableDropdownValidation() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleSheetsTableLocalRangeDropdownValidation.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var summary = document.AddWorkSheet("Summary");

                    summary.CellValue(1, 1, "Name");
                    summary.CellValue(1, 2, "Status");
                    summary.CellValue(1, 4, "Allowed Status");
                    summary.CellValue(2, 1, "Alpha");
                    summary.CellValue(2, 2, "Open");
                    summary.CellValue(2, 4, "Open");
                    summary.CellValue(3, 1, "Beta");
                    summary.CellValue(3, 2, "Closed");
                    summary.CellValue(3, 4, "Closed");
                    summary.CellValue(4, 4, "Pending");

                    summary.AddTable("A1:B3", hasHeader: true, name: "StatusTable", style: OfficeIMO.Excel.TableStyle.TableStyleMedium2, includeAutoFilter: true);
                    summary.ValidationListRange("B2:B3", "D2:D4");

                    document.Save();
                }

                using var reloadedDocument = ExcelDocument.Load(filePath);
                var snapshot = reloadedDocument.CreateInspectionSnapshot();
                var summarySheet = Assert.Single(snapshot.Worksheets, worksheet => worksheet.Name == "Summary");
                var validation = Assert.Single(summarySheet.Validations);
                Assert.Equal("list", validation.Type);
                Assert.Equal("=D2:D4", validation.Formula1);

                var batch = reloadedDocument.CreateGoogleSheetsBatch(new GoogleSheetsSaveOptions {
                    Title = "Local Range Dropdown Export"
                });

                var tableRequest = Assert.Single(batch.Requests.OfType<GoogleSheetsAddTableRequest>(), r => r.SheetName == "Summary");
                var statusColumn = Assert.Single(tableRequest.Columns, column => column.Name == "Status");
                Assert.Equal("DROPDOWN", statusColumn.ColumnType);
                Assert.NotNull(statusColumn.DataValidationRule);
                Assert.Equal("ONE_OF_LIST", statusColumn.DataValidationRule!.ConditionType);
                Assert.Equal(new[] { "Open", "Closed", "Pending" }, statusColumn.DataValidationRule.Values);

                var payload = GoogleSheetsApiPayloadBuilder.BuildBatchUpdatePayload(batch, GoogleSheetsApiPayloadBuilder.BuildSheetIdMap(batch));
                var apiTableRequest = Assert.Single(payload.Requests, r => r.AddTable != null);
                var apiStatusColumn = Assert.Single(apiTableRequest.AddTable!.Table.ColumnProperties!, column => column.Name == "Status");
                Assert.Equal("DROPDOWN", apiStatusColumn.ColumnType);
                Assert.NotNull(apiStatusColumn.DataValidationRule);
                Assert.Equal("ONE_OF_LIST", apiStatusColumn.DataValidationRule!.Condition.Type);
                Assert.Equal(new[] { "Open", "Closed", "Pending" }, apiStatusColumn.DataValidationRule.Condition.Values!.Select(value => value.UserEnteredValue).ToArray());
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Test_GoogleSheetsBatchCompiler_AndApiPayloadBuilder_MapNumericCellValidations() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleSheetsNumericCellValidations.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var data = document.AddWorkSheet("Data");

                    data.CellValue(1, 1, "Quantity");
                    data.CellValue(1, 2, "Score");
                    data.CellValue(2, 1, 2);
                    data.CellValue(3, 1, 8);
                    data.CellValue(2, 2, 5.75);
                    data.CellValue(3, 2, 7.25);

                    data.ValidationWholeNumber("A2:A3", DocumentFormat.OpenXml.Spreadsheet.DataValidationOperatorValues.Between, 1, 10);
                    data.ValidationDecimal("B2:B3", DocumentFormat.OpenXml.Spreadsheet.DataValidationOperatorValues.GreaterThanOrEqual, 5.5);

                    document.Save();
                }

                using var reloadedDocument = ExcelDocument.Load(filePath);
                var snapshot = reloadedDocument.CreateInspectionSnapshot();
                var dataSheet = Assert.Single(snapshot.Worksheets, worksheet => worksheet.Name == "Data");
                Assert.Equal(2, dataSheet.Validations.Count);
                Assert.Contains(dataSheet.Validations, validation =>
                    validation.Type == "whole"
                    && validation.Operator == "between"
                    && validation.Formula1 == "1"
                    && validation.Formula2 == "10"
                    && validation.A1Ranges.SequenceEqual(new[] { "A2:A3" }));
                Assert.Contains(dataSheet.Validations, validation =>
                    validation.Type == "decimal"
                    && validation.Operator == "greaterThanOrEqual"
                    && validation.Formula1 == "5.5"
                    && validation.A1Ranges.SequenceEqual(new[] { "B2:B3" }));

                var batch = reloadedDocument.CreateGoogleSheetsBatch(new GoogleSheetsSaveOptions {
                    Title = "Numeric Validation Export"
                });

                var updateRequest = Assert.Single(batch.Requests.OfType<GoogleSheetsUpdateCellsRequest>(), r => r.SheetName == "Data");
                var quantityCell = Assert.Single(updateRequest.Cells, cell => cell.RowIndex == 1 && cell.ColumnIndex == 0);
                Assert.NotNull(quantityCell.DataValidationRule);
                Assert.Equal("NUMBER_BETWEEN", quantityCell.DataValidationRule!.ConditionType);
                Assert.Equal(new[] { "1", "10" }, quantityCell.DataValidationRule.Values);
                Assert.True(quantityCell.DataValidationRule.Strict);
                Assert.False(quantityCell.DataValidationRule.ShowCustomUi);

                var scoreCell = Assert.Single(updateRequest.Cells, cell => cell.RowIndex == 1 && cell.ColumnIndex == 1);
                Assert.NotNull(scoreCell.DataValidationRule);
                Assert.Equal("NUMBER_GREATER_THAN_EQ", scoreCell.DataValidationRule!.ConditionType);
                Assert.Equal(new[] { "5.5" }, scoreCell.DataValidationRule.Values);

                var payload = GoogleSheetsApiPayloadBuilder.BuildBatchUpdatePayload(batch, GoogleSheetsApiPayloadBuilder.BuildSheetIdMap(batch));
                var validationPayloads = payload.Requests
                    .Where(request => request.UpdateCells != null && request.UpdateCells.Fields.Contains("dataValidationRule"))
                    .Select(request => request.UpdateCells!)
                    .ToList();
                Assert.Equal(2, validationPayloads.Count);

                var firstDataRowPayload = Assert.Single(validationPayloads, request => request.Start.RowIndex == 1 && request.Start.ColumnIndex == 0);
                Assert.Contains("dataValidationRule", firstDataRowPayload.Fields);
                var quantityValidationRule = Assert.IsType<GoogleSheetsApiDataValidationRulePayload>(firstDataRowPayload.Rows[0].Values[0].DataValidationRule);
                var quantityCondition = Assert.IsType<GoogleSheetsApiBooleanConditionPayload>(quantityValidationRule.Condition);
                Assert.Equal("NUMBER_BETWEEN", quantityCondition.Type);
                Assert.Equal(new[] { "1", "10" }, Assert.IsAssignableFrom<IEnumerable<GoogleSheetsApiConditionValuePayload>>(quantityCondition.Values).Select(value => value.UserEnteredValue).ToArray());

                var scoreValidationRule = Assert.IsType<GoogleSheetsApiDataValidationRulePayload>(firstDataRowPayload.Rows[0].Values[1].DataValidationRule);
                var scoreCondition = Assert.IsType<GoogleSheetsApiBooleanConditionPayload>(scoreValidationRule.Condition);
                Assert.Equal("NUMBER_GREATER_THAN_EQ", scoreCondition.Type);
                Assert.Equal(new[] { "5.5" }, Assert.IsAssignableFrom<IEnumerable<GoogleSheetsApiConditionValuePayload>>(scoreCondition.Values).Select(value => value.UserEnteredValue).ToArray());

                Assert.Contains(batch.Report.Notices, notice => notice.Feature == "CellValidations");
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Test_GoogleSheetsBatchCompiler_AndApiPayloadBuilder_MapDateCellValidations() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleSheetsDateCellValidations.xlsx");
            var minimumDate = new DateTime(2024, 1, 1);
            var maximumDate = new DateTime(2024, 12, 31);

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var data = document.AddWorkSheet("Data");

                    data.CellValue(1, 1, "Start");
                    data.CellValue(2, 1, minimumDate);
                    data.CellValue(3, 1, new DateTime(2024, 6, 15));

                    data.ValidationDate("A2:A3", DocumentFormat.OpenXml.Spreadsheet.DataValidationOperatorValues.Between, minimumDate, maximumDate);

                    document.Save();
                }

                using var reloadedDocument = ExcelDocument.Load(filePath);
                var snapshot = reloadedDocument.CreateInspectionSnapshot();
                var dataSheet = Assert.Single(snapshot.Worksheets, worksheet => worksheet.Name == "Data");
                var validation = Assert.Single(dataSheet.Validations);
                Assert.Equal("date", validation.Type);
                Assert.Equal("between", validation.Operator);
                Assert.Equal(new[] { "A2:A3" }, validation.A1Ranges);

                var batch = reloadedDocument.CreateGoogleSheetsBatch(new GoogleSheetsSaveOptions {
                    Title = "Date Validation Export"
                });

                var updateRequest = Assert.Single(batch.Requests.OfType<GoogleSheetsUpdateCellsRequest>(), r => r.SheetName == "Data");
                var firstDateCell = Assert.Single(updateRequest.Cells, cell => cell.RowIndex == 1 && cell.ColumnIndex == 0);
                Assert.NotNull(firstDateCell.DataValidationRule);
                Assert.Equal("DATE_BETWEEN", firstDateCell.DataValidationRule!.ConditionType);
                Assert.Equal(new[] { "2024-01-01", "2024-12-31" }, firstDateCell.DataValidationRule.Values);

                var secondDateCell = Assert.Single(updateRequest.Cells, cell => cell.RowIndex == 2 && cell.ColumnIndex == 0);
                Assert.NotNull(secondDateCell.DataValidationRule);
                Assert.Equal("DATE_BETWEEN", secondDateCell.DataValidationRule!.ConditionType);
                Assert.Equal(new[] { "2024-01-01", "2024-12-31" }, secondDateCell.DataValidationRule.Values);

                var payload = GoogleSheetsApiPayloadBuilder.BuildBatchUpdatePayload(batch, GoogleSheetsApiPayloadBuilder.BuildSheetIdMap(batch));
                var validationPayload = Assert.Single(payload.Requests
                    .Where(request => request.UpdateCells != null
                        && request.UpdateCells.Fields.Contains("dataValidationRule")
                        && request.UpdateCells.Start.RowIndex == 1
                        && request.UpdateCells.Start.ColumnIndex == 0)
                    .Select(request => request.UpdateCells!));

                var firstValidationRule = Assert.IsType<GoogleSheetsApiDataValidationRulePayload>(validationPayload.Rows[0].Values[0].DataValidationRule);
                var firstValidationCondition = Assert.IsType<GoogleSheetsApiBooleanConditionPayload>(firstValidationRule.Condition);
                Assert.Equal("DATE_BETWEEN", firstValidationCondition.Type);
                Assert.Equal(new[] { "2024-01-01", "2024-12-31" }, Assert.IsAssignableFrom<IEnumerable<GoogleSheetsApiConditionValuePayload>>(firstValidationCondition.Values).Select(value => value.UserEnteredValue).ToArray());

                Assert.Contains(batch.Report.Notices, notice => notice.Feature == "CellValidations");
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Test_GoogleSheetsBatchCompiler_AndApiPayloadBuilder_MapTextLengthCellValidations() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleSheetsTextLengthCellValidations.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var data = document.AddWorkSheet("Data");

                    data.CellValue(1, 1, "Code");
                    data.CellValue(2, 1, "ABCD");
                    data.CellValue(3, 1, "WXYZ");

                    data.ValidationTextLength("A2:A3", DocumentFormat.OpenXml.Spreadsheet.DataValidationOperatorValues.Equal, 4);

                    document.Save();
                }

                using var reloadedDocument = ExcelDocument.Load(filePath);
                var snapshot = reloadedDocument.CreateInspectionSnapshot();
                var dataSheet = Assert.Single(snapshot.Worksheets, worksheet => worksheet.Name == "Data");
                var validation = Assert.Single(dataSheet.Validations);
                Assert.Equal("textlength", validation.Type);
                Assert.Equal("equal", validation.Operator);
                Assert.Equal("4", validation.Formula1);
                Assert.Equal(new[] { "A2:A3" }, validation.A1Ranges);

                var batch = reloadedDocument.CreateGoogleSheetsBatch(new GoogleSheetsSaveOptions {
                    Title = "Text Length Validation Export"
                });

                var updateRequest = Assert.Single(batch.Requests.OfType<GoogleSheetsUpdateCellsRequest>(), r => r.SheetName == "Data");
                var firstTextCell = Assert.Single(updateRequest.Cells, cell => cell.RowIndex == 1 && cell.ColumnIndex == 0);
                Assert.NotNull(firstTextCell.DataValidationRule);
                Assert.Equal("CUSTOM_FORMULA", firstTextCell.DataValidationRule!.ConditionType);
                Assert.Equal(new[] { "=LEN(A2)=4" }, firstTextCell.DataValidationRule.Values);

                var secondTextCell = Assert.Single(updateRequest.Cells, cell => cell.RowIndex == 2 && cell.ColumnIndex == 0);
                Assert.NotNull(secondTextCell.DataValidationRule);
                Assert.Equal("CUSTOM_FORMULA", secondTextCell.DataValidationRule!.ConditionType);
                Assert.Equal(new[] { "=LEN(A3)=4" }, secondTextCell.DataValidationRule.Values);

                var payload = GoogleSheetsApiPayloadBuilder.BuildBatchUpdatePayload(batch, GoogleSheetsApiPayloadBuilder.BuildSheetIdMap(batch));
                var validationPayload = Assert.Single(payload.Requests
                    .Where(request => request.UpdateCells != null
                        && request.UpdateCells.Fields.Contains("dataValidationRule")
                        && request.UpdateCells.Start.RowIndex == 1
                        && request.UpdateCells.Start.ColumnIndex == 0)
                    .Select(request => request.UpdateCells!));

                var firstValidationRule = Assert.IsType<GoogleSheetsApiDataValidationRulePayload>(validationPayload.Rows[0].Values[0].DataValidationRule);
                var firstValidationCondition = Assert.IsType<GoogleSheetsApiBooleanConditionPayload>(firstValidationRule.Condition);
                Assert.Equal("CUSTOM_FORMULA", firstValidationCondition.Type);
                Assert.Equal(new[] { "=LEN(A2)=4" }, Assert.IsAssignableFrom<IEnumerable<GoogleSheetsApiConditionValuePayload>>(firstValidationCondition.Values).Select(value => value.UserEnteredValue).ToArray());

                Assert.Contains(batch.Report.Notices, notice => notice.Feature == "CellValidations");
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Test_GoogleSheetsBatchCompiler_AndApiPayloadBuilder_EmitValidationOnlyEmptyCells() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleSheetsValidationOnlyEmptyCells.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var data = document.AddWorkSheet("Data");

                    data.CellValue(1, 1, "Quantity");
                    data.CellValue(2, 1, 2);
                    data.ValidationWholeNumber("A2:A4", DocumentFormat.OpenXml.Spreadsheet.DataValidationOperatorValues.Between, 1, 10);

                    document.Save();
                }

                using var reloadedDocument = ExcelDocument.Load(filePath);
                var batch = reloadedDocument.CreateGoogleSheetsBatch(new GoogleSheetsSaveOptions {
                    Title = "Validation Only Empty Cells Export"
                });

                var updateRequest = Assert.Single(batch.Requests.OfType<GoogleSheetsUpdateCellsRequest>(), r => r.SheetName == "Data");

                var populatedValidatedCell = Assert.Single(updateRequest.Cells, cell => cell.RowIndex == 1 && cell.ColumnIndex == 0);
                Assert.NotNull(populatedValidatedCell.DataValidationRule);
                Assert.Equal("NUMBER_BETWEEN", populatedValidatedCell.DataValidationRule!.ConditionType);
                Assert.Equal(new[] { "1", "10" }, populatedValidatedCell.DataValidationRule.Values);

                var emptyValidatedCellOne = Assert.Single(updateRequest.Cells, cell => cell.RowIndex == 2 && cell.ColumnIndex == 0);
                Assert.Equal(GoogleSheetsCellValueKind.Blank, emptyValidatedCellOne.Value.Kind);
                Assert.NotNull(emptyValidatedCellOne.DataValidationRule);
                Assert.Equal("NUMBER_BETWEEN", emptyValidatedCellOne.DataValidationRule!.ConditionType);
                Assert.Equal(new[] { "1", "10" }, emptyValidatedCellOne.DataValidationRule.Values);

                var emptyValidatedCellTwo = Assert.Single(updateRequest.Cells, cell => cell.RowIndex == 3 && cell.ColumnIndex == 0);
                Assert.Equal(GoogleSheetsCellValueKind.Blank, emptyValidatedCellTwo.Value.Kind);
                Assert.NotNull(emptyValidatedCellTwo.DataValidationRule);
                Assert.Equal("NUMBER_BETWEEN", emptyValidatedCellTwo.DataValidationRule!.ConditionType);
                Assert.Equal(new[] { "1", "10" }, emptyValidatedCellTwo.DataValidationRule.Values);

                var payload = GoogleSheetsApiPayloadBuilder.BuildBatchUpdatePayload(batch, GoogleSheetsApiPayloadBuilder.BuildSheetIdMap(batch));
                var validationPayloads = payload.Requests
                    .Where(request => request.UpdateCells != null
                        && request.UpdateCells.Fields.Contains("dataValidationRule")
                        && request.UpdateCells.Start.ColumnIndex == 0
                        && request.UpdateCells.Start.RowIndex >= 1
                        && request.UpdateCells.Start.RowIndex <= 3)
                    .Select(request => request.UpdateCells!)
                    .OrderBy(request => request.Start.RowIndex)
                    .ToList();

                Assert.Equal(3, validationPayloads.Count);
                Assert.Equal(1, validationPayloads[0].Start.RowIndex);
                Assert.Equal(2, validationPayloads[1].Start.RowIndex);
                Assert.Equal(3, validationPayloads[2].Start.RowIndex);

                foreach (var validationPayload in validationPayloads) {
                    var validationRule = Assert.IsType<GoogleSheetsApiDataValidationRulePayload>(validationPayload.Rows[0].Values[0].DataValidationRule);
                    var validationCondition = Assert.IsType<GoogleSheetsApiBooleanConditionPayload>(validationRule.Condition);
                    Assert.Equal("NUMBER_BETWEEN", validationCondition.Type);
                    Assert.Equal(new[] { "1", "10" }, Assert.IsAssignableFrom<IEnumerable<GoogleSheetsApiConditionValuePayload>>(validationCondition.Values).Select(value => value.UserEnteredValue).ToArray());
                }

                Assert.Contains(batch.Report.Notices, notice => notice.Feature == "CellValidations");
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Test_GoogleSheetsBatchCompiler_AndApiPayloadBuilder_EmitWorksheetListValidationsOutsideTables() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleSheetsWorksheetListValidationOutsideTables.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var data = document.AddWorkSheet("Data");

                    data.CellValue(1, 1, "Status");
                    data.CellValue(2, 1, "Open");
                    data.ValidationList("A2:A4", new[] { "Open", "Closed", "Pending" });

                    document.Save();
                }

                using var reloadedDocument = ExcelDocument.Load(filePath);
                var batch = reloadedDocument.CreateGoogleSheetsBatch(new GoogleSheetsSaveOptions {
                    Title = "Worksheet List Validation Export"
                });

                var updateRequest = Assert.Single(batch.Requests.OfType<GoogleSheetsUpdateCellsRequest>(), r => r.SheetName == "Data");

                var populatedDropdownCell = Assert.Single(updateRequest.Cells, cell => cell.RowIndex == 1 && cell.ColumnIndex == 0);
                Assert.NotNull(populatedDropdownCell.DataValidationRule);
                Assert.Equal("ONE_OF_LIST", populatedDropdownCell.DataValidationRule!.ConditionType);
                Assert.Equal(new[] { "Open", "Closed", "Pending" }, populatedDropdownCell.DataValidationRule.Values);
                Assert.True(populatedDropdownCell.DataValidationRule.ShowCustomUi);

                var emptyDropdownCellOne = Assert.Single(updateRequest.Cells, cell => cell.RowIndex == 2 && cell.ColumnIndex == 0);
                Assert.Equal(GoogleSheetsCellValueKind.Blank, emptyDropdownCellOne.Value.Kind);
                Assert.NotNull(emptyDropdownCellOne.DataValidationRule);
                Assert.Equal("ONE_OF_LIST", emptyDropdownCellOne.DataValidationRule!.ConditionType);
                Assert.Equal(new[] { "Open", "Closed", "Pending" }, emptyDropdownCellOne.DataValidationRule.Values);

                var emptyDropdownCellTwo = Assert.Single(updateRequest.Cells, cell => cell.RowIndex == 3 && cell.ColumnIndex == 0);
                Assert.Equal(GoogleSheetsCellValueKind.Blank, emptyDropdownCellTwo.Value.Kind);
                Assert.NotNull(emptyDropdownCellTwo.DataValidationRule);
                Assert.Equal("ONE_OF_LIST", emptyDropdownCellTwo.DataValidationRule!.ConditionType);
                Assert.Equal(new[] { "Open", "Closed", "Pending" }, emptyDropdownCellTwo.DataValidationRule.Values);

                var payload = GoogleSheetsApiPayloadBuilder.BuildBatchUpdatePayload(batch, GoogleSheetsApiPayloadBuilder.BuildSheetIdMap(batch));
                var validationPayloads = payload.Requests
                    .Where(request => request.UpdateCells != null
                        && request.UpdateCells.Fields.Contains("dataValidationRule")
                        && request.UpdateCells.Start.ColumnIndex == 0
                        && request.UpdateCells.Start.RowIndex >= 1
                        && request.UpdateCells.Start.RowIndex <= 3)
                    .Select(request => request.UpdateCells!)
                    .OrderBy(request => request.Start.RowIndex)
                    .ToList();

                Assert.Equal(3, validationPayloads.Count);

                foreach (var validationPayload in validationPayloads) {
                    var validationRule = Assert.IsType<GoogleSheetsApiDataValidationRulePayload>(validationPayload.Rows[0].Values[0].DataValidationRule);
                    var validationCondition = Assert.IsType<GoogleSheetsApiBooleanConditionPayload>(validationRule.Condition);
                    Assert.Equal("ONE_OF_LIST", validationCondition.Type);
                    Assert.Equal(new[] { "Open", "Closed", "Pending" }, Assert.IsAssignableFrom<IEnumerable<GoogleSheetsApiConditionValuePayload>>(validationCondition.Values).Select(value => value.UserEnteredValue).ToArray());
                }

                Assert.Contains(batch.Report.Notices, notice => notice.Feature == "CellValidations");
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Test_GoogleSheetsApiPayloadBuilder_TranslatesNeutralBatchToSheetsPayloads() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleSheetsApiPayloadBuilder.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var summary = document.AddWorkSheet("Summary");
                    var hidden = document.AddWorkSheet("Hidden");

                    summary.CellValue(1, 1, "Name");
                    summary.CellValue(2, 2, 12);
                    summary.SetHyperlink(2, 1, "https://alpha.example/", display: "Alpha");
                    summary.SetComment(2, 1, "External link note", author: "Tester", initials: "TT");
                    summary.FormatCell(2, 2, "0.00%");
                    summary.CellBackground(2, 2, "#00FF00");
                    summary.CellBold(2, 2, true);
                    summary.CellAlign(2, 2, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center);
                    summary.SetColumnWidth(2, 20);
                    summary.CellValue(3, 2, "Wrapped\nRow");
                    summary.WrapCells(3, 3, 2, 20);
                    summary.AutoFitRow(3);
                    summary.CellValue(1, 7, "Status");
                    summary.CellValue(1, 8, "Region");
                    summary.CellValue(1, 9, "Score");
                    summary.CellValue(1, 10, "Budget");
                    summary.CellValue(2, 7, "Open");
                    summary.CellValue(2, 8, "North");
                    summary.CellValue(2, 9, 10d);
                    summary.CellValue(2, 10, 8d);
                    summary.CellValue(3, 7, "Closed");
                    summary.CellValue(3, 8, "South");
                    summary.CellValue(3, 9, 20d);
                    summary.CellValue(3, 10, 18d);
                    summary.CellValue(4, 7, "Open");
                    summary.CellValue(4, 8, "East");
                    summary.CellValue(4, 9, 30d);
                    summary.CellValue(4, 10, 28d);
                    summary.AddTable("A1:B3", hasHeader: true, name: "SummaryTable", style: OfficeIMO.Excel.TableStyle.TableStyleMedium2, includeAutoFilter: true);
                    summary.SetTableTotals("A1:B3", new Dictionary<string, DocumentFormat.OpenXml.Spreadsheet.TotalsRowFunctionValues> {
                        ["Name"] = DocumentFormat.OpenXml.Spreadsheet.TotalsRowFunctionValues.Count,
                        ["Count"] = DocumentFormat.OpenXml.Spreadsheet.TotalsRowFunctionValues.Sum,
                    });
                    summary.CellBackground(3, 1, "#D9EAD3");
                    summary.CellBackground(3, 2, "#D9EAD3");
                    summary.AddAutoFilter("G1:J4", new Dictionary<uint, IEnumerable<string>> {
                        { 0, new[] { "Open" } }
                    });
                    summary.AutoFilterByHeaderContains("Region", "or");
                    summary.AutoFilterByHeaderGreaterThanOrEqual("Score", 15d);
                    summary.AutoFilterByHeaderBetween("Budget", 10d, 20d);
                    summary.Freeze(topRows: 1, leftCols: 1);
                    summary.Protect(new ExcelSheetProtectionOptions {
                        AllowSelectLockedCells = false,
                        AllowSelectUnlockedCells = false,
                        AllowSort = true,
                        AllowAutoFilter = true,
                        AllowInsertRows = true,
                    });
                    summary.SetInternalLink(5, 1, hidden, "B5", display: "Go hidden");
                    summary.SetInternalLink(6, 1, "LocalData", display: "Go local");
                    summary.SetComment(6, 1, "Jump note", author: "Tester", initials: "TT");
                    hidden.SetHidden(true);
                    document.SetNamedRange("GlobalData", "'Summary'!A1:B3", save: false);
                    summary.SetNamedRange("LocalData", "B2:B3", save: false);
                    document.Save();
                }

                ApplyBorderToCell(filePath, "Summary", "B2");
                ApplySheetDisplaySettings(filePath, "Summary", "FF336699", rightToLeft: true);

                using var reloadedDocument = ExcelDocument.Load(filePath);
                var batch = reloadedDocument.CreateGoogleSheetsBatch(new GoogleSheetsSaveOptions {
                    Title = "API Export"
                });

                var createPayload = GoogleSheetsApiPayloadBuilder.BuildCreateSpreadsheetPayload(batch);
                Assert.Equal("API Export", createPayload.Properties.Title);
                Assert.Equal(2, createPayload.Sheets.Count);
                Assert.Contains(createPayload.Sheets, s =>
                    s.Properties.SheetId == 1
                    && s.Properties.Title == "Summary"
                    && s.Properties.GridProperties.FrozenRowCount == 1
                    && s.Properties.RightToLeft == true
                    && s.Properties.TabColor != null
                    && s.Properties.TabColor.Red > 0.19d
                    && s.Properties.TabColor.Blue > 0.59d);
                Assert.Contains(createPayload.Sheets, s => s.Properties.SheetId == 2 && s.Properties.Title == "Hidden" && s.Properties.Hidden);

                var batchPayload = GoogleSheetsApiPayloadBuilder.BuildBatchUpdatePayload(batch, GoogleSheetsApiPayloadBuilder.BuildSheetIdMap(batch), "spread123");
                Assert.NotEmpty(batchPayload.Requests);

                Assert.Contains(batchPayload.Requests, r =>
                    r.UpdateDimensionProperties != null
                    && r.UpdateDimensionProperties.Range.Dimension == "COLUMNS"
                    && r.UpdateDimensionProperties.Range.SheetId == 1
                    && r.UpdateDimensionProperties.Range.StartIndex == 1);

                Assert.Contains(batchPayload.Requests, r =>
                    r.UpdateDimensionProperties != null
                    && r.UpdateDimensionProperties.Range.Dimension == "ROWS"
                    && r.UpdateDimensionProperties.Range.SheetId == 1
                    && r.UpdateDimensionProperties.Range.StartIndex == 2);

                var basicFilterRequest = Assert.Single(batchPayload.Requests, r => r.SetBasicFilter != null);
                Assert.Equal(1, basicFilterRequest.SetBasicFilter!.Filter.Range.SheetId);
                Assert.Equal(6, basicFilterRequest.SetBasicFilter.Filter.Range.StartColumnIndex);
                Assert.Contains("Closed", basicFilterRequest.SetBasicFilter.Filter.Criteria!["0"].HiddenValues!);
                Assert.Equal("TEXT_CONTAINS", basicFilterRequest.SetBasicFilter.Filter.Criteria["1"].Condition!.Type);
                Assert.Equal("or", Assert.Single(basicFilterRequest.SetBasicFilter.Filter.Criteria["1"].Condition!.Values!).UserEnteredValue);
                Assert.Equal("NUMBER_GREATER_THAN_EQ", basicFilterRequest.SetBasicFilter.Filter.Criteria["2"].Condition!.Type);
                Assert.Equal("15", Assert.Single(basicFilterRequest.SetBasicFilter.Filter.Criteria["2"].Condition!.Values!).UserEnteredValue);
                Assert.Equal("NUMBER_BETWEEN", basicFilterRequest.SetBasicFilter.Filter.Criteria["3"].Condition!.Type);
                Assert.Equal(new[] { "10", "20" }, basicFilterRequest.SetBasicFilter.Filter.Criteria["3"].Condition!.Values!.Select(v => v.UserEnteredValue).ToArray());

                var filterViewRequest = Assert.Single(batchPayload.Requests, r => r.AddFilterView != null);
                Assert.Equal("SummaryTable Filter", filterViewRequest.AddFilterView!.Filter.Title);
                Assert.Equal(0, filterViewRequest.AddFilterView.Filter.Range.StartColumnIndex);

                var tableRequest = Assert.Single(batchPayload.Requests, r => r.AddTable != null);
                Assert.Equal("SummaryTable", tableRequest.AddTable!.Table.Name);
                Assert.Equal(1, tableRequest.AddTable.Table.Range.SheetId);
                var rowsProperties = Assert.IsType<GoogleSheetsApiTableRowsPropertiesPayload>(tableRequest.AddTable.Table.RowsProperties);
                var footerColorStyle = Assert.IsType<GoogleSheetsApiColorStylePayload>(rowsProperties.FooterColorStyle);
                var footerColor = Assert.IsType<GoogleSheetsApiColorPayload>(footerColorStyle.RgbColor);
                Assert.True(footerColor.Red > 0.84d);
                Assert.True(footerColor.Green > 0.91d);
                Assert.True(footerColor.Blue > 0.82d);
                Assert.Equal("Name", tableRequest.AddTable.Table.ColumnProperties![0].Name);
                Assert.Equal("PERCENT", tableRequest.AddTable.Table.ColumnProperties[1].ColumnType);

                var protectedRange = Assert.Single(batchPayload.Requests, r => r.AddProtectedRange != null);
                Assert.Equal(1, protectedRange.AddProtectedRange!.ProtectedRange.Range.SheetId);
                Assert.False(protectedRange.AddProtectedRange.ProtectedRange.WarningOnly);
                Assert.Contains("sort", protectedRange.AddProtectedRange.ProtectedRange.Description, StringComparison.OrdinalIgnoreCase);

                var hyperlinkCell = batchPayload.Requests
                    .Where(r => r.UpdateCells != null)
                    .SelectMany(r => r.UpdateCells!.Rows)
                    .SelectMany(r => r.Values)
                    .First(c => c.UserEnteredValue?.FormulaValue != null && c.UserEnteredValue.FormulaValue.Contains("HYPERLINK", StringComparison.Ordinal));
                Assert.Contains("https://alpha.example/", hyperlinkCell.UserEnteredValue!.FormulaValue);
                Assert.Equal("Tester (TT): External link note", hyperlinkCell.Note);

                var internalHyperlinkCell = batchPayload.Requests
                    .Where(r => r.UpdateCells != null)
                    .SelectMany(r => r.UpdateCells!.Rows)
                    .SelectMany(r => r.Values)
                    .First(c => c.UserEnteredValue?.FormulaValue != null && c.UserEnteredValue.FormulaValue.Contains("gid=2", StringComparison.Ordinal));
                Assert.Contains("docs.google.com/spreadsheets/d/spread123/edit#gid=2", internalHyperlinkCell.UserEnteredValue!.FormulaValue);
                Assert.Equal("OfficeIMO internal link target: Hidden!B5", internalHyperlinkCell.Note);

                var localNamedRangeHyperlinkCell = batchPayload.Requests
                    .Where(r => r.UpdateCells != null)
                    .SelectMany(r => r.UpdateCells!.Rows)
                    .SelectMany(r => r.Values)
                    .First(c => c.UserEnteredValue?.FormulaValue != null && c.UserEnteredValue.FormulaValue.Contains("Go local", StringComparison.Ordinal));
                Assert.Contains("docs.google.com/spreadsheets/d/spread123/edit#gid=1", localNamedRangeHyperlinkCell.UserEnteredValue!.FormulaValue);
                Assert.Equal("Tester (TT): Jump note" + Environment.NewLine + Environment.NewLine + "OfficeIMO internal link target: LocalData -> Summary!B2:B3", localNamedRangeHyperlinkCell.Note);

                var styledCell = batchPayload.Requests
                    .Where(r => r.UpdateCells != null)
                    .SelectMany(r => r.UpdateCells!.Rows)
                    .SelectMany(r => r.Values)
                    .First(c => c.UserEnteredFormat?.HorizontalAlignment == "CENTER");
                Assert.Equal("PERCENT", styledCell.UserEnteredFormat!.NumberFormat!.Type);
                Assert.Equal("0.00%", styledCell.UserEnteredFormat.NumberFormat.Pattern);
                Assert.Equal("CENTER", styledCell.UserEnteredFormat.HorizontalAlignment);
                Assert.NotNull(styledCell.UserEnteredFormat.Borders);
                Assert.Equal("SOLID_MEDIUM", styledCell.UserEnteredFormat.Borders!.Left!.Style);
                Assert.Equal(1d, styledCell.UserEnteredFormat.Borders.Left.Color!.Red);
                Assert.Equal("DASHED", styledCell.UserEnteredFormat.Borders.Top!.Style);
                Assert.Equal(1d, styledCell.UserEnteredFormat.Borders.Top.Color!.Blue);

                var namedRange = Assert.Single(batchPayload.Requests, r => r.AddNamedRange?.NamedRange.Name == "GlobalData");
                Assert.Equal("GlobalData", namedRange.AddNamedRange!.NamedRange.Name);
                Assert.Equal(1, namedRange.AddNamedRange.NamedRange.Range.SheetId);
                Assert.Equal(0, namedRange.AddNamedRange.NamedRange.Range.StartRowIndex);
                Assert.Equal(2, namedRange.AddNamedRange.NamedRange.Range.EndColumnIndex);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public async Task Test_GoogleSheetsExporter_UsesConfiguredHttpPipeline_ForCreateFlow() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleSheetsExporterCreate.xlsx");

            try {
                using var document = ExcelDocument.Create(filePath);
                var summary = document.AddWorkSheet("Summary");
                var target = document.AddWorkSheet("Target");
                summary.CellValue(1, 1, "Name");
                summary.SetHyperlink(2, 1, "https://alpha.example/", display: "Alpha");
                summary.SetInternalLink(3, 1, target, "B5", display: "Target");
                summary.SetNamedRange("LocalData", "B2:B3", save: false);
                summary.SetInternalLink(4, 1, "LocalData", display: "Local");
                summary.SetComment(4, 1, "Jump note", author: "Tester", initials: "TT");
                summary.CellValue(2, 2, 5);

                var recordedRequests = new List<(Uri Uri, string Method, string? Body, string? Authorization)>();
                using var httpClient = new HttpClient(new FakeHttpMessageHandler(async request => {
                    string? body = request.Content == null ? null : await request.Content.ReadAsStringAsync().ConfigureAwait(false);
                    recordedRequests.Add((request.RequestUri!, request.Method.Method, body, request.Headers.Authorization?.ToString()));

                    if (request.Method == HttpMethod.Post && request.RequestUri!.AbsoluteUri == "https://sheets.googleapis.com/v4/spreadsheets") {
                        return CreateJsonResponse("{\"spreadsheetId\":\"spread123\",\"spreadsheetUrl\":\"https://docs.google.com/spreadsheets/d/spread123/edit\",\"properties\":{\"title\":\"Create Export\"}}");
                    }

                    if (request.Method == HttpMethod.Post && request.RequestUri!.AbsoluteUri == "https://sheets.googleapis.com/v4/spreadsheets/spread123:batchUpdate") {
                        return CreateJsonResponse("{}");
                    }

                    return new HttpResponseMessage(HttpStatusCode.NotFound) {
                        Content = new StringContent("unexpected request", Encoding.UTF8, "text/plain")
                    };
                }));

                var session = new GoogleWorkspaceSession(
                    new FakeGoogleWorkspaceCredentialSource(),
                    new GoogleWorkspaceSessionOptions {
                        HttpClient = httpClient,
                    });

                var result = await document.ExportToGoogleSheetsAsync(session, new GoogleSheetsSaveOptions {
                    Title = "Create Export",
                });

                Assert.Equal("spread123", result.SpreadsheetId);
                Assert.Equal("https://docs.google.com/spreadsheets/d/spread123/edit", result.WebViewLink);
                Assert.Equal(2, recordedRequests.Count);
                Assert.All(recordedRequests, r => Assert.Equal("Bearer fake-access-token", r.Authorization));

                var createRequest = Assert.Single(recordedRequests, r => r.Uri.AbsoluteUri == "https://sheets.googleapis.com/v4/spreadsheets");
                Assert.Equal("POST", createRequest.Method);
                using (var json = JsonDocument.Parse(createRequest.Body!)) {
                    Assert.Equal("Create Export", json.RootElement.GetProperty("properties").GetProperty("title").GetString());
                }

                var updateRequest = Assert.Single(recordedRequests, r => r.Uri.AbsoluteUri == "https://sheets.googleapis.com/v4/spreadsheets/spread123:batchUpdate");
                Assert.Equal("POST", updateRequest.Method);
                Assert.Contains("HYPERLINK", updateRequest.Body!);
                Assert.Contains("docs.google.com/spreadsheets/d/spread123/edit#gid=2", updateRequest.Body!);
                Assert.Contains("OfficeIMO internal link target: Target!B5", updateRequest.Body!);
                Assert.Contains("docs.google.com/spreadsheets/d/spread123/edit#gid=1", updateRequest.Body!);
                Assert.Contains("Tester (TT): Jump note", updateRequest.Body!);
                Assert.Contains("OfficeIMO internal link target: LocalData -\\u003E Summary!B2:B3", updateRequest.Body!);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public async Task Test_GoogleSheetsExporter_MovesCreatedSpreadsheet_ToRequestedFolder() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleSheetsExporterCreateMove.xlsx");

            try {
                using var document = ExcelDocument.Create(filePath);
                var summary = document.AddWorkSheet("Summary");
                summary.CellValue(1, 1, "Name");
                summary.CellValue(2, 1, "Alpha");

                var recordedRequests = new List<(Uri Uri, string Method, string? Body)>();
                using var httpClient = new HttpClient(new FakeHttpMessageHandler(async request => {
                    string? body = request.Content == null ? null : await request.Content.ReadAsStringAsync().ConfigureAwait(false);
                    recordedRequests.Add((request.RequestUri!, request.Method.Method, body));

                    if (request.Method == HttpMethod.Post && request.RequestUri!.AbsoluteUri == "https://sheets.googleapis.com/v4/spreadsheets") {
                        return CreateJsonResponse("{\"spreadsheetId\":\"spreadMove\",\"spreadsheetUrl\":\"https://docs.google.com/spreadsheets/d/spreadMove/edit\",\"properties\":{\"title\":\"Move Export\"}}");
                    }

                    if (request.Method == HttpMethod.Post && request.RequestUri!.AbsoluteUri == "https://sheets.googleapis.com/v4/spreadsheets/spreadMove:batchUpdate") {
                        return CreateJsonResponse("{}");
                    }

                    if (request.Method == HttpMethod.Get && request.RequestUri!.AbsoluteUri == "https://www.googleapis.com/drive/v3/files/spreadMove?fields=id,parents,webViewLink&supportsAllDrives=true") {
                        return CreateJsonResponse("{\"id\":\"spreadMove\",\"parents\":[\"oldParent\"],\"webViewLink\":\"https://docs.google.com/spreadsheets/d/spreadMove/edit\"}");
                    }

                    if (string.Equals(request.Method.Method, "PATCH", StringComparison.Ordinal) && request.RequestUri!.AbsoluteUri.Contains("https://www.googleapis.com/drive/v3/files/spreadMove?", StringComparison.Ordinal)) {
                        return CreateJsonResponse("{\"id\":\"spreadMove\",\"parents\":[\"folder123\"],\"webViewLink\":\"https://docs.google.com/spreadsheets/d/spreadMove/edit\"}");
                    }

                    return new HttpResponseMessage(HttpStatusCode.NotFound) {
                        Content = new StringContent("unexpected request", Encoding.UTF8, "text/plain")
                    };
                }));

                var session = new GoogleWorkspaceSession(
                    new FakeGoogleWorkspaceCredentialSource(),
                    new GoogleWorkspaceSessionOptions {
                        HttpClient = httpClient,
                    });

                var result = await document.ExportToGoogleSheetsAsync(session, new GoogleSheetsSaveOptions {
                    Title = "Move Export",
                    Location = new GoogleDriveFileLocation {
                        FolderId = "folder123",
                        SharedDriveAware = true,
                    }
                });

                Assert.Equal("spreadMove", result.SpreadsheetId);
                Assert.Equal(4, recordedRequests.Count);
                Assert.Contains(recordedRequests, r => r.Method == "GET" && r.Uri.AbsoluteUri.Contains("/drive/v3/files/spreadMove?", StringComparison.Ordinal));
                var patchRequest = Assert.Single(recordedRequests, r => r.Method == "PATCH");
                Assert.Contains("addParents=folder123", patchRequest.Uri.Query);
                Assert.Contains("removeParents=oldParent", patchRequest.Uri.Query);
                Assert.DoesNotContain(result.Report.Notices, n => n.Feature == "DrivePlacement" && n.Severity >= OfficeIMO.GoogleWorkspace.TranslationSeverity.Warning);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public async Task Test_GoogleSheetsExporter_CanReplaceExistingSpreadsheet() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleSheetsExporterUpdate.xlsx");

            try {
                using var document = ExcelDocument.Create(filePath);
                var summary = document.AddWorkSheet("Summary");
                summary.CellValue(1, 1, "Name");
                summary.CellValue(2, 1, "Alpha");
                summary.SetColumnWidth(1, 18);

                var recordedRequests = new List<(Uri Uri, string Method, string? Body)>();
                using var httpClient = new HttpClient(new FakeHttpMessageHandler(async request => {
                    string? body = request.Content == null ? null : await request.Content.ReadAsStringAsync().ConfigureAwait(false);
                    recordedRequests.Add((request.RequestUri!, request.Method.Method, body));

                    if (request.Method == HttpMethod.Get && request.RequestUri!.AbsoluteUri.Contains("/v4/spreadsheets/existing123?", StringComparison.Ordinal)) {
                        return CreateJsonResponse("{\"spreadsheetId\":\"existing123\",\"spreadsheetUrl\":\"https://docs.google.com/spreadsheets/d/existing123/edit\",\"properties\":{\"title\":\"Old Title\"},\"sheets\":[{\"properties\":{\"sheetId\":7}},{\"properties\":{\"sheetId\":8}}]}");
                    }

                    if (request.Method == HttpMethod.Post && request.RequestUri!.AbsoluteUri == "https://sheets.googleapis.com/v4/spreadsheets/existing123:batchUpdate") {
                        return CreateJsonResponse("{}");
                    }

                    return new HttpResponseMessage(HttpStatusCode.NotFound) {
                        Content = new StringContent("unexpected request", Encoding.UTF8, "text/plain")
                    };
                }));

                var session = new GoogleWorkspaceSession(
                    new FakeGoogleWorkspaceCredentialSource(),
                    new GoogleWorkspaceSessionOptions {
                        HttpClient = httpClient,
                    });

                var result = await document.ExportToGoogleSheetsAsync(session, new GoogleSheetsSaveOptions {
                    Title = "Replacement Export",
                    Location = new GoogleDriveFileLocation {
                        ExistingFileId = "existing123",
                    }
                });

                Assert.Equal("existing123", result.SpreadsheetId);
                Assert.Equal("https://docs.google.com/spreadsheets/d/existing123/edit", result.WebViewLink);
                Assert.Equal(3, recordedRequests.Count);
                Assert.Equal("GET", recordedRequests[0].Method);
                Assert.Equal("POST", recordedRequests[1].Method);
                Assert.Equal("POST", recordedRequests[2].Method);
                Assert.DoesNotContain(recordedRequests, r => r.Uri.AbsoluteUri == "https://sheets.googleapis.com/v4/spreadsheets");
                Assert.Contains(result.Report.Notices, n => n.Feature == "ExistingSpreadsheet");

                using (var resetJson = JsonDocument.Parse(recordedRequests[1].Body!)) {
                    var requests = resetJson.RootElement.GetProperty("requests");
                    Assert.True(requests.GetArrayLength() >= 3);
                    var requestKinds = requests.EnumerateArray()
                        .SelectMany(r => r.EnumerateObject().Select(p => p.Name))
                        .ToList();
                    Assert.Contains("deleteSheet", requestKinds);
                    Assert.Contains("addSheet", requestKinds);
                    Assert.Contains("updateSpreadsheetProperties", requestKinds);
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        private static HttpResponseMessage CreateJsonResponse(string json) {
            return new HttpResponseMessage(HttpStatusCode.OK) {
                Content = new StringContent(json, Encoding.UTF8, "application/json")
            };
        }

        private static void ApplyBorderToCell(string filePath, string sheetName, string cellReference) {
            using var document = SpreadsheetDocument.Open(filePath, true);
            var workbookPart = document.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart is missing.");
            var stylesheet = workbookPart.WorkbookStylesPart?.Stylesheet ?? throw new InvalidOperationException("Stylesheet is missing.");
            var sheet = workbookPart.Workbook.Sheets?.Elements<Sheet>().FirstOrDefault(s => string.Equals(s.Name?.Value, sheetName, StringComparison.Ordinal))
                ?? throw new InvalidOperationException($"Sheet '{sheetName}' was not found.");
            var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!);
            var cell = worksheetPart.Worksheet.Descendants<Cell>().FirstOrDefault(c => string.Equals(c.CellReference?.Value, cellReference, StringComparison.Ordinal))
                ?? throw new InvalidOperationException($"Cell '{cellReference}' was not found.");

            stylesheet.Borders ??= new Borders(new Border());
            stylesheet.Borders.Count = (uint)stylesheet.Borders.Count();
            stylesheet.CellFormats ??= new CellFormats(new CellFormat());
            stylesheet.CellFormats.Count = (uint)stylesheet.CellFormats.Count();

            var border = new Border(
                new LeftBorder(new Color { Rgb = "FFFF0000" }) { Style = BorderStyleValues.Medium },
                new RightBorder(),
                new TopBorder(new Color { Rgb = "FF0000FF" }) { Style = BorderStyleValues.Dashed },
                new BottomBorder(),
                new DiagonalBorder());

            stylesheet.Borders.Append(border);
            stylesheet.Borders.Count = (uint)stylesheet.Borders.Count();
            var borderId = stylesheet.Borders.Count!.Value - 1;

            var existingStyleIndex = cell.StyleIndex?.Value ?? 0U;
            var existingFormat = stylesheet.CellFormats.Elements<CellFormat>().ElementAtOrDefault((int)existingStyleIndex) ?? new CellFormat();
            var clonedFormat = (CellFormat)existingFormat.CloneNode(true);
            clonedFormat.BorderId = borderId;
            clonedFormat.ApplyBorder = true;
            stylesheet.CellFormats.Append(clonedFormat);
            stylesheet.CellFormats.Count = (uint)stylesheet.CellFormats.Count();
            cell.StyleIndex = stylesheet.CellFormats.Count!.Value - 1;

            stylesheet.Save();
            worksheetPart.Worksheet.Save();
            workbookPart.Workbook.Save();
        }

        private static void ApplySheetDisplaySettings(string filePath, string sheetName, string tabColorArgb, bool rightToLeft) {
            using var document = SpreadsheetDocument.Open(filePath, true);
            var workbookPart = document.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart is missing.");
            var sheet = workbookPart.Workbook.Sheets?.Elements<Sheet>().FirstOrDefault(s => string.Equals(s.Name?.Value, sheetName, StringComparison.Ordinal))
                ?? throw new InvalidOperationException($"Sheet '{sheetName}' was not found.");
            var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!);
            var worksheet = worksheetPart.Worksheet;

            var sheetProperties = worksheet.GetFirstChild<SheetProperties>();
            if (sheetProperties == null) {
                sheetProperties = new SheetProperties();
                worksheet.InsertAt(sheetProperties, 0);
            }

            sheetProperties.TabColor = new TabColor {
                Rgb = tabColorArgb,
            };

            var sheetViews = worksheet.GetFirstChild<SheetViews>();
            if (sheetViews == null) {
                sheetViews = new SheetViews();
                worksheet.InsertAfter(sheetViews, sheetProperties);
            }

            var sheetView = sheetViews.Elements<SheetView>().FirstOrDefault();
            if (sheetView == null) {
                sheetView = new SheetView {
                    WorkbookViewId = 0U,
                };
                sheetViews.Append(sheetView);
            }

            sheetView.RightToLeft = rightToLeft;

            worksheet.Save();
            workbookPart.Workbook.Save();
        }

        private sealed class FakeGoogleWorkspaceCredentialSource : IGoogleWorkspaceCredentialSource {
            public Task<GoogleWorkspaceAccessToken> AcquireAccessTokenAsync(IEnumerable<string> scopes, CancellationToken cancellationToken = default) {
                return Task.FromResult(new GoogleWorkspaceAccessToken(
                    "fake-access-token",
                    DateTimeOffset.UtcNow.AddHours(1),
                    scopes.ToList()));
            }
        }

        private sealed class FakeHttpMessageHandler : HttpMessageHandler {
            private readonly Func<HttpRequestMessage, Task<HttpResponseMessage>> _handler;

            public FakeHttpMessageHandler(Func<HttpRequestMessage, Task<HttpResponseMessage>> handler) {
                _handler = handler ?? throw new ArgumentNullException(nameof(handler));
            }

            protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken) {
                return _handler(request);
            }
        }
    }
}
