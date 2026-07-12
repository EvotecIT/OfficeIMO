using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using OfficeIMO.Excel.LegacyXls;
using OfficeIMO.Excel.LegacyXls.Model;
using System.Globalization;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void LegacyXls_NativeSave_WritesAdditionalWorksheetAutoFilterCriteria() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorkSheet("AdvancedFilter");
                    sheet.CellValue(1, 1, "Optional");
                    sheet.CellValue(1, 2, "Required");
                    sheet.CellValue(1, 3, "Score");
                    sheet.CellValue(1, 4, "Ratio");
                    sheet.CellValue(2, 1, string.Empty);
                    sheet.CellValue(2, 2, "Alpha");
                    sheet.CellValue(2, 3, 95d);
                    sheet.CellValue(2, 4, 0.1d);
                    sheet.CellValue(3, 1, "Filled");
                    sheet.CellValue(3, 2, "Beta");
                    sheet.CellValue(3, 3, 85d);
                    sheet.CellValue(3, 4, 0.2d);
                    sheet.CellValue(4, 1, string.Empty);
                    sheet.CellValue(4, 2, "Gamma");
                    sheet.CellValue(4, 3, 75d);
                    sheet.CellValue(4, 4, 0.3d);
                    sheet.CellValue(5, 1, "Filled");
                    sheet.CellValue(5, 2, "Delta");
                    sheet.CellValue(5, 3, 65d);
                    sheet.CellValue(5, 4, 0.4d);

                    const string range = "A1:D5";
                    sheet.AutoFilterAdd(range);
                    AutoFilter autoFilter = Assert.Single(sheet.WorksheetPart.Worksheet.Elements<AutoFilter>());
                    FilterColumn? existingColumn = autoFilter.Elements<FilterColumn>().FirstOrDefault(column => column.ColumnId?.Value == 0U);
                    existingColumn?.Remove();
                    autoFilter.Append(new FilterColumn(
                        new Filters(
                            new Filter { Val = "Filled" }) {
                            Blank = true
                        }) {
                        ColumnId = 0U
                    });
                    sheet.ApplyAutoFilterCustomCriteria(
                        range,
                        1,
                        matchAll: false,
                        new[] { (FilterOperatorValues.NotEqual, " ") });
                    sheet.ApplyAutoFilterTop10Criteria(range, 2, 3, isTop: true, isPercent: false);
                    sheet.ApplyAutoFilterTop10Criteria(range, 3, 25, isTop: false, isPercent: true);

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                Assert.False(result.HasImportErrors);
                Assert.False(result.HasUnsupportedFeatures);

                LegacyXlsWorksheet legacySheet = Assert.Single(result.Workbook.Worksheets);
                Assert.Equal((ushort)4, legacySheet.AutoFilterDropDownCount);
                Assert.Equal(4, legacySheet.AutoFilterCriteria.Count);

                LegacyXlsAutoFilterCriteria blankCriteria = Assert.Single(legacySheet.AutoFilterCriteria, criteria => criteria.ColumnId == 0U);
                Assert.Equal(LegacyXlsAutoFilterKind.Custom, blankCriteria.Kind);
                Assert.Equal(LegacyXlsAutoFilterJoinOperator.Or, blankCriteria.JoinOperator);
                Assert.Contains(blankCriteria.Conditions, condition =>
                    condition.ValueKind == LegacyXlsAutoFilterValueKind.Text
                    && condition.Operator == LegacyXlsAutoFilterOperator.Equal
                    && condition.Value == "Filled");
                Assert.Contains(blankCriteria.Conditions, condition =>
                    condition.ValueKind == LegacyXlsAutoFilterValueKind.Blank
                    && condition.Operator == LegacyXlsAutoFilterOperator.Equal);

                LegacyXlsAutoFilterCriteria nonBlankCriteria = Assert.Single(legacySheet.AutoFilterCriteria, criteria => criteria.ColumnId == 1U);
                Assert.Equal(LegacyXlsAutoFilterKind.NonBlanks, nonBlankCriteria.Kind);
                Assert.Equal(LegacyXlsAutoFilterValueKind.NonBlank, Assert.Single(nonBlankCriteria.Conditions).ValueKind);

                LegacyXlsAutoFilterCriteria topItemsCriteria = Assert.Single(legacySheet.AutoFilterCriteria, criteria => criteria.ColumnId == 2U);
                Assert.Equal(LegacyXlsAutoFilterKind.Top10, topItemsCriteria.Kind);
                Assert.Equal((ushort)3, topItemsCriteria.Top10Value);
                Assert.True(topItemsCriteria.Top10IsTop);
                Assert.False(topItemsCriteria.Top10IsPercent);

                LegacyXlsAutoFilterCriteria bottomPercentCriteria = Assert.Single(legacySheet.AutoFilterCriteria, criteria => criteria.ColumnId == 3U);
                Assert.Equal(LegacyXlsAutoFilterKind.Top10, bottomPercentCriteria.Kind);
                Assert.Equal((ushort)25, bottomPercentCriteria.Top10Value);
                Assert.False(bottomPercentCriteria.Top10IsTop);
                Assert.True(bottomPercentCriteria.Top10IsPercent);

                AutoFilter projectedAutoFilter = Assert.Single(result.Document.Sheets[0].WorksheetPart.Worksheet.Elements<AutoFilter>());
                Assert.Equal("A1:D5", projectedAutoFilter.Reference!.Value);
                List<FilterColumn> projectedColumns = projectedAutoFilter.Elements<FilterColumn>().OrderBy(column => column.ColumnId?.Value ?? 0U).ToList();
                Assert.Equal(4, projectedColumns.Count);

                Filters projectedBlankOrValue = projectedColumns[0].GetFirstChild<Filters>()!;
                Assert.True(projectedBlankOrValue.Blank!.Value);
                Assert.Equal("Filled", Assert.Single(projectedBlankOrValue.Elements<Filter>()).Val!.Value);

                CustomFilter projectedNonBlank = Assert.Single(projectedColumns[1].GetFirstChild<CustomFilters>()!.Elements<CustomFilter>());
                Assert.Equal(FilterOperatorValues.NotEqual, projectedNonBlank.Operator!.Value);
                Assert.Equal(" ", projectedNonBlank.Val!.Value);

                Top10 projectedTopItems = projectedColumns[2].GetFirstChild<Top10>()!;
                Assert.True(projectedTopItems.Top!.Value);
                Assert.False(projectedTopItems.Percent!.Value);
                Assert.Equal(3d, projectedTopItems.Val!.Value);

                Top10 projectedBottomPercent = projectedColumns[3].GetFirstChild<Top10>()!;
                Assert.False(projectedBottomPercent.Top!.Value);
                Assert.True(projectedBottomPercent.Percent!.Value);
                Assert.Equal(25d, projectedBottomPercent.Val!.Value);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesDateGroupAutoFilterCriteriaAsSerialRange() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    document.DateSystem = ExcelDateSystem.NineteenFour;

                    ExcelSheet sheet = document.AddWorkSheet("DateFilter");
                    sheet.CellValue(1, 1, "Entered");
                    sheet.CellValue(2, 1, new DateTime(2026, 6, 27));
                    sheet.CellValue(3, 1, new DateTime(2026, 6, 28));
                    sheet.CellValue(4, 1, new DateTime(2026, 6, 29));
                    sheet.CellValue(5, 1, new DateTime(2026, 7, 1));

                    sheet.AutoFilterAdd("A1:A5");
                    AutoFilter autoFilter = Assert.Single(sheet.WorksheetPart.Worksheet.Elements<AutoFilter>());
                    autoFilter.Append(new FilterColumn(
                        new Filters(
                            new DateGroupItem {
                                Year = 2026,
                                Month = 6,
                                Day = 28,
                                DateTimeGrouping = DateTimeGroupingValues.Day
                            })) {
                        ColumnId = 0U
                    });
                    sheet.WorksheetPart.Worksheet.Save();

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                Assert.False(result.HasImportErrors);
                Assert.False(result.HasUnsupportedFeatures);
                Assert.True(result.Workbook.Uses1904DateSystem);

                LegacyXlsWorksheet legacySheet = Assert.Single(result.Workbook.Worksheets);
                LegacyXlsAutoFilterCriteria criteria = Assert.Single(legacySheet.AutoFilterCriteria);
                Assert.Equal(0U, criteria.ColumnId);
                Assert.Equal(LegacyXlsAutoFilterKind.Custom, criteria.Kind);
                Assert.True(criteria.MatchAll);
                Assert.Equal(LegacyXlsAutoFilterJoinOperator.And, criteria.JoinOperator);
                Assert.Equal(2, criteria.Conditions.Count);

                double expectedStart = new DateTime(2026, 6, 28).ToOADate() - 1462d;
                double expectedEnd = new DateTime(2026, 6, 29).ToOADate() - 1462d;
                Assert.Equal(LegacyXlsAutoFilterOperator.GreaterThanOrEqual, criteria.Conditions[0].Operator);
                Assert.Equal(LegacyXlsAutoFilterValueKind.Number, criteria.Conditions[0].ValueKind);
                Assert.Equal(expectedStart, double.Parse(criteria.Conditions[0].Value, CultureInfo.InvariantCulture), 6);
                Assert.Equal(LegacyXlsAutoFilterOperator.LessThan, criteria.Conditions[1].Operator);
                Assert.Equal(LegacyXlsAutoFilterValueKind.Number, criteria.Conditions[1].ValueKind);
                Assert.Equal(expectedEnd, double.Parse(criteria.Conditions[1].Value, CultureInfo.InvariantCulture), 6);

                AutoFilter projectedAutoFilter = Assert.Single(result.Document.Sheets[0].WorksheetPart.Worksheet.Elements<AutoFilter>());
                CustomFilters projectedFilters = Assert.Single(projectedAutoFilter.Elements<FilterColumn>()).GetFirstChild<CustomFilters>()!;
                Assert.True(projectedFilters.And!.Value);
                List<CustomFilter> projectedConditions = projectedFilters.Elements<CustomFilter>().ToList();
                Assert.Equal(2, projectedConditions.Count);
                Assert.Equal(FilterOperatorValues.GreaterThanOrEqual, projectedConditions[0].Operator!.Value);
                Assert.Equal(expectedStart, double.Parse(projectedConditions[0].Val!.Value!, CultureInfo.InvariantCulture), 6);
                Assert.Equal(FilterOperatorValues.LessThan, projectedConditions[1].Operator!.Value);
                Assert.Equal(expectedEnd, double.Parse(projectedConditions[1].Val!.Value!, CultureInfo.InvariantCulture), 6);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }
    }
}
