using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_AddPivotTableAndReadBack() {
            var filePath = Path.Combine(_directoryWithFiles, "ExcelPivotTableBasic.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");

                sheet.CellValue(1, 1, "Region");
                sheet.CellValue(1, 2, "Product");
                sheet.CellValue(1, 3, "Sales");

                sheet.CellValue(2, 1, "East");
                sheet.CellValue(2, 2, "A");
                sheet.CellValue(2, 3, 10);

                sheet.CellValue(3, 1, "West");
                sheet.CellValue(3, 2, "A");
                sheet.CellValue(3, 3, 12);

                sheet.CellValue(4, 1, "East");
                sheet.CellValue(4, 2, "B");
                sheet.CellValue(4, 3, 7);

                sheet.AddPivotTable(
                    sourceRange: "A1:C4",
                    destinationCell: "E2",
                    name: "SalesPivot",
                    rowFields: new[] { "Region" },
                    columnFields: new[] { "Product" },
                    dataFields: new[] { new ExcelPivotDataField("Sales", DataConsolidateFunctionValues.Sum, "Total Sales") },
                    pivotStyleName: "PivotStyleMedium9");

                document.Save(false);
            }

            using (var document = ExcelDocument.Load(filePath)) {
                var pivots = document.GetPivotTables();
                Assert.Single(pivots);

                var pivot = pivots[0];
                Assert.Equal("SalesPivot", pivot.Name);
                Assert.Equal("Data", pivot.SheetName);
                Assert.Equal("Data", pivot.SourceSheet);
                Assert.Equal("A1:C4", pivot.SourceRange);
                Assert.Equal("PivotStyleMedium9", pivot.PivotStyle);
                Assert.Equal(ExcelPivotLayout.Compact, pivot.Layout);

                Assert.Contains("Region", pivot.RowFields);
                Assert.Contains("Product", pivot.ColumnFields);

                Assert.Single(pivot.DataFields);
                var dataField = pivot.DataFields[0];
                Assert.Equal("Sales", dataField.FieldName);
                Assert.Equal(DataConsolidateFunctionValues.Sum, dataField.Function);
                Assert.Equal("Total Sales", dataField.DisplayName);
            }

            using (var document = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_FluentPivotBuilder_CreatesPivotTable() {
            var filePath = Path.Combine(_directoryWithFiles, "ExcelPivotTableFluent.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");

                sheet.CellValue(1, 1, "Region");
                sheet.CellValue(1, 2, "Product");
                sheet.CellValue(1, 3, "Sales");

                sheet.CellValue(2, 1, "East");
                sheet.CellValue(2, 2, "A");
                sheet.CellValue(2, 3, 10);

                sheet.CellValue(3, 1, "West");
                sheet.CellValue(3, 2, "A");
                sheet.CellValue(3, 3, 12);

                sheet.CellValue(4, 1, "East");
                sheet.CellValue(4, 2, "B");
                sheet.CellValue(4, 3, 7);

                sheet.Pivot("A1:C4")
                    .Rows("Region")
                    .Columns("Product")
                    .Sum("Sales", "Total Sales", "$#,##0")
                    .Style("PivotStyleMedium9")
                    .Layout(ExcelPivotLayout.Tabular)
                    .GrandTotals(rows: false, columns: true)
                    .Captions(rowHeader: "Rows", columnHeader: "Products", grandTotal: "Total")
                    .At("E2", "SalesPivot");

                document.Save(false);
            }

            using (var document = ExcelDocument.Load(filePath)) {
                var pivot = document.GetPivotTables().Single();

                Assert.Equal("SalesPivot", pivot.Name);
                Assert.Equal("A1:C4", pivot.SourceRange);
                Assert.Equal("PivotStyleMedium9", pivot.PivotStyle);
                Assert.Equal(ExcelPivotLayout.Tabular, pivot.Layout);
                Assert.False(pivot.RowGrandTotals);
                Assert.True(pivot.ColumnGrandTotals);
                Assert.Equal("Rows", pivot.RowHeaderCaption);
                Assert.Equal("Products", pivot.ColumnHeaderCaption);
                Assert.Equal("Total", pivot.GrandTotalCaption);
                Assert.Contains("Region", pivot.RowFields);
                Assert.Contains("Product", pivot.ColumnFields);

                var dataField = pivot.DataFields.Single();
                Assert.Equal("Sales", dataField.FieldName);
                Assert.Equal(DataConsolidateFunctionValues.Sum, dataField.Function);
                Assert.Equal("Total Sales", dataField.DisplayName);
                Assert.True(dataField.NumberFormatId >= 164);
            }

            using (var document = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_PivotPublicCompatibilityOverloadsRemainAvailable() {
            Assert.NotNull(typeof(ExcelPivotDataField).GetConstructor(new[] {
                typeof(string),
                typeof(DataConsolidateFunctionValues?),
                typeof(string),
                typeof(uint?)
            }));

            Assert.NotNull(typeof(ExcelPivotDataFieldInfo).GetConstructor(new[] {
                typeof(string),
                typeof(DataConsolidateFunctionValues),
                typeof(string)
            }));

            Assert.NotNull(typeof(ExcelPivotTableInfo).GetConstructor(new[] {
                typeof(string),
                typeof(uint),
                typeof(string),
                typeof(string),
                typeof(string),
                typeof(string),
                typeof(int),
                typeof(string),
                typeof(ExcelPivotLayout),
                typeof(bool?),
                typeof(bool?),
                typeof(bool?),
                typeof(bool?),
                typeof(bool?),
                typeof(IReadOnlyList<string>),
                typeof(IReadOnlyList<string>),
                typeof(IReadOnlyList<string>),
                typeof(IReadOnlyList<ExcelPivotDataFieldInfo>)
            }));

            Assert.NotNull(typeof(ExcelSheet).GetMethod("AddPivotTable", new[] {
                typeof(string),
                typeof(string),
                typeof(string),
                typeof(IEnumerable<string>),
                typeof(IEnumerable<string>),
                typeof(IEnumerable<string>),
                typeof(IEnumerable<ExcelPivotDataField>),
                typeof(bool),
                typeof(bool),
                typeof(string),
                typeof(ExcelPivotLayout),
                typeof(bool?),
                typeof(bool?),
                typeof(bool?),
                typeof(bool?),
                typeof(bool?)
            }));
        }

        [Fact]
        public void Test_AddPivotTable_AppliesFieldOptionsAndNumberFormats() {
            var filePath = Path.Combine(_directoryWithFiles, "ExcelPivotTableFieldOptions.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");

                sheet.CellValue(1, 1, "Region");
                sheet.CellValue(1, 2, "Product");
                sheet.CellValue(1, 3, "Sales");

                sheet.CellValue(2, 1, "East");
                sheet.CellValue(2, 2, "A");
                sheet.CellValue(2, 3, 10.25);

                sheet.CellValue(3, 1, "West");
                sheet.CellValue(3, 2, "A");
                sheet.CellValue(3, 3, 12.5);

                sheet.CellValue(4, 1, "East");
                sheet.CellValue(4, 2, "B");
                sheet.CellValue(4, 3, 7.75);

                sheet.CellValue(5, 2, "C");
                sheet.CellValue(5, 3, 3.5);

                sheet.AddPivotTable(
                    sourceRange: "A1:C5",
                    destinationCell: "E2",
                    name: "SalesPivot",
                    rowFields: new[] { "Region" },
                    pageFields: new[] { "Product" },
                    dataFields: new[] { new ExcelPivotDataField("Sales", DataConsolidateFunctionValues.Sum, "Total Sales", numberFormat: "$#,##0.00") },
                    fieldOptions: new[] {
                        new ExcelPivotFieldOptions("Region",
                            sortType: FieldSortValues.Ascending,
                            defaultSubtotal: false,
                            subtotalTop: true,
                            insertBlankRow: true,
                            insertPageBreak: true,
                            compact: false,
                            outline: true,
                            showDropDowns: true,
                            multipleItemSelectionAllowed: true,
                            includeNewItemsInFilter: true,
                            subtotalCaption: "Region subtotal",
                            hiddenItems: new[] { "West" }),
                        new ExcelPivotFieldOptions("Product", selectedItem: "A")
                    },
                    rowHeaderCaption: "Rows",
                    columnHeaderCaption: "Columns",
                    grandTotalCaption: "Grand total",
                    missingCaption: "(missing)",
                    errorCaption: "(error)",
                    showDataDropDown: false,
                    showDropZones: true,
                    showDataTips: true,
                    showMemberPropertyTips: true,
                    fieldListSortAscending: true,
                    customListSort: false);

                document.Save(false);
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                var workbookPart = spreadsheet.WorkbookPart!;
                var pivotPart = workbookPart.WorksheetParts.SelectMany(part => part.PivotTableParts).Single();
                var pivotDefinition = pivotPart.PivotTableDefinition!;

                var cacheField = pivotPart.PivotTableCacheDefinitionPart!.PivotCacheDefinition!.CacheFields!.Elements<CacheField>().ElementAt(0);
                var sharedItems = cacheField.SharedItems!.ChildElements;
                Assert.Equal(3, sharedItems.Count);
                Assert.IsType<MissingItem>(sharedItems[2]);

                var regionField = pivotDefinition.PivotFields!.Elements<PivotField>().ElementAt(0);
                Assert.Equal(FieldSortValues.Ascending, regionField.SortType!.Value);
                Assert.False(regionField.DefaultSubtotal!.Value);
                Assert.True(regionField.SubtotalTop!.Value);
                Assert.True(regionField.InsertBlankRow!.Value);
                Assert.True(regionField.InsertPageBreak!.Value);
                Assert.False(regionField.Compact!.Value);
                Assert.True(regionField.Outline!.Value);
                Assert.True(regionField.ShowDropDowns!.Value);
                Assert.True(regionField.MultipleItemSelectionAllowed!.Value);
                Assert.True(regionField.IncludeNewItemsInFilter!.Value);
                Assert.Equal("Region subtotal", regionField.SubtotalCaption!.Value);

                var regionItems = regionField.Items!.Elements<Item>().ToList();
                Assert.Equal(3, regionItems.Count);
                Assert.False(regionItems[0].Hidden?.Value ?? false);
                Assert.True(regionItems[1].Hidden!.Value);
                Assert.False(regionItems[2].Hidden?.Value ?? false);

                var pageField = pivotDefinition.PageFields!.Elements<PageField>().Single();
                Assert.Equal(1, pageField.Field!.Value);
                Assert.Equal(0U, pageField.Item!.Value);

                Assert.Equal("Rows", pivotDefinition.RowHeaderCaption!.Value);
                Assert.Equal("Columns", pivotDefinition.ColumnHeaderCaption!.Value);
                Assert.Equal("Grand total", pivotDefinition.GrandTotalCaption!.Value);
                Assert.Equal("(missing)", pivotDefinition.MissingCaption!.Value);
                Assert.Equal("(error)", pivotDefinition.ErrorCaption!.Value);
                Assert.False(pivotDefinition.ShowDataDropDown!.Value);
                Assert.True(pivotDefinition.ShowDropZones!.Value);
                Assert.True(pivotDefinition.ShowDataTips!.Value);
                Assert.True(pivotDefinition.ShowMemberPropertyTips!.Value);
                Assert.True(pivotDefinition.FieldListSortAscending!.Value);
                Assert.False(pivotDefinition.CustomListSort!.Value);

                var dataField = pivotDefinition.DataFields!.Elements<DataField>().Single();
                uint numberFormatId = dataField.NumberFormatId!.Value;
                Assert.True(numberFormatId >= 164);

                var numberFormat = workbookPart.WorkbookStylesPart!.Stylesheet!.NumberingFormats!
                    .Elements<NumberingFormat>()
                    .Single(format => format.NumberFormatId!.Value == numberFormatId);
                Assert.Equal("$#,##0.00", numberFormat.FormatCode!.Value);
            }

            using (var document = ExcelDocument.Load(filePath)) {
                var pivot = document.GetPivotTables().Single();
                Assert.True(pivot.RowGrandTotals);
                Assert.True(pivot.ColumnGrandTotals);
                Assert.Equal("Rows", pivot.RowHeaderCaption);
                Assert.Equal("Columns", pivot.ColumnHeaderCaption);
                Assert.Equal("Grand total", pivot.GrandTotalCaption);
                Assert.False(pivot.ShowDataDropDown);
                Assert.True(pivot.ShowDropZones);
                Assert.Contains("Product", pivot.PageFields);

                var region = pivot.Fields.Single(field => field.FieldName == "Region");
                Assert.Equal(FieldSortValues.Ascending, region.SortType);
                Assert.False(region.DefaultSubtotal);
                Assert.True(region.InsertPageBreak);
                Assert.False(region.Compact);
                Assert.True(region.Outline);
                Assert.Contains("West", region.HiddenItems);

                var dataField = pivot.DataFields.Single();
                Assert.True(dataField.NumberFormatId >= 164);
            }

            using (var document = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_AddPivotChartFromRange_WritesPivotSource() {
            var filePath = Path.Combine(_directoryWithFiles, "ExcelPivotChart.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");

                sheet.CellValue(1, 1, "Region");
                sheet.CellValue(1, 2, "Product");
                sheet.CellValue(1, 3, "Sales");

                sheet.CellValue(2, 1, "East");
                sheet.CellValue(2, 2, "A");
                sheet.CellValue(2, 3, 10);

                sheet.CellValue(3, 1, "West");
                sheet.CellValue(3, 2, "A");
                sheet.CellValue(3, 3, 12);

                sheet.CellValue(4, 1, "East");
                sheet.CellValue(4, 2, "B");
                sheet.CellValue(4, 3, 7);

                sheet.AddPivotTable(
                    sourceRange: "A1:C4",
                    destinationCell: "E2",
                    name: "SalesPivot",
                    rowFields: new[] { "Region" },
                    columnFields: new[] { "Product" },
                    dataFields: new[] { new ExcelPivotDataField("Sales", DataConsolidateFunctionValues.Sum, "Total Sales") });

                var chart = sheet.AddPivotChartFromRange("SalesPivot", "A1:C4", row: 7, column: 1, title: "Sales Pivot");
                Assert.True(chart.IsPivotChart);
                Assert.Equal("SalesPivot", chart.PivotTableName);

                document.Save(false);
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                var worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First(part => part.DrawingsPart?.ChartParts.Any() == true);
                var chartPart = worksheetPart.DrawingsPart!.ChartParts.Single();
                var pivotSource = chartPart.ChartSpace!.GetFirstChild<C.PivotSource>();

                Assert.NotNull(pivotSource);
                Assert.Equal("SalesPivot", pivotSource!.GetFirstChild<C.PivotTableName>()!.Text);
                Assert.Equal(0U, pivotSource.GetFirstChild<C.FormatId>()!.Val!.Value);
            }

            using (var document = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }
    }
}
