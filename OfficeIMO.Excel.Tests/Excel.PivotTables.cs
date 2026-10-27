using System;
using System.Globalization;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Excel;
using Xunit;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Tests
{
    public partial class Excel
    {
        [Fact]
        public void Test_AddPivotTableAndReadBack()
        {
            var filePath = Path.Combine(_directoryWithFiles, "ExcelPivotTableBasic.xlsx");
            using (var document = ExcelDocument.Create(filePath))
            {
                var sheet = document.AddWorksheet("Data");
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
                sheet.AddPivotTable("A1:C4", "E2", options: new ExcelPivotTableOptions { Name = "SalesPivot", RowFields = new[] { "Region" }, ColumnFields = new[] { "Product" }, DataFields = new[] { new ExcelPivotDataField("Sales", ExcelPivotDataFunction.Sum, "Total Sales") }, PivotStyleName = "PivotStyleMedium9" });
                document.Save();
            }

            using (var document = ExcelDocument.Load(filePath))
            {
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
                Assert.Equal(ExcelPivotDataFunction.Sum, dataField.Function);
                Assert.Equal("Total Sales", dataField.DisplayName);
            }

            using (var document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly }))
            {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_PivotConditionalFormatting_TargetsPivotDataBody()
        {
            var filePath = Path.Combine(_directoryWithFiles, "ExcelPivotTableConditionalFormatting.xlsx");
            using (var document = ExcelDocument.Create(filePath))
            {
                var sheet = document.AddWorksheet("Data");
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
                sheet.AddPivotTable("A1:C4", "E2", options: new ExcelPivotTableOptions { Name = "SalesPivot", RowFields = new[] { "Region" }, ColumnFields = new[] { "Product" }, DataFields = new[] { new ExcelPivotDataField("Sales", ExcelPivotDataFunction.Sum, "Total Sales") } });
                Assert.Equal("F3:G3", sheet.GetPivotTableRange("SalesPivot", ExcelPivotRangeTarget.DataBody));
                sheet.AddPivotConditionalRule("SalesPivot", ExcelConditionalFormattingOperator.GreaterThan, "0");
                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false))
            {
                WorksheetPart worksheetPart = GetWorksheetPartByName(spreadsheet, "Data");
                ConditionalFormatting formatting = worksheetPart.Worksheet.Descendants<ConditionalFormatting>().Single();
                Assert.Equal("F3:G3", formatting.SequenceOfReferences!.InnerText);
                Assert.Equal(ConditionalFormatValues.CellIs, formatting.GetFirstChild<ConditionalFormattingRule>()!.Type!.Value);
            }
        }

        [Fact]
        public void Test_FluentPivotBuilder_CreatesPivotTable()
        {
            var filePath = Path.Combine(_directoryWithFiles, "ExcelPivotTableFluent.xlsx");
            using (var document = ExcelDocument.Create(filePath))
            {
                var sheet = document.AddWorksheet("Data");
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
                {
                    var pivotOptions1 = new ExcelPivotTableOptions();
                    pivotOptions1.RowFields.Add("Region");
                    pivotOptions1.ColumnFields.Add("Product");
                    pivotOptions1.DataFields.Add(new ExcelPivotDataField("Sales", ExcelPivotDataFunction.Sum, displayName: "Total Sales", numberFormat: "$#,##0"));
                    pivotOptions1.PivotStyleName = "PivotStyleMedium9";
                    pivotOptions1.Layout = ExcelPivotLayout.Tabular;
                    pivotOptions1.ShowRowGrandTotals = false;
                    pivotOptions1.ShowColumnGrandTotals = true;
                    pivotOptions1.RowHeaderCaption = "Rows";
                    pivotOptions1.ColumnHeaderCaption = "Products";
                    pivotOptions1.GrandTotalCaption = "Total";
                    pivotOptions1.Name = "SalesPivot";
                    sheet.AddPivotTable("A1:C4", "E2", pivotOptions1);
                }

                document.Save();
            }

            using (var document = ExcelDocument.Load(filePath))
            {
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
                Assert.Equal(ExcelPivotDataFunction.Sum, dataField.Function);
                Assert.Equal("Total Sales", dataField.DisplayName);
                Assert.True(dataField.NumberFormatId >= 164);
            }

            using (var document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly }))
            {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_PivotTableInteractionOptions_AreWrittenAndReadBack()
        {
            var filePath = Path.Combine(_directoryWithFiles, "ExcelPivotTableInteractionOptions.xlsx");
            using (var document = ExcelDocument.Create(filePath))
            {
                var sheet = document.AddWorksheet("Data");
                sheet.CellValue(1, 1, "Region");
                sheet.CellValue(1, 2, "Sales");
                sheet.CellValue(2, 1, "East");
                sheet.CellValue(2, 2, 10);
                sheet.CellValue(3, 1, "West");
                sheet.CellValue(3, 2, 12);
                {
                    var pivotOptions2 = new ExcelPivotTableOptions();
                    pivotOptions2.RowFields.Add("Region");
                    pivotOptions2.DataFields.Add(new ExcelPivotDataField("Sales", ExcelPivotDataFunction.Sum, displayName: "Total Sales"));
                    pivotOptions2.RefreshOnOpen = true;
                    pivotOptions2.SaveSourceData = false;
                    pivotOptions2.PreserveFormatting = false;
                    pivotOptions2.EnableDrill = false;
                    pivotOptions2.Name = "SalesPivot";
                    sheet.AddPivotTable("A1:B3", "E2", pivotOptions2);
                }

                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false))
            {
                var pivotPart = spreadsheet.WorkbookPart!.WorksheetParts.SelectMany(part => part.PivotTableParts).Single();
                var cachePart = spreadsheet.WorkbookPart!.PivotTableCacheDefinitionParts.Single();
                Assert.True(cachePart.PivotCacheDefinition!.RefreshOnLoad!.Value);
                Assert.False(cachePart.PivotCacheDefinition!.SaveData!.Value);
                Assert.False(pivotPart.PivotTableDefinition!.PreserveFormatting!.Value);
                Assert.False(pivotPart.PivotTableDefinition!.EnableDrill!.Value);
            }

            using (var document = ExcelDocument.Load(filePath))
            {
                var pivot = document.GetPivotTables().Single();
                Assert.True(pivot.RefreshOnOpen);
                Assert.False(pivot.SaveSourceData);
                Assert.False(pivot.PreserveFormatting);
                Assert.False(pivot.EnableDrill);
            }
        }

        [Fact]
        public void Test_PivotTableCalculatedFields_DoNotPersistSourceRowsWithoutOptIn()
        {
            var filePath = Path.Combine(_directoryWithFiles, "ExcelPivotTableNoSourceCacheByDefault.xlsx");
            using (var document = ExcelDocument.Create(filePath))
            {
                var sheet = document.AddWorksheet("Data");
                sheet.CellValue(1, 1, "Region");
                sheet.CellValue(1, 2, "Sales");
                sheet.CellValue(2, 1, "Confidential East");
                sheet.CellValue(2, 2, 10);
                sheet.CellValue(3, 1, "Confidential West");
                sheet.CellValue(3, 2, 12);
                sheet.AddPivotTable("A1:B3", "E2", options: new ExcelPivotTableOptions { Name = "SalesPivot", RowFields = new[] { "Region" }, DataFields = new[] { new ExcelPivotDataField("Sales", ExcelPivotDataFunction.Sum, "Total Sales") }, CalculatedFields = new[] { new ExcelPivotCalculatedField("DoubleSales", "'Sales' * 2") } });
                document.Save();
            }

            using var spreadsheet = SpreadsheetDocument.Open(filePath, false);
            var cachePart = spreadsheet.WorkbookPart!.PivotTableCacheDefinitionParts.Single();
            Assert.False(cachePart.PivotCacheDefinition!.SaveData!.Value);
            Assert.True(cachePart.PivotCacheDefinition.RefreshOnLoad!.Value);
            var recordsPart = Assert.Single(cachePart.GetPartsOfType<PivotTableCacheRecordsPart>());
            Assert.Equal(0U, recordsPart.PivotCacheRecords!.Count!.Value);
        }

        [Fact]
        public void Test_FluentPivotBuilder_AppliesItemAndPageFilterHelpers()
        {
            var filePath = Path.Combine(_directoryWithFiles, "ExcelPivotTableFluentItemFilters.xlsx");
            using (var document = ExcelDocument.Create(filePath))
            {
                var sheet = document.AddWorksheet("Data");
                sheet.CellValue(1, 1, "Region");
                sheet.CellValue(1, 2, "Product");
                sheet.CellValue(1, 3, "Sales");
                sheet.CellValue(2, 1, "East");
                sheet.CellValue(2, 2, "A");
                sheet.CellValue(2, 3, 10);
                sheet.CellValue(3, 1, "West");
                sheet.CellValue(3, 2, "A");
                sheet.CellValue(3, 3, 12);
                sheet.CellValue(4, 1, "North");
                sheet.CellValue(4, 2, "B");
                sheet.CellValue(4, 3, 7);
                {
                    var pivotOptions3 = new ExcelPivotTableOptions();
                    pivotOptions3.RowFields.Add("Region");
                    pivotOptions3.PageFields.Add("Product");
                    pivotOptions3.DataFields.Add(new ExcelPivotDataField("Sales", ExcelPivotDataFunction.Sum, displayName: "Total Sales"));
                    pivotOptions3.FieldOptions.Add(new ExcelPivotFieldOptions("Region", sortType: ExcelPivotFieldSort.Descending, defaultSubtotal: false, subtotalTop: true, compact: false, outline: true, insertBlankRow: true, insertPageBreak: true, showDropDowns: false, multipleItemSelectionAllowed: true, includeNewItemsInFilter: true, numberFormat: "@", subtotalCaption: "Region subtotal", visibleItems: new[] { "East", "North" }));
                    pivotOptions3.FieldOptions.Add(new ExcelPivotFieldOptions("Product", selectedItem: "B"));
                    pivotOptions3.Name = "SalesPivot";
                    sheet.AddPivotTable("A1:C4", "E2", pivotOptions3);
                }

                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false))
            {
                var pivotPart = spreadsheet.WorkbookPart!.WorksheetParts.SelectMany(part => part.PivotTableParts).Single();
                var pivotDefinition = pivotPart.PivotTableDefinition!;
                var regionField = pivotDefinition.PivotFields!.Elements<PivotField>().ElementAt(0);
                Assert.Equal(FieldSortValues.Descending, regionField.SortType!.Value);
                Assert.False(regionField.DefaultSubtotal!.Value);
                Assert.True(regionField.SubtotalTop!.Value);
                Assert.True(regionField.InsertBlankRow!.Value);
                Assert.True(regionField.InsertPageBreak!.Value);
                Assert.False(regionField.Compact!.Value);
                Assert.True(regionField.Outline!.Value);
                Assert.False(regionField.ShowDropDowns!.Value);
                Assert.True(regionField.MultipleItemSelectionAllowed!.Value);
                Assert.True(regionField.IncludeNewItemsInFilter!.Value);
                Assert.True(regionField.NumberFormatId!.Value >= 164);
                Assert.Equal("Region subtotal", regionField.SubtotalCaption!.Value);
                var fieldNumberFormat = spreadsheet.WorkbookPart!.WorkbookStylesPart!.Stylesheet!.NumberingFormats!.Elements<NumberingFormat>().Single(format => format.NumberFormatId!.Value == regionField.NumberFormatId!.Value);
                Assert.Equal("@", fieldNumberFormat.FormatCode!.Value);
                var regionItems = regionField.Items!.Elements<Item>().ToList();
                Assert.Equal(4, regionItems.Count);
                Assert.False(regionItems[0].Hidden?.Value ?? false);
                Assert.True(regionItems[1].Hidden!.Value);
                Assert.False(regionItems[2].Hidden?.Value ?? false);
                Assert.Equal(ItemValues.Default, regionItems[3].ItemType!.Value);
                Assert.False(regionField.ShowAll!.Value);
                var pageField = pivotDefinition.PageFields!.Elements<PageField>().Single();
                Assert.Equal(1, pageField.Field!.Value);
                Assert.Equal(1U, pageField.Item!.Value);
            }

            using (var document = ExcelDocument.Load(filePath))
            {
                var pivot = document.GetPivotTables().Single();
                Assert.Contains("Product", pivot.PageFields);
                var region = pivot.Fields.Single(field => field.FieldName == "Region");
                Assert.Equal(ExcelPivotFieldSort.Descending, region.SortType);
                Assert.False(region.DefaultSubtotal);
                Assert.True(region.SubtotalTop);
                Assert.True(region.InsertBlankRow);
                Assert.True(region.InsertPageBreak);
                Assert.False(region.Compact);
                Assert.True(region.Outline);
                Assert.False(region.ShowDropDowns);
                Assert.True(region.MultipleItemSelectionAllowed);
                Assert.True(region.IncludeNewItemsInFilter);
                Assert.True(region.NumberFormatId >= 164);
                Assert.Equal("@", region.NumberFormatCode);
                Assert.Equal("Region subtotal", region.SubtotalCaption);
                Assert.Contains("West", region.HiddenItems);
                Assert.DoesNotContain("East", region.HiddenItems);
                Assert.DoesNotContain("North", region.HiddenItems);
                Assert.Equal(new[] { "East", "North" }, region.VisibleItems);
                var product = pivot.Fields.Single(field => field.FieldName == "Product");
                Assert.Equal("B", product.SelectedItem);
            }

            using (var document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly }))
            {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_FluentPivotBuilder_ReadsBuiltInNumberFormatCodes()
        {
            var filePath = Path.Combine(_directoryWithFiles, "ExcelPivotTableBuiltInNumberFormats.xlsx");
            using (var document = ExcelDocument.Create(filePath))
            {
                var sheet = document.AddWorksheet("Data");
                sheet.CellValue(1, 1, "Region");
                sheet.CellValue(1, 2, "Sales");
                sheet.CellValue(2, 1, "East");
                sheet.CellValue(2, 2, 10.25);
                sheet.CellValue(3, 1, "West");
                sheet.CellValue(3, 2, 12.5);
                {
                    var pivotOptions4 = new ExcelPivotTableOptions();
                    pivotOptions4.RowFields.Add("Region");
                    pivotOptions4.DataFields.Add(new ExcelPivotDataField("Sales", ExcelPivotDataFunction.Sum, "Total Sales", numberFormatId: 4));
                    pivotOptions4.FieldOptions.Add(new ExcelPivotFieldOptions("Region", numberFormatId: 49));
                    pivotOptions4.Name = "SalesPivot";
                    sheet.AddPivotTable("A1:B3", "D2", pivotOptions4);
                }

                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false))
            {
                var pivotPart = spreadsheet.WorkbookPart!.WorksheetParts.SelectMany(part => part.PivotTableParts).Single();
                var pivotDefinition = pivotPart.PivotTableDefinition!;
                var regionField = pivotDefinition.PivotFields!.Elements<PivotField>().ElementAt(0);
                Assert.Equal(49U, regionField.NumberFormatId!.Value);
                var dataField = pivotDefinition.DataFields!.Elements<DataField>().Single();
                Assert.Equal(4U, dataField.NumberFormatId!.Value);
            }

            using (var document = ExcelDocument.Load(filePath))
            {
                var pivot = document.GetPivotTables().Single();
                var region = pivot.Fields.Single(field => field.FieldName == "Region");
                Assert.Equal(49U, region.NumberFormatId);
                Assert.Equal("@", region.NumberFormatCode);
                var dataField = pivot.DataFields.Single();
                Assert.Equal(4U, dataField.NumberFormatId);
                Assert.Equal("#,##0.00", dataField.NumberFormatCode);
            }

            using (var document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly }))
            {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_AddPivotTable_AppliesFieldOptionsAndNumberFormats()
        {
            var filePath = Path.Combine(_directoryWithFiles, "ExcelPivotTableFieldOptions.xlsx");
            using (var document = ExcelDocument.Create(filePath))
            {
                var sheet = document.AddWorksheet("Data");
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
                sheet.AddPivotTable("A1:C5", "E2", options: new ExcelPivotTableOptions { Name = "SalesPivot", RowFields = new[] { "Region" }, PageFields = new[] { "Product" }, DataFields = new[] { new ExcelPivotDataField("Sales", ExcelPivotDataFunction.Sum, "Total Sales", numberFormat: "$#,##0.00") }, FieldOptions = new[] { new ExcelPivotFieldOptions("Region", sortType: ExcelPivotFieldSort.Ascending, defaultSubtotal: false, subtotalTop: true, insertBlankRow: true, insertPageBreak: true, compact: false, outline: true, showDropDowns: true, multipleItemSelectionAllowed: true, includeNewItemsInFilter: true, subtotalCaption: "Region subtotal", hiddenItems: new[] { "West" }), new ExcelPivotFieldOptions("Product", selectedItem: "A") }, RowHeaderCaption = "Rows", ColumnHeaderCaption = "Columns", GrandTotalCaption = "Grand total", MissingCaption = "(missing)", ErrorCaption = "(error)", ShowDataDropDown = false, ShowDropZones = true, ShowDataTips = true, ShowMemberPropertyTips = true, FieldListSortAscending = true, CustomListSort = false });
                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false))
            {
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
                Assert.Equal(4, regionItems.Count);
                Assert.False(regionItems[0].Hidden?.Value ?? false);
                Assert.True(regionItems[1].Hidden!.Value);
                Assert.False(regionItems[2].Hidden?.Value ?? false);
                Assert.Equal(ItemValues.Default, regionItems[3].ItemType!.Value);
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
                var numberFormat = workbookPart.WorkbookStylesPart!.Stylesheet!.NumberingFormats!.Elements<NumberingFormat>().Single(format => format.NumberFormatId!.Value == numberFormatId);
                Assert.Equal("$#,##0.00", numberFormat.FormatCode!.Value);
            }

            using (var document = ExcelDocument.Load(filePath))
            {
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
                Assert.Equal(ExcelPivotFieldSort.Ascending, region.SortType);
                Assert.False(region.DefaultSubtotal);
                Assert.True(region.InsertPageBreak);
                Assert.False(region.Compact);
                Assert.True(region.Outline);
                Assert.Contains("West", region.HiddenItems);
                Assert.Equal(new[] { "East", string.Empty }, region.VisibleItems);
                var product = pivot.Fields.Single(field => field.FieldName == "Product");
                Assert.Equal("A", product.SelectedItem);
                var dataField = pivot.DataFields.Single();
                Assert.True(dataField.NumberFormatId >= 164);
                Assert.Equal("$#,##0.00", dataField.NumberFormatCode);
            }

            using (var document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly }))
            {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_AddPivotTable_AppliesShowValuesAsDataFieldOptions()
        {
            var filePath = Path.Combine(_directoryWithFiles, "ExcelPivotTableShowValuesAs.xlsx");
            using (var document = ExcelDocument.Create(filePath))
            {
                var sheet = document.AddWorksheet("Data");
                sheet.CellValue(1, 1, "Region");
                sheet.CellValue(1, 2, "Product");
                sheet.CellValue(1, 3, "Sales");
                sheet.CellValue(2, 1, "East");
                sheet.CellValue(2, 2, "A");
                sheet.CellValue(2, 3, 10);
                sheet.CellValue(3, 1, "West");
                sheet.CellValue(3, 2, "A");
                sheet.CellValue(3, 3, 20);
                sheet.CellValue(4, 1, "East");
                sheet.CellValue(4, 2, "B");
                sheet.CellValue(4, 3, 30);
                {
                    var pivotOptions5 = new ExcelPivotTableOptions();
                    pivotOptions5.RowFields.Add("Region");
                    pivotOptions5.DataFields.Add(new ExcelPivotDataField("Sales", ExcelPivotDataFunction.Sum, displayName: "% Total Sales", showDataAs: ExcelPivotShowDataAs.PercentOfTotal, numberFormat: "0.0%"));
                    pivotOptions5.Name = "SalesPivot";
                    sheet.AddPivotTable("A1:C4", "E2", pivotOptions5);
                }

                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false))
            {
                var pivotPart = spreadsheet.WorkbookPart!.WorksheetParts.SelectMany(part => part.PivotTableParts).Single();
                var dataField = pivotPart.PivotTableDefinition!.DataFields!.Elements<DataField>().Single();
                Assert.Equal("% Total Sales", dataField.Name!.Value);
                Assert.Equal(ShowDataAsValues.PercentOfTotal, dataField.ShowDataAs!.Value);
                Assert.True(dataField.NumberFormatId!.Value >= 164);
            }

            using (var document = ExcelDocument.Load(filePath))
            {
                var dataField = document.GetPivotTables().Single().DataFields.Single();
                Assert.Equal("% Total Sales", dataField.DisplayName);
                Assert.Equal(ExcelPivotShowDataAs.PercentOfTotal, dataField.ShowDataAs);
                Assert.True(dataField.NumberFormatId >= 164);
            }

            using (var document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly }))
            {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_AddPivotTable_AppliesLabelAndValueFilters()
        {
            var filePath = Path.Combine(_directoryWithFiles, "ExcelPivotTableFilters.xlsx");
            using (var document = ExcelDocument.Create(filePath))
            {
                var sheet = document.AddWorksheet("Data");
                sheet.CellValue(1, 1, "Region");
                sheet.CellValue(1, 2, "Product");
                sheet.CellValue(1, 3, "Sales");
                sheet.CellValue(2, 1, "East");
                sheet.CellValue(2, 2, "A");
                sheet.CellValue(2, 3, 10);
                sheet.CellValue(3, 1, "West");
                sheet.CellValue(3, 2, "A");
                sheet.CellValue(3, 3, 20);
                sheet.CellValue(4, 1, "East");
                sheet.CellValue(4, 2, "B");
                sheet.CellValue(4, 3, 30);
                {
                    var pivotOptions6 = new ExcelPivotTableOptions();
                    pivotOptions6.RowFields.Add("Region");
                    pivotOptions6.ColumnFields.Add("Product");
                    pivotOptions6.DataFields.Add(new ExcelPivotDataField("Sales", ExcelPivotDataFunction.Sum, displayName: "Total Sales"));
                    pivotOptions6.PivotFilters.Add(ExcelPivotFilter.LabelContains("Region", "Ea", name: "Region contains Ea"));
                    pivotOptions6.PivotFilters.Add(ExcelPivotFilter.ValueGreaterThan("Region", "Total Sales", 15, name: "Sales above 15"));
                    pivotOptions6.Name = "SalesPivot";
                    sheet.AddPivotTable("A1:C4", "E2", pivotOptions6);
                }

                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false))
            {
                var pivotPart = spreadsheet.WorkbookPart!.WorksheetParts.SelectMany(part => part.PivotTableParts).Single();
                var filters = pivotPart.PivotTableDefinition!.PivotFilters!.Elements<PivotFilter>().ToList();
                Assert.Equal(2, filters.Count);
                Assert.Equal(PivotFilterValues.CaptionContains, filters[0].Type!.Value);
                Assert.Equal(0U, filters[0].Field!.Value);
                Assert.Equal("Ea", filters[0].StringValue1!.Value);
                Assert.Equal("Region contains Ea", filters[0].Name!.Value);
                var labelCustomFilter = Assert.Single(filters[0].AutoFilter!.Elements<FilterColumn>().Single().GetFirstChild<CustomFilters>()!.Elements<CustomFilter>());
                Assert.Equal(FilterOperatorValues.Equal, labelCustomFilter.Operator!.Value);
                Assert.Equal("*Ea*", labelCustomFilter.Val!.Value);
                Assert.Equal(PivotFilterValues.ValueGreaterThan, filters[1].Type!.Value);
                Assert.Equal(0U, filters[1].Field!.Value);
                Assert.Equal(0U, filters[1].MeasureField!.Value);
                Assert.Equal("15", filters[1].StringValue1!.Value);
                Assert.Equal("Sales above 15", filters[1].Name!.Value);
                var valueCustomFilter = Assert.Single(filters[1].AutoFilter!.Elements<FilterColumn>().Single().GetFirstChild<CustomFilters>()!.Elements<CustomFilter>());
                Assert.Equal(FilterOperatorValues.GreaterThan, valueCustomFilter.Operator!.Value);
                Assert.Equal("15", valueCustomFilter.Val!.Value);
            }

            using (var document = ExcelDocument.Load(filePath))
            {
                var pivot = document.GetPivotTables().Single();
                Assert.Equal(2, pivot.Filters.Count);
                Assert.Equal("Region", pivot.Filters[0].FieldName);
                Assert.Equal(ExcelPivotFilterType.CaptionContains, pivot.Filters[0].Type);
                Assert.Equal("Ea", pivot.Filters[0].Value1);
                Assert.Equal("Total Sales", pivot.Filters[1].DataFieldName);
                Assert.Equal(ExcelPivotFilterType.ValueGreaterThan, pivot.Filters[1].Type);
            }

            using (var document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly }))
            {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_AddPivotTable_AppliesExpandedFilterVariants()
        {
            var filePath = Path.Combine(_directoryWithFiles, "ExcelPivotTableExpandedFilters.xlsx");
            using (var document = ExcelDocument.Create(filePath))
            {
                var sheet = document.AddWorksheet("Data");
                sheet.CellValue(1, 1, "Region");
                sheet.CellValue(1, 2, "Product");
                sheet.CellValue(1, 3, "Sales");
                sheet.CellValue(2, 1, "East");
                sheet.CellValue(2, 2, "A");
                sheet.CellValue(2, 3, 10);
                sheet.CellValue(3, 1, "West");
                sheet.CellValue(3, 2, "A");
                sheet.CellValue(3, 3, 20);
                sheet.CellValue(4, 1, "North");
                sheet.CellValue(4, 2, "B");
                sheet.CellValue(4, 3, 30);
                sheet.CellValue(5, 1, "South");
                sheet.CellValue(5, 2, "B");
                sheet.CellValue(5, 3, 40);
                {
                    var pivotOptions7 = new ExcelPivotTableOptions();
                    pivotOptions7.RowFields.Add("Region");
                    pivotOptions7.DataFields.Add(new ExcelPivotDataField("Sales", ExcelPivotDataFunction.Sum, displayName: "Total Sales"));
                    pivotOptions7.PivotFilters.Add(ExcelPivotFilter.LabelNotContains("Region", "East", name: "Exclude East"));
                    pivotOptions7.PivotFilters.Add(ExcelPivotFilter.LabelNotBetween("Region", "M", "S", name: "Outside M-S"));
                    pivotOptions7.PivotFilters.Add(ExcelPivotFilter.ValueLessThanOrEqual("Region", "Total Sales", 30, name: "Sales <= 30"));
                    pivotOptions7.PivotFilters.Add(ExcelPivotFilter.ValueNotBetween("Region", "Total Sales", 15, 35, name: "Outside sales band"));
                    pivotOptions7.Name = "SalesPivot";
                    sheet.AddPivotTable("A1:C5", "E2", pivotOptions7);
                }

                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false))
            {
                var pivotPart = spreadsheet.WorkbookPart!.WorksheetParts.SelectMany(part => part.PivotTableParts).Single();
                var filters = pivotPart.PivotTableDefinition!.PivotFilters!.Elements<PivotFilter>().ToList();
                Assert.Equal(4, filters.Count);
                Assert.Equal(PivotFilterValues.CaptionNotContains, filters[0].Type!.Value);
                var notContains = Assert.Single(filters[0].AutoFilter!.Elements<FilterColumn>().Single().GetFirstChild<CustomFilters>()!.Elements<CustomFilter>());
                Assert.Equal(FilterOperatorValues.NotEqual, notContains.Operator!.Value);
                Assert.Equal("*East*", notContains.Val!.Value);
                Assert.Equal(PivotFilterValues.CaptionNotBetween, filters[1].Type!.Value);
                var labelBand = filters[1].AutoFilter!.Elements<FilterColumn>().Single().GetFirstChild<CustomFilters>()!;
                Assert.False(labelBand.And!.Value);
                var labelBandFilters = labelBand.Elements<CustomFilter>().ToList();
                Assert.Equal(FilterOperatorValues.LessThan, labelBandFilters[0].Operator!.Value);
                Assert.Equal("M", labelBandFilters[0].Val!.Value);
                Assert.Equal(FilterOperatorValues.GreaterThan, labelBandFilters[1].Operator!.Value);
                Assert.Equal("S", labelBandFilters[1].Val!.Value);
                Assert.Equal(PivotFilterValues.ValueLessThanOrEqual, filters[2].Type!.Value);
                Assert.Equal(0U, filters[2].MeasureField!.Value);
                var lessThanOrEqual = Assert.Single(filters[2].AutoFilter!.Elements<FilterColumn>().Single().GetFirstChild<CustomFilters>()!.Elements<CustomFilter>());
                Assert.Equal(FilterOperatorValues.LessThanOrEqual, lessThanOrEqual.Operator!.Value);
                Assert.Equal("30", lessThanOrEqual.Val!.Value);
                Assert.Equal(PivotFilterValues.ValueNotBetween, filters[3].Type!.Value);
                Assert.Equal(0U, filters[3].MeasureField!.Value);
                var valueBand = filters[3].AutoFilter!.Elements<FilterColumn>().Single().GetFirstChild<CustomFilters>()!;
                Assert.False(valueBand.And!.Value);
                var valueBandFilters = valueBand.Elements<CustomFilter>().ToList();
                Assert.Equal(FilterOperatorValues.LessThan, valueBandFilters[0].Operator!.Value);
                Assert.Equal("15", valueBandFilters[0].Val!.Value);
                Assert.Equal(FilterOperatorValues.GreaterThan, valueBandFilters[1].Operator!.Value);
                Assert.Equal("35", valueBandFilters[1].Val!.Value);
            }

            using (var document = ExcelDocument.Load(filePath))
            {
                var pivot = document.GetPivotTables().Single();
                Assert.Equal(4, pivot.Filters.Count);
                Assert.Equal(ExcelPivotFilterType.CaptionNotContains, pivot.Filters[0].Type);
                Assert.Equal("East", pivot.Filters[0].Value1);
                Assert.Equal(ExcelPivotFilterType.CaptionNotBetween, pivot.Filters[1].Type);
                Assert.Equal("M", pivot.Filters[1].Value1);
                Assert.Equal("S", pivot.Filters[1].Value2);
                Assert.Equal(ExcelPivotFilterType.ValueLessThanOrEqual, pivot.Filters[2].Type);
                Assert.Equal("Total Sales", pivot.Filters[2].DataFieldName);
                Assert.Equal(ExcelPivotFilterType.ValueNotBetween, pivot.Filters[3].Type);
                Assert.Equal("15", pivot.Filters[3].Value1);
                Assert.Equal("35", pivot.Filters[3].Value2);
            }

            using (var document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly }))
            {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_AddPivotTable_AppliesDynamicDateFilters()
        {
            var filePath = Path.Combine(_directoryWithFiles, "ExcelPivotTableDynamicDateFilters.xlsx");
            using (var document = ExcelDocument.Create(filePath))
            {
                var sheet = document.AddWorksheet("Data");
                sheet.CellValue(1, 1, "OrderDate");
                sheet.CellValue(1, 2, "Product");
                sheet.CellValue(1, 3, "Sales");
                sheet.CellValue(2, 1, new DateTime(2026, 1, 10));
                sheet.CellValue(2, 2, "A");
                sheet.CellValue(2, 3, 10);
                sheet.CellValue(3, 1, new DateTime(2026, 2, 15));
                sheet.CellValue(3, 2, "A");
                sheet.CellValue(3, 3, 20);
                sheet.CellValue(4, 1, new DateTime(2026, 3, 20));
                sheet.CellValue(4, 2, "B");
                sheet.CellValue(4, 3, 30);
                sheet.CellValue(5, 1, new DateTime(2026, 4, 25));
                sheet.CellValue(5, 2, "B");
                sheet.CellValue(5, 3, 40);
                {
                    var pivotOptions8 = new ExcelPivotTableOptions();
                    pivotOptions8.RowFields.Add("OrderDate");
                    pivotOptions8.DataFields.Add(new ExcelPivotDataField("Sales", ExcelPivotDataFunction.Sum, displayName: "Total Sales"));
                    pivotOptions8.PivotFilters.Add(ExcelPivotFilter.DateThisMonth("OrderDate", name: "This month"));
                    pivotOptions8.PivotFilters.Add(ExcelPivotFilter.DateYearToDate("OrderDate", name: "YTD"));
                    pivotOptions8.PivotFilters.Add(ExcelPivotFilter.DateQuarter("OrderDate", 1, name: "Q1"));
                    pivotOptions8.PivotFilters.Add(ExcelPivotFilter.DateMonth("OrderDate", 2, name: "February"));
                    pivotOptions8.Name = "SalesPivot";
                    sheet.AddPivotTable("A1:C5", "E2", pivotOptions8);
                }

                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false))
            {
                var pivotPart = spreadsheet.WorkbookPart!.WorksheetParts.SelectMany(part => part.PivotTableParts).Single();
                var filters = pivotPart.PivotTableDefinition!.PivotFilters!.Elements<PivotFilter>().ToList();
                Assert.Equal(4, filters.Count);
                Assert.Equal(PivotFilterValues.ThisMonth, filters[0].Type!.Value);
                Assert.Equal(DynamicFilterValues.ThisMonth, filters[0].AutoFilter!.Elements<FilterColumn>().Single().GetFirstChild<DynamicFilter>()!.Type!.Value);
                Assert.Null(filters[0].StringValue1);
                Assert.Equal(PivotFilterValues.YearToDate, filters[1].Type!.Value);
                Assert.Equal(DynamicFilterValues.YearToDate, filters[1].AutoFilter!.Elements<FilterColumn>().Single().GetFirstChild<DynamicFilter>()!.Type!.Value);
                Assert.Equal(PivotFilterValues.Quarter1, filters[2].Type!.Value);
                Assert.Equal(DynamicFilterValues.Quarter1, filters[2].AutoFilter!.Elements<FilterColumn>().Single().GetFirstChild<DynamicFilter>()!.Type!.Value);
                Assert.Equal(PivotFilterValues.February, filters[3].Type!.Value);
                Assert.Equal(DynamicFilterValues.February, filters[3].AutoFilter!.Elements<FilterColumn>().Single().GetFirstChild<DynamicFilter>()!.Type!.Value);
            }

            using (var document = ExcelDocument.Load(filePath))
            {
                var pivot = document.GetPivotTables().Single();
                Assert.Equal(4, pivot.Filters.Count);
                Assert.Equal(ExcelPivotFilterType.ThisMonth, pivot.Filters[0].Type);
                Assert.Equal("This month", pivot.Filters[0].Name);
                Assert.Null(pivot.Filters[0].Value1);
                Assert.Equal(ExcelPivotFilterType.YearToDate, pivot.Filters[1].Type);
                Assert.Equal(ExcelPivotFilterType.Quarter1, pivot.Filters[2].Type);
                Assert.Equal(ExcelPivotFilterType.February, pivot.Filters[3].Type);
            }

            using (var document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly }))
            {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_AddPivotTable_AppliesFixedDateFilters()
        {
            var filePath = Path.Combine(_directoryWithFiles, "ExcelPivotTableFixedDateFilters.xlsx");
            string DateSerial(DateTime value) => value.ToOADate().ToString("G17", CultureInfo.InvariantCulture);
            using (var document = ExcelDocument.Create(filePath))
            {
                var sheet = document.AddWorksheet("Data");
                sheet.CellValue(1, 1, "OrderDate");
                sheet.CellValue(1, 2, "Product");
                sheet.CellValue(1, 3, "Sales");
                sheet.CellValue(2, 1, new DateTime(2026, 1, 10));
                sheet.CellValue(2, 2, "A");
                sheet.CellValue(2, 3, 10);
                sheet.CellValue(3, 1, new DateTime(2026, 2, 15));
                sheet.CellValue(3, 2, "A");
                sheet.CellValue(3, 3, 20);
                sheet.CellValue(4, 1, new DateTime(2026, 3, 20));
                sheet.CellValue(4, 2, "B");
                sheet.CellValue(4, 3, 30);
                sheet.CellValue(5, 1, new DateTime(2026, 4, 25));
                sheet.CellValue(5, 2, "B");
                sheet.CellValue(5, 3, 40);
                {
                    var pivotOptions9 = new ExcelPivotTableOptions();
                    pivotOptions9.RowFields.Add("OrderDate");
                    pivotOptions9.DataFields.Add(new ExcelPivotDataField("Sales", ExcelPivotDataFunction.Sum, displayName: "Total Sales"));
                    pivotOptions9.PivotFilters.Add(ExcelPivotFilter.DateNewerThanOrEqual("OrderDate", new DateTime(2026, 2, 1), name: "On/after Feb"));
                    pivotOptions9.PivotFilters.Add(ExcelPivotFilter.DateOlderThan("OrderDate", new DateTime(2026, 4, 1), name: "Before Apr"));
                    pivotOptions9.PivotFilters.Add(ExcelPivotFilter.DateBetween("OrderDate", new DateTime(2026, 1, 1), new DateTime(2026, 3, 31), name: "Q1 band"));
                    pivotOptions9.PivotFilters.Add(ExcelPivotFilter.DateNotBetween("OrderDate", new DateTime(2026, 2, 1), new DateTime(2026, 2, 28), name: "Not Feb"));
                    pivotOptions9.Name = "SalesPivot";
                    sheet.AddPivotTable("A1:C5", "E2", pivotOptions9);
                }

                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false))
            {
                var pivotPart = spreadsheet.WorkbookPart!.WorksheetParts.SelectMany(part => part.PivotTableParts).Single();
                var filters = pivotPart.PivotTableDefinition!.PivotFilters!.Elements<PivotFilter>().ToList();
                Assert.Equal(4, filters.Count);
                Assert.Equal(PivotFilterValues.DateNewerThanOrEqual, filters[0].Type!.Value);
                Assert.Equal(DateSerial(new DateTime(2026, 2, 1)), filters[0].StringValue1!.Value);
                var newer = Assert.Single(filters[0].AutoFilter!.Elements<FilterColumn>().Single().GetFirstChild<CustomFilters>()!.Elements<CustomFilter>());
                Assert.Equal(FilterOperatorValues.GreaterThanOrEqual, newer.Operator!.Value);
                Assert.Equal(PivotFilterValues.DateOlderThan, filters[1].Type!.Value);
                Assert.Equal(DateSerial(new DateTime(2026, 4, 1)), filters[1].StringValue1!.Value);
                var older = Assert.Single(filters[1].AutoFilter!.Elements<FilterColumn>().Single().GetFirstChild<CustomFilters>()!.Elements<CustomFilter>());
                Assert.Equal(FilterOperatorValues.LessThan, older.Operator!.Value);
                Assert.Equal(PivotFilterValues.DateBetween, filters[2].Type!.Value);
                var between = filters[2].AutoFilter!.Elements<FilterColumn>().Single().GetFirstChild<CustomFilters>()!;
                Assert.True(between.And!.Value);
                var betweenFilters = between.Elements<CustomFilter>().ToList();
                Assert.Equal(DateSerial(new DateTime(2026, 1, 1)), betweenFilters[0].Val!.Value);
                Assert.Equal(DateSerial(new DateTime(2026, 3, 31)), betweenFilters[1].Val!.Value);
                Assert.Equal(PivotFilterValues.DateNotBetween, filters[3].Type!.Value);
                var notBetween = filters[3].AutoFilter!.Elements<FilterColumn>().Single().GetFirstChild<CustomFilters>()!;
                Assert.False(notBetween.And!.Value);
            }

            using (var document = ExcelDocument.Load(filePath))
            {
                var pivot = document.GetPivotTables().Single();
                Assert.Equal(4, pivot.Filters.Count);
                Assert.Equal(ExcelPivotFilterType.DateNewerThanOrEqual, pivot.Filters[0].Type);
                Assert.Equal(DateSerial(new DateTime(2026, 2, 1)), pivot.Filters[0].Value1);
                Assert.Equal(ExcelPivotFilterType.DateOlderThan, pivot.Filters[1].Type);
                Assert.Equal(ExcelPivotFilterType.DateBetween, pivot.Filters[2].Type);
                Assert.Equal(DateSerial(new DateTime(2026, 3, 31)), pivot.Filters[2].Value2);
                Assert.Equal(ExcelPivotFilterType.DateNotBetween, pivot.Filters[3].Type);
            }

            using (var document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly }))
            {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_AddPivotTable_AppliesTopBottomFilters()
        {
            var filePath = Path.Combine(_directoryWithFiles, "ExcelPivotTableTopBottomFilters.xlsx");
            using (var document = ExcelDocument.Create(filePath))
            {
                var sheet = document.AddWorksheet("Data");
                sheet.CellValue(1, 1, "Region");
                sheet.CellValue(1, 2, "Product");
                sheet.CellValue(1, 3, "Sales");
                sheet.CellValue(2, 1, "East");
                sheet.CellValue(2, 2, "A");
                sheet.CellValue(2, 3, 10);
                sheet.CellValue(3, 1, "West");
                sheet.CellValue(3, 2, "A");
                sheet.CellValue(3, 3, 20);
                sheet.CellValue(4, 1, "North");
                sheet.CellValue(4, 2, "B");
                sheet.CellValue(4, 3, 30);
                sheet.CellValue(5, 1, "South");
                sheet.CellValue(5, 2, "B");
                sheet.CellValue(5, 3, 40);
                {
                    var pivotOptions10 = new ExcelPivotTableOptions();
                    pivotOptions10.RowFields.Add("Region");
                    pivotOptions10.DataFields.Add(new ExcelPivotDataField("Sales", ExcelPivotDataFunction.Sum, displayName: "Total Sales"));
                    pivotOptions10.PivotFilters.Add(ExcelPivotFilter.TopCount("Region", "Total Sales", 2, name: "Top 2 regions"));
                    pivotOptions10.PivotFilters.Add(ExcelPivotFilter.BottomPercent("Region", "Total Sales", 25, name: "Bottom quarter"));
                    pivotOptions10.PivotFilters.Add(ExcelPivotFilter.TopSum("Region", "Total Sales", 50, name: "Top 50 sum"));
                    pivotOptions10.PivotFilters.Add(ExcelPivotFilter.BottomSum("Region", "Total Sales", 30.5, name: "Bottom 30.5 sum"));
                    pivotOptions10.Name = "SalesPivot";
                    sheet.AddPivotTable("A1:C5", "E2", pivotOptions10);
                }

                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false))
            {
                var pivotPart = spreadsheet.WorkbookPart!.WorksheetParts.SelectMany(part => part.PivotTableParts).Single();
                var filters = pivotPart.PivotTableDefinition!.PivotFilters!.Elements<PivotFilter>().ToList();
                Assert.Equal(4, filters.Count);
                Assert.Equal(PivotFilterValues.Count, filters[0].Type!.Value);
                Assert.Equal("2", filters[0].StringValue1!.Value);
                Assert.Equal(0U, filters[0].MeasureField!.Value);
                var topCount = filters[0].AutoFilter!.Elements<FilterColumn>().Single().GetFirstChild<Top10>()!;
                Assert.True(topCount.Top!.Value);
                Assert.False(topCount.Percent!.Value);
                Assert.Equal(2D, topCount.Val!.Value);
                Assert.Equal(PivotFilterValues.Percent, filters[1].Type!.Value);
                Assert.Equal("25", filters[1].StringValue1!.Value);
                Assert.Equal(0U, filters[1].MeasureField!.Value);
                var bottomPercent = filters[1].AutoFilter!.Elements<FilterColumn>().Single().GetFirstChild<Top10>()!;
                Assert.False(bottomPercent.Top!.Value);
                Assert.True(bottomPercent.Percent!.Value);
                Assert.Equal(25D, bottomPercent.Val!.Value);
                Assert.Equal(PivotFilterValues.Sum, filters[2].Type!.Value);
                Assert.Equal("50", filters[2].StringValue1!.Value);
                Assert.Equal(0U, filters[2].MeasureField!.Value);
                var topSum = filters[2].AutoFilter!.Elements<FilterColumn>().Single().GetFirstChild<Top10>()!;
                Assert.True(topSum.Top!.Value);
                Assert.False(topSum.Percent!.Value);
                Assert.Equal(50D, topSum.Val!.Value);
                Assert.Equal(PivotFilterValues.Sum, filters[3].Type!.Value);
                Assert.Equal("30.5", filters[3].StringValue1!.Value);
                Assert.Equal(0U, filters[3].MeasureField!.Value);
                var bottomSum = filters[3].AutoFilter!.Elements<FilterColumn>().Single().GetFirstChild<Top10>()!;
                Assert.False(bottomSum.Top!.Value);
                Assert.False(bottomSum.Percent!.Value);
                Assert.Equal(30.5D, bottomSum.Val!.Value);
            }

            using (var document = ExcelDocument.Load(filePath))
            {
                var pivot = document.GetPivotTables().Single();
                Assert.Equal(4, pivot.Filters.Count);
                Assert.Equal(ExcelPivotFilterType.Count, pivot.Filters[0].Type);
                Assert.True(pivot.Filters[0].IsTop);
                Assert.False(pivot.Filters[0].IsPercent);
                Assert.Equal("2", pivot.Filters[0].Value1);
                Assert.Equal("Total Sales", pivot.Filters[0].DataFieldName);
                Assert.Equal(ExcelPivotFilterType.Percent, pivot.Filters[1].Type);
                Assert.False(pivot.Filters[1].IsTop);
                Assert.True(pivot.Filters[1].IsPercent);
                Assert.Equal("25", pivot.Filters[1].Value1);
                Assert.Equal(ExcelPivotFilterType.Sum, pivot.Filters[2].Type);
                Assert.True(pivot.Filters[2].IsTop);
                Assert.False(pivot.Filters[2].IsPercent);
                Assert.Equal("50", pivot.Filters[2].Value1);
                Assert.Equal(ExcelPivotFilterType.Sum, pivot.Filters[3].Type);
                Assert.False(pivot.Filters[3].IsTop);
                Assert.False(pivot.Filters[3].IsPercent);
                Assert.Equal("30.5", pivot.Filters[3].Value1);
            }

            using (var document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly }))
            {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_AddPivotTable_AppliesCalculatedFields()
        {
            var filePath = Path.Combine(_directoryWithFiles, "ExcelPivotTableCalculatedFields.xlsx");
            using (var document = ExcelDocument.Create(filePath))
            {
                var sheet = document.AddWorksheet("Data");
                sheet.CellValue(1, 1, "Region");
                sheet.CellValue(1, 2, "Revenue");
                sheet.CellValue(1, 3, "Cost");
                sheet.CellValue(2, 1, "East");
                sheet.CellValue(2, 2, 100);
                sheet.CellValue(2, 3, 60);
                sheet.CellValue(3, 1, "West");
                sheet.CellValue(3, 2, 120);
                sheet.CellValue(3, 3, 90);
                sheet.CellValue(4, 1, "East");
                sheet.CellValue(4, 2, 80);
                sheet.CellValue(4, 3, 45);
                {
                    var pivotOptions11 = new ExcelPivotTableOptions();
                    pivotOptions11.RowFields.Add("Region");
                    pivotOptions11.CalculatedFields.Add(new ExcelPivotCalculatedField("Margin", "'Revenue' - 'Cost'", numberFormat: "$#,##0"));
                    pivotOptions11.DataFields.Add(new ExcelPivotDataField("Margin", ExcelPivotDataFunction.Sum, displayName: "Total Margin", numberFormat: "$#,##0"));
                    pivotOptions11.Name = "MarginPivot";
                    sheet.AddPivotTable("A1:C4", "E2", pivotOptions11);
                }

                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false))
            {
                var pivotPart = spreadsheet.WorkbookPart!.WorksheetParts.SelectMany(part => part.PivotTableParts).Single();
                var cacheFields = pivotPart.PivotTableCacheDefinitionPart!.PivotCacheDefinition!.CacheFields!.Elements<CacheField>().ToList();
                Assert.Equal(4, cacheFields.Count);
                Assert.Equal("Margin", cacheFields[3].Name!.Value);
                Assert.Equal("'Revenue' - 'Cost'", cacheFields[3].Formula!.Value);
                Assert.False(cacheFields[3].DatabaseField!.Value);
                Assert.True(cacheFields[3].NumberFormatId!.Value >= 164);
                var pivotField = pivotPart.PivotTableDefinition!.PivotFields!.Elements<PivotField>().ElementAt(3);
                Assert.True(pivotField.DataField!.Value);
                var dataField = pivotPart.PivotTableDefinition!.DataFields!.Elements<DataField>().Single();
                Assert.Equal(3U, dataField.Field!.Value);
                Assert.Equal("Total Margin", dataField.Name!.Value);
                Assert.True(dataField.NumberFormatId!.Value >= 164);
            }

            using (var document = ExcelDocument.Load(filePath))
            {
                var pivot = document.GetPivotTables().Single();
                var calculated = pivot.CalculatedFields.Single();
                Assert.Equal("Margin", calculated.Name);
                Assert.Equal("'Revenue' - 'Cost'", calculated.Formula);
                Assert.True(calculated.NumberFormatId >= 164);
                var dataField = pivot.DataFields.Single();
                Assert.Equal("Margin", dataField.FieldName);
                Assert.Equal("Total Margin", dataField.DisplayName);
                Assert.Equal("$#,##0", dataField.NumberFormatCode);
                Assert.Equal("$#,##0", calculated.NumberFormatCode);
            }

            using (var document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly }))
            {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_AddPivotTable_AppliesDateAndNumberGrouping()
        {
            var filePath = Path.Combine(_directoryWithFiles, "ExcelPivotTableGrouping.xlsx");
            using (var document = ExcelDocument.Create(filePath))
            {
                var sheet = document.AddWorksheet("Data");
                sheet.CellValue(1, 1, "OrderDate");
                sheet.CellValue(1, 2, "Quantity");
                sheet.CellValue(1, 3, "Sales");
                sheet.CellValue(2, 1, new DateTime(2026, 1, 3));
                sheet.CellValue(2, 2, 4);
                sheet.CellValue(2, 3, 40);
                sheet.CellValue(3, 1, new DateTime(2026, 1, 18));
                sheet.CellValue(3, 2, 12);
                sheet.CellValue(3, 3, 120);
                sheet.CellValue(4, 1, new DateTime(2026, 2, 8));
                sheet.CellValue(4, 2, 18);
                sheet.CellValue(4, 3, 180);
                {
                    var pivotOptions12 = new ExcelPivotTableOptions();
                    pivotOptions12.RowFields.Add("OrderDate");
                    pivotOptions12.RowFields.Add("Quantity");
                    pivotOptions12.DataFields.Add(new ExcelPivotDataField("Sales", ExcelPivotDataFunction.Sum, displayName: "Total Sales"));
                    pivotOptions12.Groupings.Add(ExcelPivotGrouping.Date("OrderDate", ExcelPivotGroupBy.Months, new DateTime(2026, 1, 1), new DateTime(2026, 12, 31)));
                    pivotOptions12.Groupings.Add(ExcelPivotGrouping.Number("Quantity", 10, 0, 30));
                    pivotOptions12.Name = "SalesPivot";
                    sheet.AddPivotTable("A1:C4", "E2", pivotOptions12);
                }

                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false))
            {
                var pivotPart = spreadsheet.WorkbookPart!.WorksheetParts.SelectMany(part => part.PivotTableParts).Single();
                var cacheFields = pivotPart.PivotTableCacheDefinitionPart!.PivotCacheDefinition!.CacheFields!.Elements<CacheField>().ToList();
                var dateField = cacheFields[0];
                var dateRange = dateField.FieldGroup!.GetFirstChild<RangeProperties>()!;
                Assert.Equal(GroupByValues.Months, dateRange.GroupBy!.Value);
                Assert.Equal(new DateTime(2026, 1, 1), dateRange.StartDate!.Value);
                Assert.Equal(new DateTime(2026, 12, 31), dateRange.EndDate!.Value);
                Assert.False(dateRange.AutoStart!.Value);
                Assert.False(dateRange.AutoEnd!.Value);
                Assert.True(dateField.SharedItems!.ContainsDate!.Value);
                Assert.Equal(3, dateField.SharedItems.Elements<DateTimeItem>().Count());
                var numberField = cacheFields[1];
                var numberRange = numberField.FieldGroup!.GetFirstChild<RangeProperties>()!;
                Assert.Equal(GroupByValues.Range, numberRange.GroupBy!.Value);
                Assert.Equal(10d, numberRange.GroupInterval!.Value);
                Assert.Equal(0d, numberRange.StartNumber!.Value);
                Assert.Equal(30d, numberRange.EndNum!.Value);
                Assert.True(numberField.SharedItems!.ContainsNumber!.Value);
                Assert.Equal(3, numberField.SharedItems.Elements<NumberItem>().Count());
                var fields = pivotPart.PivotTableDefinition!.PivotFields!.Elements<PivotField>().ToArray();
                for (int fieldIndex = 0; fieldIndex < 2; fieldIndex++)
                {
                    var field = cacheFields[fieldIndex];
                    int keyCount = field.FieldGroup?.GetFirstChild<GroupItems>()?.ChildElements.Count ?? field.SharedItems!.ChildElements.Count;
                    Assert.Equal(Enumerable.Range(0, keyCount).Select(index => (uint)index), fields[fieldIndex].Items!.Elements<Item>().Where(item => item.Index != null).Select(item => item.Index!.Value));
                }

                OpenXmlValidator validator = new OpenXmlValidator();
                var errors = validator.Validate(spreadsheet).ToList();
                Assert.True(errors.Count == 0, FormatValidationErrors(errors));
            }

            using (var document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly }))
            {
                var pivot = document.GetPivotTables().Single();
                Assert.Equal(2, pivot.Groupings.Count);
                Assert.Equal("OrderDate", pivot.Groupings[0].FieldName);
                Assert.Equal(ExcelPivotGroupBy.Months, pivot.Groupings[0].GroupBy);
                Assert.Equal(new DateTime(2026, 1, 1), pivot.Groupings[0].StartDate);
                Assert.Equal("Quantity", pivot.Groupings[1].FieldName);
                Assert.Equal(ExcelPivotGroupBy.Range, pivot.Groupings[1].GroupBy);
                Assert.Equal(10d, pivot.Groupings[1].Interval);
            }
        }

        [Fact]
        public void Test_AddPivotTable_AppliesGeneratedDateHierarchyGrouping()
        {
            var filePath = Path.Combine(_directoryWithFiles, "ExcelPivotTableDateHierarchyGrouping.xlsx");
            using (var document = ExcelDocument.Create(filePath))
            {
                var sheet = document.AddWorksheet("Data");
                sheet.CellValue(1, 1, "OrderDate");
                sheet.CellValue(1, 2, "Region");
                sheet.CellValue(1, 3, "Sales");
                sheet.CellValue(2, 1, new DateTime(2025, 12, 30));
                sheet.CellValue(2, 2, "West");
                sheet.CellValue(2, 3, 75);
                sheet.CellValue(3, 1, new DateTime(2026, 1, 3));
                sheet.CellValue(3, 2, "East");
                sheet.CellValue(3, 3, 40);
                sheet.CellValue(4, 1, new DateTime(2026, 4, 8));
                sheet.CellValue(4, 2, "East");
                sheet.CellValue(4, 3, 180);
                {
                    var pivotOptions13 = new ExcelPivotTableOptions();
                    pivotOptions13.RowFields.Add("OrderDate");
                    pivotOptions13.DataFields.Add(new ExcelPivotDataField("Sales", ExcelPivotDataFunction.Sum, displayName: "Total Sales"));
                    pivotOptions13.Groupings.Add(ExcelPivotGrouping.DateHierarchy("OrderDate", ExcelPivotGroupBy.Years, ExcelPivotGroupBy.Quarters, ExcelPivotGroupBy.Months));
                    pivotOptions13.Name = "SalesPivot";
                    sheet.AddPivotTable("A1:C4", "E2", pivotOptions13);
                }

                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false))
            {
                var pivotPart = spreadsheet.WorkbookPart!.WorksheetParts.SelectMany(part => part.PivotTableParts).Single();
                var cacheFields = pivotPart.PivotTableCacheDefinitionPart!.PivotCacheDefinition!.CacheFields!.Elements<CacheField>().ToList();
                Assert.Equal(6, cacheFields.Count);
                Assert.Equal("OrderDate Years", cacheFields[3].Name!.Value);
                Assert.Equal("OrderDate Quarters", cacheFields[4].Name!.Value);
                Assert.Equal("OrderDate Months", cacheFields[5].Name!.Value);
                Assert.False(cacheFields[3].DatabaseField!.Value);
                Assert.False(cacheFields[4].DatabaseField!.Value);
                Assert.False(cacheFields[5].DatabaseField!.Value);
                Assert.Equal(GroupByValues.Years, cacheFields[3].FieldGroup!.GetFirstChild<RangeProperties>()!.GroupBy!.Value);
                Assert.Equal(GroupByValues.Quarters, cacheFields[4].FieldGroup!.GetFirstChild<RangeProperties>()!.GroupBy!.Value);
                Assert.Equal(GroupByValues.Months, cacheFields[5].FieldGroup!.GetFirstChild<RangeProperties>()!.GroupBy!.Value);
                Assert.Equal(0U, cacheFields[3].FieldGroup!.Base!.Value);
                Assert.Equal(0U, cacheFields[4].FieldGroup!.Base!.Value);
                Assert.Equal(0U, cacheFields[5].FieldGroup!.Base!.Value);
                Assert.Null(cacheFields[3].FieldGroup!.ParentId);
                Assert.Equal(3U, cacheFields[4].FieldGroup!.ParentId!.Value);
                Assert.Equal(4U, cacheFields[5].FieldGroup!.ParentId!.Value);
                Assert.Equal(3U, cacheFields[0].FieldGroup!.ParentId!.Value);
                Assert.Equal(4U, cacheFields[3].FieldGroup!.GetFirstChild<GroupItems>()!.Count!.Value);
                Assert.Equal(6U, cacheFields[4].FieldGroup!.GetFirstChild<GroupItems>()!.Count!.Value);
                Assert.Equal(14U, cacheFields[5].FieldGroup!.GetFirstChild<GroupItems>()!.Count!.Value);
                Assert.Contains(cacheFields[3].FieldGroup!.GetFirstChild<GroupItems>()!.Elements<StringItem>(), item => item.Val!.Value == "2026");
                Assert.Contains(cacheFields[4].FieldGroup!.GetFirstChild<GroupItems>()!.Elements<StringItem>(), item => item.Val!.Value == "Q2");
                Assert.Contains(cacheFields[5].FieldGroup!.GetFirstChild<GroupItems>()!.Elements<StringItem>(), item => item.Val!.Value == "April");
                Assert.Contains(cacheFields[3].SharedItems!.Elements<StringItem>(), item => item.Val!.Value == "2026");
                Assert.Contains(cacheFields[4].SharedItems!.Elements<StringItem>(), item => item.Val!.Value == "Q2");
                Assert.Contains(cacheFields[5].SharedItems!.Elements<StringItem>(), item => item.Val!.Value == "April");
                var rowFields = pivotPart.PivotTableDefinition!.RowFields!.Elements<Field>().ToList();
                Assert.Equal(3, rowFields.Count);
                Assert.Equal(3, rowFields[0].Index!.Value);
                Assert.Equal(4, rowFields[1].Index!.Value);
                Assert.Equal(5, rowFields[2].Index!.Value);
                OpenXmlValidator validator = new OpenXmlValidator();
                var errors = validator.Validate(spreadsheet).ToList();
                Assert.True(errors.Count == 0, FormatValidationErrors(errors));
            }

            using (var document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly }))
            {
                var pivot = document.GetPivotTables().Single();
                Assert.Equal(new[] { "OrderDate Years", "OrderDate Quarters", "OrderDate Months" }, pivot.RowFields);
                Assert.Contains(pivot.Groupings, grouping => grouping.FieldName == "OrderDate Years" && grouping.GroupBy == ExcelPivotGroupBy.Years && grouping.BaseFieldIndex == 0U && grouping.ParentFieldIndex == null && grouping.GroupItems.Contains("2026"));
                Assert.Contains(pivot.Groupings, grouping => grouping.FieldName == "OrderDate Quarters" && grouping.GroupBy == ExcelPivotGroupBy.Quarters && grouping.BaseFieldIndex == 0U && grouping.ParentFieldIndex == 3U && grouping.GroupItems.Contains("Q2"));
                Assert.Contains(pivot.Groupings, grouping => grouping.FieldName == "OrderDate Months" && grouping.GroupBy == ExcelPivotGroupBy.Months && grouping.BaseFieldIndex == 0U && grouping.ParentFieldIndex == 4U && grouping.GroupItems.Contains("April"));
            }
        }

        [Fact]
        public void Test_AddPivotTable_AppliesFieldOptionsToGeneratedDateHierarchyFields()
        {
            var filePath = Path.Combine(_directoryWithFiles, "ExcelPivotTableDateHierarchyFieldOptions.xlsx");
            using (var document = ExcelDocument.Create(filePath))
            {
                var sheet = document.AddWorksheet("Data");
                sheet.CellValue(1, 1, "OrderDate");
                sheet.CellValue(1, 2, "Region");
                sheet.CellValue(1, 3, "Sales");
                sheet.CellValue(2, 1, new DateTime(2025, 12, 30));
                sheet.CellValue(2, 2, "West");
                sheet.CellValue(2, 3, 75);
                sheet.CellValue(3, 1, new DateTime(2026, 1, 3));
                sheet.CellValue(3, 2, "East");
                sheet.CellValue(3, 3, 40);
                sheet.CellValue(4, 1, new DateTime(2026, 4, 8));
                sheet.CellValue(4, 2, "East");
                sheet.CellValue(4, 3, 180);
                {
                    var pivotOptions14 = new ExcelPivotTableOptions();
                    pivotOptions14.RowFields.Add("OrderDate");
                    pivotOptions14.DataFields.Add(new ExcelPivotDataField("Sales", ExcelPivotDataFunction.Sum, displayName: "Total Sales"));
                    pivotOptions14.Groupings.Add(ExcelPivotGrouping.DateHierarchy("OrderDate", ExcelPivotGroupBy.Years, ExcelPivotGroupBy.Months));
                    pivotOptions14.FieldOptions.Add(new ExcelPivotFieldOptions("OrderDate", sortType: ExcelPivotFieldSort.Descending, defaultSubtotal: false, compact: false, outline: true, hiddenItems: new[] { "2025" }));
                    pivotOptions14.Name = "SalesPivot";
                    sheet.AddPivotTable("A1:C4", "E2", pivotOptions14);
                }

                {
                    var pivotOptions15 = new ExcelPivotTableOptions();
                    pivotOptions15.DataFields.Add(new ExcelPivotDataField("Sales", ExcelPivotDataFunction.Sum, displayName: "Filtered Sales"));
                    pivotOptions15.Groupings.Add(ExcelPivotGrouping.DateHierarchy("OrderDate", ExcelPivotGroupBy.Years, ExcelPivotGroupBy.Months));
                    pivotOptions15.PageFields.Add("OrderDate");
                    pivotOptions15.FieldOptions.Add(new ExcelPivotFieldOptions("OrderDate", selectedItem: "2026"));
                    pivotOptions15.Name = "FilterPivot";
                    sheet.AddPivotTable("A1:C4", "E12", pivotOptions15);
                }

                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false))
            {
                var pivotParts = spreadsheet.WorkbookPart!.WorksheetParts.SelectMany(part => part.PivotTableParts).ToList();
                var pivotPart = pivotParts.Single(part => part.PivotTableDefinition!.Name == "SalesPivot");
                var pivotFields = pivotPart.PivotTableDefinition!.PivotFields!.Elements<PivotField>().ToList();
                Assert.Equal(5, pivotFields.Count);
                var sourceField = pivotFields[0];
                Assert.Null(sourceField.Items);
                var yearsField = pivotFields[3];
                Assert.Equal(FieldSortValues.Descending, yearsField.SortType!.Value);
                Assert.False(yearsField.DefaultSubtotal!.Value);
                Assert.False(yearsField.Compact!.Value);
                Assert.True(yearsField.Outline!.Value);
                var hiddenYear = Assert.Single(yearsField.Items!.Elements<Item>(), item => item.Hidden?.Value == true);
                Assert.Equal(1U, hiddenYear.Index!.Value);
                var monthsField = pivotFields[4];
                Assert.Equal(FieldSortValues.Descending, monthsField.SortType!.Value);
                Assert.False(monthsField.DefaultSubtotal!.Value);
                Assert.False(monthsField.Compact!.Value);
                Assert.True(monthsField.Outline!.Value);
                var monthCacheField = pivotPart.PivotTableCacheDefinitionPart!.PivotCacheDefinition!.CacheFields!.Elements<CacheField>().ElementAt(4);
                int monthKeys = monthCacheField.FieldGroup!.GetFirstChild<GroupItems>()!.ChildElements.Count;
                Assert.Equal(Enumerable.Range(0, monthKeys).Select(index => (uint)index), monthsField.Items!.Elements<Item>().Select(item => item.Index!.Value));
                var filterPivotPart = pivotParts.Single(part => part.PivotTableDefinition!.Name == "FilterPivot");
                var pageFields = filterPivotPart.PivotTableDefinition!.PageFields!.Elements<PageField>().ToList();
                Assert.Equal(2, pageFields.Count);
                Assert.Equal(3, pageFields[0].Field!.Value);
                Assert.Equal(2U, pageFields[0].Item!.Value);
                Assert.Equal(4, pageFields[1].Field!.Value);
                Assert.Null(pageFields[1].Item);
                OpenXmlValidator validator = new OpenXmlValidator();
                var errors = validator.Validate(spreadsheet).ToList();
                Assert.True(errors.Count == 0, FormatValidationErrors(errors));
            }
        }

        [Fact]
        public void Test_AddPivotChartFromRange_WritesPivotSource()
        {
            var filePath = Path.Combine(_directoryWithFiles, "ExcelPivotChart.xlsx");
            using (var document = ExcelDocument.Create(filePath))
            {
                var sheet = document.AddWorksheet("Data");
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
                sheet.AddPivotTable("A1:C4", "E2", options: new ExcelPivotTableOptions { Name = "SalesPivot", RowFields = new[] { "Region" }, ColumnFields = new[] { "Product" }, DataFields = new[] { new ExcelPivotDataField("Sales", ExcelPivotDataFunction.Sum, "Total Sales") } });
                var chart = sheet.AddPivotChartFromRange("SalesPivot", "A1:C4", row: 7, column: 1, title: "Sales Pivot");
                Assert.True(chart.IsPivotChart);
                Assert.Equal("SalesPivot", chart.PivotTableName);
                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false))
            {
                var worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First(part => part.DrawingsPart?.ChartParts.Any() == true);
                var chartPart = worksheetPart.DrawingsPart!.ChartParts.Single();
                var pivotSource = chartPart.ChartSpace!.GetFirstChild<C.PivotSource>();
                Assert.NotNull(pivotSource);
                Assert.Equal("SalesPivot", pivotSource!.GetFirstChild<C.PivotTableName>()!.Text);
                Assert.Equal(0U, pivotSource.GetFirstChild<C.FormatId>()!.Val!.Value);
            }

            using (var document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly }))
            {
                Assert.Empty(document.ValidateOpenXml());
            }
        }
    }
}