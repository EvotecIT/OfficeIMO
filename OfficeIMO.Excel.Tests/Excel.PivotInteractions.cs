using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_PivotInteractionCaches_ValidateBindingsAndReadBackMetadata() {
            string filePath = Path.Combine(_directoryWithFiles, "Excel.PivotInteractions.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("Sales");
                sheet.CellValue(1, 1, "Region");
                sheet.CellValue(1, 2, "OrderDate");
                sheet.CellValue(1, 3, "Sales");
                sheet.CellValue(2, 1, "East");
                sheet.CellValue(2, 2, new DateTime(2026, 1, 2));
                sheet.CellValue(2, 3, 10d);
                sheet.CellValue(3, 1, "West");
                sheet.CellValue(3, 2, new DateTime(2026, 2, 3));
                sheet.CellValue(3, 3, 20d);
                sheet.AddPivotTable(
                    sourceRange: "A1:C3",
                    destinationCell: "E2",
                    name: "SalesPivot",
                    rowFields: new[] { "Region" },
                    dataFields: new[] { new ExcelPivotDataField("Sales", DataConsolidateFunctionValues.Sum) });

                Assert.Throws<ArgumentException>(() => document.AddPivotSlicerCache("MissingPivot", "Region"));
                Assert.Throws<ArgumentException>(() => document.AddPivotSlicerCache("SalesPivot", "MissingField"));

                document.AddPivotSlicerCache("SalesPivot", "Region");
                Assert.Throws<ArgumentException>(() => document.AddPivotTimelineCache("SalesPivot", "Sales"));
                document.AddPivotTimelineCache("SalesPivot", "OrderDate", "SalesDateTimeline");
                Assert.Throws<InvalidOperationException>(() => document.AddPivotTimelineCache("SalesPivot", "OrderDate", "SalesDateTimeline"));

                ExcelPivotInteractionCacheInfo slicer = Assert.Single(document.GetWorkbookSlicerCaches());
                Assert.Equal(ExcelPivotInteractionCacheKind.Slicer, slicer.Kind);
                Assert.Equal("Slicer_Region", slicer.Name);
                Assert.Equal("Region", slicer.SourceName);
                Assert.Equal("SalesPivot", slicer.PivotTableName);
                Assert.False(string.IsNullOrWhiteSpace(slicer.RelationshipId));

                ExcelPivotInteractionCacheInfo timeline = Assert.Single(document.GetWorkbookTimelineCaches());
                Assert.Equal(ExcelPivotInteractionCacheKind.Timeline, timeline.Kind);
                Assert.Equal("SalesDateTimeline", timeline.Name);
                Assert.Equal("OrderDate", timeline.SourceName);
                Assert.Equal("SalesPivot", timeline.PivotTableName);
                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                Assert.Equal("Slicer_Region", Assert.Single(document.GetWorkbookSlicerCaches()).Name);
                Assert.Equal("SalesDateTimeline", Assert.Single(document.GetWorkbookTimelineCaches()).Name);
                ExcelWorkbookSnapshot snapshot = document.CreateInspectionSnapshot();
                Assert.False(snapshot.HasSlicers);
                Assert.False(snapshot.HasTimelines);
                Assert.Equal(1, snapshot.SlicerBindingMetadataPartCount);
                Assert.Equal(1, snapshot.TimelineBindingMetadataPartCount);
                Assert.True(snapshot.HasSlicerBindingMetadata);
                Assert.True(snapshot.HasTimelineBindingMetadata);
                Assert.Empty(document.ValidateOpenXml());
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                var metadataParts = spreadsheet.WorkbookPart!.Parts
                    .Select(pair => pair.OpenXmlPart)
                    .Where(part => part.ContentType.IndexOf("Cache-metadata", StringComparison.OrdinalIgnoreCase) >= 0)
                    .ToList();
                Assert.Equal(2, metadataParts.Count);
                Assert.All(metadataParts, part => Assert.StartsWith("application/vnd.officeimo.excel.", part.ContentType));
                Assert.DoesNotContain(spreadsheet.WorkbookPart.Parts, pair =>
                    pair.OpenXmlPart.RelationshipType.StartsWith("http://schemas.microsoft.com/office/", StringComparison.OrdinalIgnoreCase));
            }
        }

        [Fact]
        public void Test_PivotInteractionCaches_GenerateUniqueDefaultNames() {
            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("Sales");
            sheet.CellValue(1, 1, "Region");
            sheet.CellValue(1, 2, "OrderDate");
            sheet.CellValue(1, 3, "Sales");
            sheet.CellValue(2, 1, "East");
            sheet.CellValue(2, 2, new DateTime(2026, 1, 2));
            sheet.CellValue(2, 3, 10d);
            sheet.AddPivotTable(
                sourceRange: "A1:C2",
                destinationCell: "E2",
                name: "SalesPivot",
                rowFields: new[] { "Region" },
                dataFields: new[] { new ExcelPivotDataField("Sales", DataConsolidateFunctionValues.Sum) });

            document.AddPivotSlicerCache("SalesPivot", "Region");
            document.AddPivotSlicerCache("SalesPivot", "Region");
            document.AddPivotTimelineCache("SalesPivot", "OrderDate");
            document.AddPivotTimelineCache("SalesPivot", "OrderDate");

            Assert.Equal(new[] { "Slicer_Region", "Slicer_Region_2" },
                document.GetWorkbookSlicerCaches().Select(cache => cache.Name));
            Assert.Equal(new[] { "Timeline_OrderDate", "Timeline_OrderDate_2" },
                document.GetWorkbookTimelineCaches().Select(cache => cache.Name));

            document.AddPivotSlicerCache("SalesPivot", "Region", "ExplicitRegion");
            Assert.Throws<InvalidOperationException>(() =>
                document.AddPivotSlicerCache("SalesPivot", "Region", "ExplicitRegion"));
        }

        [Fact]
        public void Test_PivotTimelineCache_ValidatesRetargetedLiveSourceBeforeStaleCacheMetadata() {
            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet original = document.AddWorksheet("Original");
            original.CellValue(1, 1, "Region");
            original.CellValue(1, 2, "OrderDate");
            original.CellValue(1, 3, "Sales");
            original.CellValue(2, 1, "East");
            original.CellValue(2, 2, new DateTime(2026, 1, 2));
            original.CellValue(2, 3, 10d);
            original.AddPivotTable(
                sourceRange: "A1:C2",
                destinationCell: "E2",
                name: "SalesPivot",
                rowFields: new[] { "Region" },
                dataFields: new[] { new ExcelPivotDataField("Sales", DataConsolidateFunctionValues.Sum) });

            ExcelSheet replacement = document.AddWorksheet("Replacement");
            replacement.CellValue(1, 1, "Region");
            replacement.CellValue(1, 2, "OrderDate");
            replacement.CellValue(1, 3, "Sales");
            replacement.CellValue(2, 1, "West");
            replacement.CellValue(2, 2, "not a date");
            replacement.CellValue(2, 3, 20d);

            original.UpdatePivotTableSource("SalesPivot", replacement, "A1:C2");

            Assert.Throws<ArgumentException>(() => document.AddPivotTimelineCache("SalesPivot", "OrderDate"));
        }

        [Fact]
        public void Test_PivotTimelineCache_MapsRelaxedHeadersByCacheFieldPosition() {
            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet original = document.AddWorksheet("Original");
            original.CellValue(1, 1, "Region");
            original.CellValue(1, 2, "OrderDate");
            original.CellValue(1, 3, "Sales");
            original.CellValue(2, 1, "East");
            original.CellValue(2, 2, new DateTime(2026, 1, 2));
            original.CellValue(2, 3, 10d);
            original.AddPivotTable(
                sourceRange: "A1:C2",
                destinationCell: "E2",
                name: "SalesPivot",
                rowFields: new[] { "Region" },
                dataFields: new[] { new ExcelPivotDataField("Sales", DataConsolidateFunctionValues.Sum) });

            ExcelSheet replacement = document.AddWorksheet("Replacement");
            replacement.CellValue(1, 1, "OrderDate");
            replacement.CellValue(1, 2, "Date");
            replacement.CellValue(1, 3, "Amount");
            replacement.CellValue(2, 1, "not a date");
            replacement.CellValue(2, 2, new DateTime(2026, 2, 3));
            replacement.CellValue(2, 3, 20d);

            original.UpdatePivotTableSource(
                "SalesPivot",
                replacement,
                "A1:C2",
                new ExcelPivotSourceUpdateOptions { RequireMatchingHeaders = false });
            document.AddPivotTimelineCache("SalesPivot", "OrderDate");

            Assert.Equal("OrderDate", Assert.Single(document.GetWorkbookTimelineCaches()).SourceName);

            ExcelSheet misleading = document.AddWorksheet("Misleading");
            misleading.CellValue(1, 1, "OrderDate");
            misleading.CellValue(1, 2, "Date");
            misleading.CellValue(1, 3, "Amount");
            misleading.CellValue(2, 1, new DateTime(2026, 3, 4));
            misleading.CellValue(2, 2, "not a date");
            misleading.CellValue(2, 3, 30d);
            original.UpdatePivotTableSource(
                "SalesPivot",
                misleading,
                "A1:C2",
                new ExcelPivotSourceUpdateOptions { RequireMatchingHeaders = false });

            Assert.Throws<ArgumentException>(() =>
                document.AddPivotTimelineCache("SalesPivot", "OrderDate", "InvalidTimeline"));
        }

        [Fact]
        public void Test_PivotTimelineCache_UsesNormalizedDuplicateAndBlankSourceHeaders() {
            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("Sales");
            sheet.CellValue(1, 1, "Region");
            sheet.CellValue(1, 2, "Date");
            sheet.CellValue(1, 3, "Date");
            sheet.CellValue(1, 5, "Amount");
            sheet.CellValue(2, 1, "East");
            sheet.CellValue(2, 2, "not a date");
            sheet.CellValue(2, 3, new DateTime(2026, 1, 2));
            sheet.CellValue(2, 4, new DateTime(2026, 2, 3));
            sheet.CellValue(2, 5, 10d);
            sheet.AddPivotTable(
                sourceRange: "A1:E2",
                destinationCell: "G2",
                name: "SalesPivot",
                rowFields: new[] { "Region" },
                dataFields: new[] { new ExcelPivotDataField("Amount", DataConsolidateFunctionValues.Sum) });

            document.AddPivotTimelineCache("SalesPivot", "Date_2");
            document.AddPivotTimelineCache("SalesPivot", "Column4");

            Assert.Equal(
                new[] { "Column4", "Date_2" },
                document.GetWorkbookTimelineCaches()
                    .Select(cache => cache.SourceName)
                    .OrderBy(name => name, StringComparer.OrdinalIgnoreCase));
        }

        [Fact]
        public void Test_PivotTimelineCache_AcceptsStyledFormulaDateCaches() {
            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("Sales");
            sheet.CellValue(1, 1, "Region");
            sheet.CellValue(1, 2, "OrderDate");
            sheet.CellValue(1, 3, "Amount");
            sheet.CellValue(2, 1, "East");
            sheet.CellFormula(2, 2, "DATE(2026,1,2)");
            sheet.CellAt(2, 2).DateTime("yyyy-mm-dd");
            sheet.CellValue(2, 3, 10d);
            Assert.Equal(1, document.Calculate());

            Cell formulaCell = Assert.Single(sheet.WorksheetPart.Worksheet.Descendants<Cell>(), cell =>
                cell.CellReference?.Value == "B2");
            formulaCell.DataType = null;
            sheet.WorksheetPart.Worksheet.Save();

            sheet.AddPivotTable(
                sourceRange: "A1:C2",
                destinationCell: "E2",
                name: "SalesPivot",
                rowFields: new[] { "Region" },
                dataFields: new[] { new ExcelPivotDataField("Amount", DataConsolidateFunctionValues.Sum) });

            document.AddPivotTimelineCache("SalesPivot", "OrderDate");
            ExcelPivotInteractionCacheInfo timeline = Assert.Single(document.GetWorkbookTimelineCaches());
            Assert.Equal("OrderDate", timeline.SourceName);
        }
    }
}
