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
    }
}
