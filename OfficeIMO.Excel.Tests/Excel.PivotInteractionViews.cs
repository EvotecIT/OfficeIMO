using System;
using System.IO;
using System.Linq;
using System.Threading;
using DocumentFormat.OpenXml.Office2010.Drawing.Slicer;
using DocumentFormat.OpenXml.Office2013.Drawing.TimeSlicer;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using X15 = DocumentFormat.OpenXml.Office2013.Excel;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_PivotInteractionViews_CreateRoundTripReuseAndRemoveNativeParts() {
            string filePath = Path.Combine(_directoryWithFiles, "Excel.PivotInteractionViews.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                document.AddWorksheet("Dummy");
                ExcelSheet data = document.AddWorksheet("Sales");
                data.CellValue(1, 1, "Region");
                data.CellValue(1, 2, "OrderDate");
                data.CellValue(1, 3, "Sales");
                data.CellValue(2, 1, "East");
                data.CellValue(2, 2, new DateTime(2026, 1, 2));
                data.CellValue(2, 3, 10d);
                data.CellValue(3, 1, "West");
                data.CellValue(3, 2, new DateTime(2026, 2, 3));
                data.CellValue(3, 3, 20d);
                data.AddPivotTable(
                    sourceRange: "A1:C3",
                    destinationCell: "E2",
                    name: "SalesPivot",
                    rowFields: new[] { "Region" },
                    dataFields: new[] { new ExcelPivotDataField("Sales", DataConsolidateFunctionValues.Sum) });
                data.AddPivotTable(
                    sourceRange: "A1:C3",
                    destinationCell: "E10",
                    name: "SalesPivot2",
                    rowFields: new[] { "Region" },
                    dataFields: new[] { new ExcelPivotDataField("Sales", DataConsolidateFunctionValues.Sum) });
                PivotTablePart[] pivotParts = data.WorksheetPart.PivotTableParts
                    .OrderBy(part => part.PivotTableDefinition!.Name!.Value, StringComparer.Ordinal)
                    .ToArray();
                PivotTableCacheDefinitionPart sharedDefinition = pivotParts[0].PivotTableCacheDefinitionPart!;
                PivotTableCacheDefinitionPart discardedDefinition = pivotParts[1].PivotTableCacheDefinitionPart!;
                uint sharedCacheId = pivotParts[0].PivotTableDefinition!.CacheId!.Value;
                uint discardedCacheId = pivotParts[1].PivotTableDefinition!.CacheId!.Value;
                pivotParts[1].DeletePart(discardedDefinition);
                pivotParts[1].AddPart(sharedDefinition);
                pivotParts[1].PivotTableDefinition.CacheId = sharedCacheId;
                pivotParts[1].PivotTableDefinition.Save();
                WorkbookPart workbookPart = document._spreadSheetDocument.WorkbookPart!;
                PivotCache discardedCache = workbookPart.Workbook.PivotCaches!.Elements<PivotCache>()
                    .Single(cache => cache.CacheId?.Value == discardedCacheId);
                discardedCache.Remove();
                workbookPart.DeletePart(discardedDefinition);
                workbookPart.Workbook.Save();
                ExcelSheet dashboard = document.AddWorksheet("Dashboard");
                document.RemoveWorksheet("Dummy");

                ExcelPivotInteractionInfo slicer = document.AddPivotSlicer(
                    "SalesPivot",
                    "Region",
                    dashboard.Name,
                    new ExcelSlicerViewOptions { Name = "RegionFilter", Row = 2, Column = 2 });
                ExcelPivotInteractionInfo secondSlicer = document.AddPivotSlicer(
                    "SalesPivot2",
                    "Region",
                    dashboard.Name,
                    new ExcelSlicerViewOptions { Name = "RegionFilter2", CacheName = slicer.CacheName, Row = 2, Column = 5 });
                ExcelPivotInteractionInfo timeline = document.AddPivotTimeline(
                    "SalesPivot",
                    "OrderDate",
                    dashboard.Name,
                    new ExcelTimelineViewOptions { Name = "OrderTimeline", Row = 16, Column = 2 });
                ExcelPivotInteractionInfo salesSlicer = document.AddPivotSlicer(
                    "SalesPivot2",
                    "Sales",
                    dashboard.Name,
                    new ExcelSlicerViewOptions { Name = "SalesFilter", Row = 9, Column = 5 });

                Assert.Equal(slicer.CacheName, secondSlicer.CacheName);
                Assert.NotEqual(slicer.CacheName, salesSlicer.CacheName);
                Assert.Equal(4, document.GetPivotInteractions().Count);
                Assert.Equal(2, document._spreadSheetDocument.WorkbookPart!.SlicerCacheParts.Count());
                Assert.Equal(
                    2,
                    document._spreadSheetDocument.WorkbookPart.SlicerCacheParts.Single(part =>
                        string.Equals(part.SlicerCacheDefinition!.Name!.Value, slicer.CacheName, StringComparison.OrdinalIgnoreCase))
                        .SlicerCacheDefinition!.SlicerCachePivotTables!.ChildElements.Count);
                uint stableSalesSheetId = workbookPart.Workbook.Sheets!.Elements<Sheet>()
                    .Single(sheet => sheet.Name!.Value == "Sales").SheetId!.Value;
                Assert.Equal(2U, stableSalesSheetId);
                Assert.All(
                    document._spreadSheetDocument.WorkbookPart.SlicerCacheParts
                        .SelectMany(part => part.SlicerCacheDefinition!.SlicerCachePivotTables!.Elements<X14.SlicerCachePivotTable>()),
                    target => Assert.Equal(stableSalesSheetId, target.TabId!.Value));
                Assert.Single(document._spreadSheetDocument.WorkbookPart.TimeLineCacheParts);
                Assert.Empty(document.ValidateOpenXml());
                Assert.Equal(
                    ExcelFeatureSupportLevel.PartiallyEditable,
                    Assert.Single(document.InspectFeatures().FindFeatures("Slicers")).SupportLevel);

                SlicerCachePart originalCache = document._spreadSheetDocument.WorkbookPart.SlicerCacheParts.First();
                SlicersPart originalViews = dashboard.WorksheetPart.SlicersParts.Single();
                Assert.Throws<InvalidOperationException>(() => dashboard.ApplyTransactionalMutation(_ => {
                    document._spreadSheetDocument.WorkbookPart.DeletePart(originalCache);
                    dashboard.WorksheetPart.DeletePart(originalViews);
                    throw new InvalidOperationException("Rollback probe");
                }, new ExcelMutationPlanOptions(), CancellationToken.None));
                Assert.Equal(2, document._spreadSheetDocument.WorkbookPart.SlicerCacheParts.Count());
                Assert.Equal(3, dashboard.WorksheetPart.SlicersParts.Single().Slicers!.ChildElements.Count);

                Assert.True(document.RemovePivotInteraction("RegionFilter"));
                Assert.Equal(2, document._spreadSheetDocument.WorkbookPart.SlicerCacheParts.Count());
                Assert.True(document.RemovePivotInteraction("RegionFilter2"));
                Assert.Single(document._spreadSheetDocument.WorkbookPart.SlicerCacheParts);
                PivotCacheDefinition slicerPivotCache = data.WorksheetPart.PivotTableParts.First()
                    .PivotTableCacheDefinitionPart!.PivotCacheDefinition!;
                Assert.Contains(
                    slicerPivotCache.PivotCacheDefinitionExtensionList!.Elements<PivotCacheDefinitionExtension>(),
                    extension => extension.Uri?.Value == "{725AE2AE-9491-48BE-B2B4-4EB974FC3084}");
                Assert.True(document.RemovePivotInteraction("SalesFilter"));
                Assert.Empty(document._spreadSheetDocument.WorkbookPart.SlicerCacheParts);
                Assert.DoesNotContain(
                    slicerPivotCache.PivotCacheDefinitionExtensionList?.Elements<PivotCacheDefinitionExtension>()
                        ?? Enumerable.Empty<PivotCacheDefinitionExtension>(),
                    extension => extension.Uri?.Value == "{725AE2AE-9491-48BE-B2B4-4EB974FC3084}");
                Assert.Single(document.GetPivotInteractions());
                Assert.Equal(timeline.Name, document.GetPivotInteractions()[0].Name);
                Assert.Empty(document.ValidateOpenXml());
                document.Save();
            }

            using (SpreadsheetDocument package = SpreadsheetDocument.Open(filePath, false)) {
                WorkbookPart workbookPart = package.WorkbookPart!;
                Assert.Empty(workbookPart.SlicerCacheParts);
                Assert.Single(workbookPart.TimeLineCacheParts);
                WorksheetPart dashboardPart = workbookPart.WorksheetParts.Single(part =>
                    part.Worksheet?.Descendants<X15.TimelineReference>().Any() == true);
                Assert.Empty(dashboardPart.SlicersParts);
                Assert.Single(dashboardPart.TimeLineParts);
                Assert.Single(dashboardPart.DrawingsPart!.WorksheetDrawing!.Descendants<TimeSlicer>());
                Assert.Empty(dashboardPart.DrawingsPart.WorksheetDrawing.Descendants<Slicer>());
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath)) {
                ExcelPivotInteractionInfo timeline = Assert.Single(document.GetPivotInteractions());
                Assert.Equal(ExcelPivotInteractionCacheKind.Timeline, timeline.Kind);
                Assert.Equal("OrderTimeline", timeline.Name);
                document.RemoveWorksheet("Sales");
                Assert.Empty(document.GetPivotInteractions());
                Assert.Empty(document._spreadSheetDocument.WorkbookPart!.TimeLineCacheParts);
                Assert.Null(document["Dashboard"].WorksheetPart.DrawingsPart);
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_PivotInteractionViews_RejectInvalidBindingsAndPreserveDrawingPlacement() {
            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("Sales");
            sheet.CellValue(1, 1, "Region");
            sheet.CellValue(1, 2, "Sales");
            sheet.CellValue(2, 1, "East");
            sheet.CellValue(2, 2, 10d);
            sheet.AddPivotTable(
                sourceRange: "A1:B2",
                destinationCell: "D2",
                name: "SalesPivot",
                rowFields: new[] { "Region" },
                dataFields: new[] { new ExcelPivotDataField("Sales", DataConsolidateFunctionValues.Sum) });
            var beforeErrors = document.ValidateOpenXml().ToArray();
            Assert.True(
                beforeErrors.Length == 0,
                string.Join(Environment.NewLine, beforeErrors)
                + Environment.NewLine
                + sheet.WorksheetPart.PivotTableParts.Single().PivotTableDefinition!.OuterXml);

            Assert.Throws<ArgumentException>(() => document.AddPivotSlicer("Missing", "Region", "Sales"));
            Assert.Throws<ArgumentException>(() => document.AddPivotTimeline("SalesPivot", "Region", "Sales"));
            Assert.Throws<ArgumentException>(() => document.AddPivotSlicer(
                "SalesPivot",
                "Region",
                "Sales",
                new ExcelSlicerViewOptions { Style = "Unsupported" }));
            Assert.Throws<ArgumentOutOfRangeException>(() => document.AddPivotSlicer(
                "SalesPivot",
                "Region",
                "Sales",
                new ExcelSlicerViewOptions { Column = 16385 }));

            document.AddPivotSlicer(
                "SalesPivot",
                "Region",
                "Sales",
                new ExcelSlicerViewOptions { Name = "RegionFilter", Row = 5, Column = 6 });
            Xdr.OneCellAnchor anchor = Assert.Single(
                sheet.WorksheetPart.DrawingsPart!.WorksheetDrawing!.Elements<Xdr.OneCellAnchor>(),
                item => item.Descendants<Slicer>().Any());
            Assert.Equal("5", anchor.FromMarker!.ColumnId!.Text);
            Assert.Equal("4", anchor.FromMarker.RowId!.Text);

            sheet.InsertColumns(3, 2);
            anchor = Assert.Single(
                sheet.WorksheetPart.DrawingsPart.WorksheetDrawing!.Elements<Xdr.OneCellAnchor>(),
                item => item.Descendants<Slicer>().Any());
            Assert.Equal("7", anchor.FromMarker!.ColumnId!.Text);
            Assert.Equal("4", anchor.FromMarker.RowId!.Text);
            var errors = document.ValidateOpenXml().ToArray();
            Assert.True(errors.Length == 0, string.Join(Environment.NewLine, errors));

            document.AddWorksheet("Keep");
            document.RemoveWorksheet("Sales");
            Assert.Empty(document.GetPivotInteractions());
            Assert.Empty(document._spreadSheetDocument.WorkbookPart!.SlicerCacheParts);
            Assert.Empty(document.ValidateOpenXml());
        }
    }
}
