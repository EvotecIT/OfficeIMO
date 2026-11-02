using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests
{
    public partial class Excel
    {
        [Fact]
        public void Test_PivotTimeline_UsesCachedDateResultsFromFormulaCells()
        {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            try
            {
                using (ExcelDocument document = ExcelDocument.Create(filePath))
                {
                    ExcelSheet sheet = document.AddSheet("Sales");
                    sheet.Cell(1, 1).SetValue("Region");
                    sheet.Cell(1, 2).SetValue("OrderDate");
                    sheet.Cell(1, 3).SetValue("Sales");
                    sheet.Cell(2, 1).SetValue("East");
                    sheet.Cell(2, 2).SetValue(new DateTime(2026, 1, 2));
                    sheet.Cell(2, 3).SetValue(10D);
                    sheet.Cell(3, 1).SetValue("West");
                    sheet.Cell(3, 2).SetValue(new DateTime(2026, 2, 3));
                    sheet.Cell(3, 3).SetValue(20D);
                    document.Save();
                }

                using (SpreadsheetDocument package = SpreadsheetDocument.Open(filePath, true))
                {
                    Worksheet worksheet = package.WorkbookPart!.WorksheetParts.Single().Worksheet;
                    Cell[] dates = worksheet.Descendants<Cell>().Where(cell => cell.CellReference?.Value == "B2" || cell.CellReference?.Value == "B3").OrderBy(cell => cell.CellReference!.Value).ToArray();
                    dates[0].CellFormula = new CellFormula("DATE(2026,1,2)");
                    dates[1].CellFormula = new CellFormula("DATE(2026,2,3)");
                    worksheet.Save();
                }

                using (ExcelDocument document = ExcelDocument.Load(filePath))
                {
                    ExcelSheet sheet = document.Sheets[0];
                    sheet.AddPivotTable("A1:C3", "E2", options: new ExcelPivotTableOptions { Name = "FormulaDatePivot", RowFields = new[] { "Region" }, DataFields = new[] { new ExcelPivotDataField("Sales", ExcelPivotDataFunction.Sum) } });
                    document.AddPivotTimelineCache("FormulaDatePivot", "OrderDate", "FormulaDateTimeline");
                    Assert.Equal("FormulaDateTimeline", Assert.Single(document.GetWorkbookTimelineCaches()).Name);
                }
            }
            finally
            {
                if (File.Exists(filePath))
                    File.Delete(filePath);
            }
        }

        [Fact]
        public void Test_PivotInteractionCaches_ValidateBindingsAndReadBackMetadata()
        {
            string filePath = Path.Combine(_directoryWithFiles, "Excel.PivotInteractions.xlsx");
            using (ExcelDocument document = ExcelDocument.Create(filePath))
            {
                ExcelSheet sheet = document.AddSheet("Sales");
                sheet.Cell(1, 1).SetValue("Region");
                sheet.Cell(1, 2).SetValue("OrderDate");
                sheet.Cell(1, 3).SetValue("Sales");
                sheet.Cell(2, 1).SetValue("East");
                sheet.Cell(2, 2).SetValue(new DateTime(2026, 1, 2));
                sheet.Cell(2, 3).SetValue(10d);
                sheet.Cell(3, 1).SetValue("West");
                sheet.Cell(3, 2).SetValue(new DateTime(2026, 2, 3));
                sheet.Cell(3, 3).SetValue(20d);
                sheet.Cell(4, 1).SetValue("North");
                sheet.Cell(4, 2).SetValue(string.Empty);
                sheet.Cell(4, 3).SetValue(5d);
                sheet.AddPivotTable("A1:C4", "E2", options: new ExcelPivotTableOptions { Name = "SalesPivot", RowFields = new[] { "Region" }, DataFields = new[] { new ExcelPivotDataField("Sales", ExcelPivotDataFunction.Sum) } });
                Assert.True(sheet.TryGetCellValueSnapshot(4, 2, out ExcelNativeCellValueInfo? blankDate));
                Assert.Equal(ExcelCellValueKind.Text, blankDate!.Kind);
                Assert.Equal(string.Empty, blankDate.Text);
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

            using (ExcelDocument document = ExcelDocument.Load(filePath, new ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly }))
            {
                Assert.Equal("Slicer_Region", Assert.Single(document.GetWorkbookSlicerCaches()).Name);
                Assert.Equal("SalesDateTimeline", Assert.Single(document.GetWorkbookTimelineCaches()).Name);
                ExcelWorkbookInfo snapshot = document.Inspect();
                Assert.False(snapshot.HasSlicers);
                Assert.False(snapshot.HasTimelines);
                Assert.Equal(1, snapshot.SlicerBindingMetadataPartCount);
                Assert.Equal(1, snapshot.TimelineBindingMetadataPartCount);
                Assert.True(snapshot.HasSlicerBindingMetadata);
                Assert.True(snapshot.HasTimelineBindingMetadata);
                Assert.Empty(document.ValidateOpenXml());
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false))
            {
                WorkbookPart workbookPart = spreadsheet.WorkbookPart!;
                CustomXmlPart metadataPart = Assert.Single(workbookPart.CustomXmlParts, part => ReadPivotInteractionMetadataText(part) != null);
                Assert.Equal("application/xml", metadataPart.ContentType);
                Assert.EndsWith("/customXml", metadataPart.RelationshipType, StringComparison.Ordinal);
                Assert.DoesNotContain(spreadsheet.WorkbookPart.Parts, pair => pair.OpenXmlPart.RelationshipType.StartsWith("http://schemas.microsoft.com/office/", StringComparison.OrdinalIgnoreCase));
            }
        }

        [Fact]
        public void Test_PivotInteractionCaches_GenerateUniqueDefaultNames()
        {
            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddSheet("Sales");
            sheet.Cell(1, 1).SetValue("Region");
            sheet.Cell(1, 2).SetValue("OrderDate");
            sheet.Cell(1, 3).SetValue("Sales");
            sheet.Cell(2, 1).SetValue("East");
            sheet.Cell(2, 2).SetValue(new DateTime(2026, 1, 2));
            sheet.Cell(2, 3).SetValue(10d);
            sheet.AddPivotTable("A1:C2", "E2", options: new ExcelPivotTableOptions { Name = "SalesPivot", RowFields = new[] { "Region" }, DataFields = new[] { new ExcelPivotDataField("Sales", ExcelPivotDataFunction.Sum) } });
            WorkbookPart workbookPart = document._spreadSheetDocument.WorkbookPart!;
            WriteExtendedPart(workbookPart.AddExtendedPart("http://schemas.microsoft.com/office/2007/relationships/slicerCache", "application/vnd.ms-excel.slicerCache+xml", "xml"), "<slicerCacheDefinition xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\" name=\"Slicer_Region\" sourceName=\"Region\" pivotTableName=\"SalesPivot\"/>");
            WriteExtendedPart(workbookPart.AddExtendedPart("http://schemas.microsoft.com/office/2011/relationships/timelineCache", "application/vnd.ms-excel.timelineCache+xml", "xml"), "<timelineCacheDefinition xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2011/1/main\" name=\"Timeline_OrderDate\" sourceName=\"OrderDate\" pivotTableName=\"SalesPivot\"/>");
            document.AddPivotSlicerCache("SalesPivot", "Region");
            document.AddPivotSlicerCache("SalesPivot", "Region");
            document.AddPivotTimelineCache("SalesPivot", "OrderDate");
            document.AddPivotTimelineCache("SalesPivot", "OrderDate");
            Assert.Equal(new[] { "Slicer_Region", "Slicer_Region_2", "Slicer_Region_3" }, document.GetWorkbookSlicerCaches().Select(cache => cache.Name));
            Assert.Equal(new[] { "Timeline_OrderDate", "Timeline_OrderDate_2", "Timeline_OrderDate_3" }, document.GetWorkbookTimelineCaches().Select(cache => cache.Name));
            document.AddPivotSlicerCache("SalesPivot", "Region", "ExplicitRegion");
            Assert.Throws<InvalidOperationException>(() => document.AddPivotSlicerCache("SalesPivot", "Region", "ExplicitRegion"));
        }

#if EXCEL_PERFORMANCE_EVIDENCE
        [Trait("Category", "Performance")]
        [Trait("Category", "ExcelPerformanceEvidence")]
#endif
        [Fact]
        public void Test_PivotInteractionCaches_BoundDefaultNameSearchToExistingCaches()
        {
            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddSheet("Sales");
            sheet.Cell(1, 1).SetValue("Region");
            sheet.Cell(1, 2).SetValue("Sales");
            sheet.Cell(2, 1).SetValue("East");
            sheet.Cell(2, 2).SetValue(10d);
            sheet.AddPivotTable("A1:B2", "D2", options: new ExcelPivotTableOptions { Name = "SalesPivot", RowFields = new[] { "Region" }, DataFields = new[] { new ExcelPivotDataField("Sales", ExcelPivotDataFunction.Sum) } });
            WorkbookPart workbookPart = document._spreadSheetDocument.WorkbookPart!;
#if EXCEL_PERFORMANCE_EVIDENCE
            const int existingCount = 512;
#else
            const int existingCount = 64;
#endif
            for (int suffix = 1; suffix <= existingCount; suffix++)
            {
                string name = suffix == 1 ? "Slicer_Region" : "Slicer_Region_" + suffix;
                WriteExtendedPart(workbookPart.AddExtendedPart("http://schemas.microsoft.com/office/2007/relationships/slicerCache", "application/vnd.ms-excel.slicerCache+xml", "xml"), "<slicerCacheDefinition xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\" name=\"" + name + "\" sourceName=\"Region\" pivotTableName=\"SalesPivot\"/>");
            }

            IReadOnlyList<ExcelPivotInteractionCacheInfo> existing = document.GetWorkbookSlicerCaches();
            Assert.True(existing.Count > 50, existing.Count.ToString());
            var existingNames = new HashSet<string>(existing.Select(cache => cache.Name), StringComparer.OrdinalIgnoreCase);
            int expectedSuffix = Enumerable.Range(2, existing.Count + 1).First(suffix => !existingNames.Contains("Slicer_Region_" + suffix));
#if EXCEL_PERFORMANCE_EVIDENCE
            var stopwatch = System.Diagnostics.Stopwatch.StartNew();
#endif
            document.AddPivotSlicerCache("SalesPivot", "Region");
#if EXCEL_PERFORMANCE_EVIDENCE
            stopwatch.Stop();
#endif
            Assert.Contains(document.GetWorkbookSlicerCaches(), cache => cache.Name == "Slicer_Region_" + expectedSuffix);
#if EXCEL_PERFORMANCE_EVIDENCE
            Assert.True(stopwatch.Elapsed < TimeSpan.FromSeconds(5), stopwatch.Elapsed.ToString());
#endif
        }

        [Fact]
        public void Test_PivotTimelineCache_UsesDecisiveCacheBeforeUntrustedSourceRange()
        {
            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet original = document.AddSheet("Original");
            original.Cell(1, 1).SetValue("Region");
            original.Cell(1, 2).SetValue("OrderDate");
            original.Cell(1, 3).SetValue("Sales");
            original.Cell(2, 1).SetValue("East");
            original.Cell(2, 2).SetValue(new DateTime(2026, 1, 2));
            original.Cell(2, 3).SetValue(10d);
            original.AddPivotTable("A1:C2", "E2", options: new ExcelPivotTableOptions { Name = "SalesPivot", RowFields = new[] { "Region" }, DataFields = new[] { new ExcelPivotDataField("Sales", ExcelPivotDataFunction.Sum) } });
            PivotTablePart pivotPart = original.WorksheetPart.PivotTableParts.Single();
            PivotCacheDefinition cache = pivotPart.PivotTableCacheDefinitionPart!.PivotCacheDefinition!;
            CacheField orderDate = cache.CacheFields!.Elements<CacheField>().Single(field => field.Name?.Value == "OrderDate");
            orderDate.SharedItems = new SharedItems(new DateTimeItem { Val = new DateTime(2026, 1, 2) })
            {
                ContainsDate = true,
                ContainsString = false,
                ContainsNumber = false,
                Count = 1U
            };
            cache.CacheSource!.WorksheetSource!.Reference = "A1:A2147483647";
            cache.Save();
            document.AddPivotTimelineCache("SalesPivot", "OrderDate");
            Assert.Equal("OrderDate", Assert.Single(document.GetWorkbookTimelineCaches()).SourceName);
        }

        [Fact]
        public void Test_PivotTimelineCache_MapsRelaxedHeadersByCacheFieldPosition()
        {
            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet original = document.AddSheet("Original");
            original.Cell(1, 1).SetValue("Region");
            original.Cell(1, 2).SetValue("OrderDate");
            original.Cell(1, 3).SetValue("Sales");
            original.Cell(2, 1).SetValue("East");
            original.Cell(2, 2).SetValue(new DateTime(2026, 1, 2));
            original.Cell(2, 3).SetValue(10d);
            original.AddPivotTable("A1:C2", "E2", options: new ExcelPivotTableOptions { Name = "SalesPivot", RowFields = new[] { "Region" }, DataFields = new[] { new ExcelPivotDataField("Sales", ExcelPivotDataFunction.Sum) } });
            ExcelSheet replacement = document.AddSheet("Replacement");
            replacement.Cell(1, 1).SetValue("OrderDate");
            replacement.Cell(1, 2).SetValue("Date");
            replacement.Cell(1, 3).SetValue("Amount");
            replacement.Cell(2, 1).SetValue("not a date");
            replacement.Cell(2, 2).SetValue(new DateTime(2026, 2, 3));
            replacement.Cell(2, 3).SetValue(20d);
            original.UpdatePivotTableSource("SalesPivot", replacement, "A1:C2", new ExcelPivotSourceUpdateOptions { RequireMatchingHeaders = false });
            document.AddPivotTimelineCache("SalesPivot", "OrderDate");
            Assert.Equal("OrderDate", Assert.Single(document.GetWorkbookTimelineCaches()).SourceName);
            ExcelSheet misleading = document.AddSheet("Misleading");
            misleading.Cell(1, 1).SetValue("OrderDate");
            misleading.Cell(1, 2).SetValue("Date");
            misleading.Cell(1, 3).SetValue("Amount");
            misleading.Cell(2, 1).SetValue(new DateTime(2026, 3, 4));
            misleading.Cell(2, 2).SetValue("not a date");
            misleading.Cell(2, 3).SetValue(30d);
            original.UpdatePivotTableSource("SalesPivot", misleading, "A1:C2", new ExcelPivotSourceUpdateOptions { RequireMatchingHeaders = false });
            Assert.Throws<ArgumentException>(() => document.AddPivotTimelineCache("SalesPivot", "OrderDate", "InvalidTimeline"));
        }

        [Fact]
        public void Test_PivotTimelineCache_UsesNormalizedDuplicateAndBlankSourceHeaders()
        {
            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddSheet("Sales");
            sheet.Cell(1, 1).SetValue("Region");
            sheet.Cell(1, 2).SetValue("Date");
            sheet.Cell(1, 3).SetValue("Date");
            sheet.Cell(1, 5).SetValue("Amount");
            sheet.Cell(2, 1).SetValue("East");
            sheet.Cell(2, 2).SetValue("not a date");
            sheet.Cell(2, 3).SetValue(new DateTime(2026, 1, 2));
            sheet.Cell(2, 4).SetValue(new DateTime(2026, 2, 3));
            sheet.Cell(2, 5).SetValue(10d);
            sheet.AddPivotTable("A1:E2", "G2", options: new ExcelPivotTableOptions { Name = "SalesPivot", RowFields = new[] { "Region" }, DataFields = new[] { new ExcelPivotDataField("Amount", ExcelPivotDataFunction.Sum) } });
            document.AddPivotTimelineCache("SalesPivot", "Date_2");
            document.AddPivotTimelineCache("SalesPivot", "Column4");
            Assert.Equal(new[] { "Column4", "Date_2" }, document.GetWorkbookTimelineCaches().Select(cache => cache.SourceName).OrderBy(name => name, StringComparer.OrdinalIgnoreCase));
        }

        [Fact]
        public void Test_PivotTimelineCache_AcceptsStyledFormulaDateCaches()
        {
            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddSheet("Sales");
            sheet.Cell(1, 1).SetValue("Region");
            sheet.Cell(1, 2).SetValue("OrderDate");
            sheet.Cell(1, 3).SetValue("Amount");
            sheet.Cell(2, 1).SetValue("East");
            sheet.Cell(2, 2).SetFormula("DATE(2026,1,2)");
            sheet.Cell(2, 2).DateTime("yyyy-mm-dd");
            sheet.Cell(2, 3).SetValue(10d);
            Assert.Equal(1, document.Calculate());
            Cell formulaCell = Assert.Single(sheet.WorksheetPart.Worksheet.Descendants<Cell>(), cell => cell.CellReference?.Value == "B2");
            formulaCell.DataType = null;
            sheet.WorksheetPart.Worksheet.Save();
            sheet.AddPivotTable("A1:C2", "E2", options: new ExcelPivotTableOptions { Name = "SalesPivot", RowFields = new[] { "Region" }, DataFields = new[] { new ExcelPivotDataField("Amount", ExcelPivotDataFunction.Sum) } });
            document.AddPivotTimelineCache("SalesPivot", "OrderDate");
            ExcelPivotInteractionCacheInfo timeline = Assert.Single(document.GetWorkbookTimelineCaches());
            Assert.Equal("OrderDate", timeline.SourceName);
        }
    }
}