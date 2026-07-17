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
        public void Test_UpdatePivotTableSource_InvalidatesStaleCacheAndRefreshesOnOpen() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelPivotTable.UpdateSource.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet original = document.AddWorksheet("Original");
                WritePivotSource(original, "Region", "Product", "Sales", 3);
                original.AddPivotTable(
                    sourceRange: "A1:C4",
                    destinationCell: "E2",
                    name: "SalesPivot",
                    rowFields: new[] { "Region" },
                    columnFields: new[] { "Product" },
                    dataFields: new[] { new ExcelPivotDataField("Sales", DataConsolidateFunctionValues.Sum, "Total Sales") },
                    options: new ExcelPivotTableOptions { SaveSourceData = true });

                ExcelSheet mismatch = document.AddWorksheet("Mismatch");
                WritePivotSource(mismatch, "Area", "Product", "Sales", 4);
                InvalidOperationException mismatchException = Assert.Throws<InvalidOperationException>(() =>
                    original.UpdatePivotTableSource("SalesPivot", mismatch, "A1:C5"));
                Assert.Contains("do not match", mismatchException.Message, StringComparison.OrdinalIgnoreCase);
                Assert.Equal("Original", Assert.Single(document.GetPivotTables()).SourceSheet);

                ExcelSheet expanded = document.AddWorksheet("Expanded");
                WritePivotSource(expanded, "Region", "Product", "Sales", 4);
                ExcelPivotSourceUpdateResult update = original.UpdatePivotTableSource("SalesPivot", expanded, "$A$1:$C$5");

                Assert.Equal("SalesPivot", update.PivotTableName);
                Assert.Equal("Expanded", update.SourceSheet);
                Assert.Equal("A1:C5", update.SourceRange);
                Assert.Equal(3U, update.InvalidatedCachedRecordCount);
                Assert.Equal(new[] { "SalesPivot" }, update.AffectedPivotTables);

                ExcelPivotTableInfo pivot = Assert.Single(document.GetPivotTables());
                Assert.Equal("Expanded", pivot.SourceSheet);
                Assert.Equal("A1:C5", pivot.SourceRange);
                Assert.True(pivot.RefreshOnOpen);
                Assert.False(pivot.SaveSourceData);
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                PivotTablePart pivotPart = spreadsheet.WorkbookPart!.WorksheetParts
                    .SelectMany(part => part.PivotTableParts)
                    .Single();
                PivotCacheDefinition cache = pivotPart.PivotTableCacheDefinitionPart!.PivotCacheDefinition!;
                WorksheetSource source = cache.CacheSource!.WorksheetSource!;
                PivotCacheRecords records = pivotPart.PivotTableCacheDefinitionPart.PivotTableCacheRecordsPart!.PivotCacheRecords!;

                Assert.Equal("Expanded", source.Sheet!.Value);
                Assert.Equal("A1:C5", source.Reference!.Value);
                Assert.Null(source.Name);
                Assert.Null(source.Id);
                Assert.Equal(4U, cache.RecordCount!.Value);
                Assert.True(cache.RefreshOnLoad!.Value);
                Assert.False(cache.SaveData!.Value);
                Assert.Equal(0U, records.Count!.Value);
                Assert.Empty(records.ChildElements);
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                ExcelPivotTableInfo pivot = Assert.Single(document.GetPivotTables());
                Assert.Equal("Expanded", pivot.SourceSheet);
                Assert.Equal("A1:C5", pivot.SourceRange);
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Theory]
        [InlineData(true)]
        [InlineData(false)]
        public void Test_UpdatePivotTableSource_ClearsNamedAndExternalSourceAttributes(bool useNamedSource) {
            string sourceKind = useNamedSource ? "Named" : "External";
            string filePath = Path.Combine(_directoryWithFiles, $"ExcelPivotTable.UpdateSource.{sourceKind}.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet original = document.AddWorksheet("Original");
                WritePivotSource(original, "Region", "Product", "Sales", 3);
                original.AddPivotTable(
                    sourceRange: "A1:C4",
                    destinationCell: "E2",
                    name: "SalesPivot",
                    rowFields: new[] { "Region" },
                    dataFields: new[] { new ExcelPivotDataField("Sales", DataConsolidateFunctionValues.Sum) });

                PivotTableCacheDefinitionPart cachePart = original.WorksheetPart.PivotTableParts
                    .Single()
                    .PivotTableCacheDefinitionPart!;
                PivotCacheDefinition cache = cachePart.PivotCacheDefinition!;
                WorksheetSource source = cache.CacheSource!.WorksheetSource!;
                source.Sheet = null;
                source.Reference = null;
                if (useNamedSource) {
                    source.Name = "ExistingNamedSource";
                } else {
                    ExternalRelationship relationship = cachePart.AddExternalRelationship(
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath",
                        new Uri("https://example.test/source.xlsx"),
                        "rIdExternalSource");
                    source.Id = relationship.Id;
                }

                ExcelSheet expanded = document.AddWorksheet("Expanded");
                WritePivotSource(expanded, "Region", "Product", "Sales", 4);
                original.UpdatePivotTableSource("SalesPivot", expanded, "A1:C5");

                Assert.Equal("Expanded", source.Sheet!.Value);
                Assert.Equal("A1:C5", source.Reference!.Value);
                Assert.Null(source.Name);
                Assert.Null(source.Id);
                Assert.Empty(cachePart.ExternalRelationships);
                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_UpdatePivotTableSource_RequiresStableFieldCountWhenHeaderNamesAreRelaxed() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelPivotTable.UpdateSource.RelaxedHeaders.xlsx");

            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet original = document.AddWorksheet("Original");
            WritePivotSource(original, "Region", "Product", "Sales", 3);
            original.AddPivotTable(
                sourceRange: "A1:C4",
                destinationCell: "E2",
                name: "SalesPivot",
                rowFields: new[] { "Region" },
                dataFields: new[] { new ExcelPivotDataField("Sales", DataConsolidateFunctionValues.Sum) });

            ExcelSheet renamed = document.AddWorksheet("Renamed");
            WritePivotSource(renamed, "Area", "Item", "Amount", 4);
            var relaxed = new ExcelPivotSourceUpdateOptions { RequireMatchingHeaders = false };
            ExcelPivotSourceUpdateResult update = original.UpdatePivotTableSource("SalesPivot", renamed, "A1:C5", relaxed);
            Assert.Equal("Renamed", update.SourceSheet);

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
                original.UpdatePivotTableSource("SalesPivot", renamed, "A1:B5", relaxed));
            Assert.Contains("field count", exception.Message, StringComparison.OrdinalIgnoreCase);
        }

        private static void WritePivotSource(ExcelSheet sheet, string firstHeader, string secondHeader, string thirdHeader, int dataRows) {
            sheet.CellValue(1, 1, firstHeader);
            sheet.CellValue(1, 2, secondHeader);
            sheet.CellValue(1, 3, thirdHeader);
            for (int row = 0; row < dataRows; row++) {
                sheet.CellValue(row + 2, 1, row % 2 == 0 ? "East" : "West");
                sheet.CellValue(row + 2, 2, row % 2 == 0 ? "A" : "B");
                sheet.CellValue(row + 2, 3, (row + 1) * 10);
            }
        }
    }
}
