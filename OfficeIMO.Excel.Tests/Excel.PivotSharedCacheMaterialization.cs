using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        private static string PivotSharedCacheOraclePath => Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
            "Documents", "ExcelPivotCorpus", "pivot-shared-cache-views-conformance.xlsx");

        [Fact]
        public void Test_PivotSharedCacheMaterialization_RefreshesBothExcelAuthoredViews() {
            string path = Path.Combine(_directoryWithFiles, "PivotSharedCache.Materialized.xlsx");
            File.Copy(PivotSharedCacheOraclePath, path, true);
            using (var document = ExcelDocument.Load(path)) {
                ExcelSheet rows = document.GetSheet("Rows");
                ExcelSheet columns = document.GetSheet("Columns");
                PivotTableCacheDefinitionPart cache = rows.WorksheetPart.PivotTableParts.Single().PivotTableCacheDefinitionPart!;
                Assert.Same(cache, columns.WorksheetPart.PivotTableParts.Single().PivotTableCacheDefinitionPart);
                Assert.Equal(100d, rows.GetPivotData("RowsPivot", "Metric").Value);
                Assert.Equal(100d, columns.GetPivotData("ColumnsPivot", "Metric").Value);
                document.GetSheet("Source").CellValue(2, 2, 50d);

                var result = rows.MaterializePivotTable("RowsPivot");

                Assert.True(result.Mutation.PackageIsValid, string.Join(Environment.NewLine,
                    result.Mutation.Diagnostics.Select(diagnostic => diagnostic.Message)));
                Assert.Equal(new[] { "RowsPivot", "ColumnsPivot" }, result.AffectedPivotTables);
                Assert.Equal(140d, rows.GetPivotData("RowsPivot", "Metric").Value);
                Assert.Equal(140d, columns.GetPivotData("ColumnsPivot", "Metric").Value);
                Assert.Equal(4U, cache.PivotCacheDefinition!.RecordCount!.Value);
                Assert.Equal(4U, cache.PivotTableCacheRecordsPart!.PivotCacheRecords!.Count!.Value);
                document.Save();
            }
            using var reopened = ExcelDocument.Load(path);
            Assert.Equal(140d, reopened.GetSheet("Rows").GetPivotData("RowsPivot", "Metric").Value);
            Assert.Equal(140d, reopened.GetSheet("Columns").GetPivotData("ColumnsPivot", "Metric").Value);
            Assert.Empty(reopened.ValidateOpenXml());
        }

        [Fact]
        public void Test_PivotSharedCacheMaterialization_CombinedBudgetPreservesBothViews() {
            using var document = ExcelDocument.Load(PivotSharedCacheOraclePath);
            ExcelSheet rows = document.GetSheet("Rows");
            ExcelSheet columns = document.GetSheet("Columns");
            PivotTableCacheDefinitionPart cache = rows.WorksheetPart.PivotTableParts.Single().PivotTableCacheDefinitionPart!;
            string rowBefore = rows.WorksheetPart.Worksheet.OuterXml;
            string columnBefore = columns.WorksheetPart.Worksheet.OuterXml;
            string cacheBefore = cache.PivotCacheDefinition!.OuterXml;
            document.GetSheet("Source").CellValue(2, 2, 50d);

            Assert.Throws<InvalidOperationException>(() => rows.MaterializePivotTable("RowsPivot",
                new ExcelMutationPlanOptions { MaximumAffectedCells = 9 }));

            Assert.Equal(rowBefore, rows.WorksheetPart.Worksheet.OuterXml);
            Assert.Equal(columnBefore, columns.WorksheetPart.Worksheet.OuterXml);
            Assert.Equal(cacheBefore, cache.PivotCacheDefinition.OuterXml);
        }

        [Fact]
        public void Test_PivotSharedCacheMaterialization_UnsupportedSiblingPreservesCacheAndViews() {
            using var document = ExcelDocument.Load(PivotSharedCacheOraclePath);
            ExcelSheet rows = document.GetSheet("Rows");
            ExcelSheet columns = document.GetSheet("Columns");
            PivotTableCacheDefinitionPart cache = rows.WorksheetPart.PivotTableParts.Single().PivotTableCacheDefinitionPart!;
            columns.WorksheetPart.PivotTableParts.Single().PivotTableDefinition!.DataFields!
                .Elements<DataField>().Single().ShowDataAs = ShowDataAsValues.PercentOfTotal;
            string rowBefore = rows.WorksheetPart.Worksheet.OuterXml;
            string columnBefore = columns.WorksheetPart.Worksheet.OuterXml;
            string cacheBefore = cache.PivotCacheDefinition!.OuterXml;
            document.GetSheet("Source").CellValue(2, 2, 50d);

            Assert.Throws<NotSupportedException>(() => rows.MaterializePivotTable("RowsPivot"));

            Assert.Equal(rowBefore, rows.WorksheetPart.Worksheet.OuterXml);
            Assert.Equal(columnBefore, columns.WorksheetPart.Worksheet.OuterXml);
            Assert.Equal(cacheBefore, cache.PivotCacheDefinition.OuterXml);
        }
    }
}
