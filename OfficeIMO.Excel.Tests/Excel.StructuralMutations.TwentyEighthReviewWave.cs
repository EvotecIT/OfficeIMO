using System;
using System.IO;
using System.Linq;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_ModernChart_RemovedWrapperCannotMutateReusedDataAllocation() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Dashboard");
            ExcelModernChart removed = sheet.AddModernChart(
                new ExcelChartData(
                    new[] { "A" },
                    new[] { new ExcelChartSeries("Value", new[] { 1d }) }),
                1,
                1,
                ExcelModernChartType.Funnel,
                "Removed");
            ExcelChartDataRange releasedRange = removed.DataRange!;
            removed.Remove();

            ExcelModernChart live = sheet.AddModernChart(
                new ExcelChartData(
                    new[] { "B" },
                    new[] { new ExcelChartSeries("Value", new[] { 2d }) }),
                1,
                4,
                ExcelModernChartType.Funnel,
                "Live");
            ExcelChartDataRange liveRange = live.DataRange!;
            Assert.Equal(releasedRange.StartRow, liveRange.StartRow);
            Assert.Equal(releasedRange.StartColumn, liveRange.StartColumn);

            Assert.Throws<InvalidOperationException>(() => removed.UpdateData(new ExcelChartData(
                new[] { "Changed" },
                new[] { new ExcelChartSeries("Value", new[] { 99d }) })));
            Assert.Throws<InvalidOperationException>(() => removed.SetTitle("Changed"));
            Assert.Throws<InvalidOperationException>(() => removed.SetPlacement(2, 2, 300, 200));
            Assert.Throws<InvalidOperationException>(() => removed.Name = "Changed");

            ExcelSheet dataSheet = document[liveRange.SheetName];
            Assert.Equal("B", dataSheet.CellAt(liveRange.CategoryStartRow, liveRange.CategoryStartColumn).GetValue<string>());
            Assert.Equal(2d, dataSheet.CellAt(liveRange.CategoryStartRow, liveRange.SeriesStartColumn).GetValue<double>());
            Assert.Equal("Live", live.Title);
            Assert.Empty(document.ValidateOpenXml());
        }

        [Fact]
        public void Test_FormulaSearch_ExcludesLetAndLambdaBoundCalls() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellFormula(1, 1, "LET(SUM,LAMBDA(x,x),SUM(1))");
            sheet.CellFormula(2, 1, "LAMBDA(SUM,SUM(1))(LAMBDA(x,x))");
            sheet.CellFormula(3, 1, "SUM(1)");

            ExcelFormulaCellInfo match = Assert.Single(sheet.SearchFormulas(
                new ExcelFormulaSearchOptions { Function = "SUM" }));

            Assert.Equal("A3", match.CellReference);
        }

        [Fact]
        public void Test_StructuralColumns_SynchronizeQueryTableFields() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            ExcelQueryBackedTableInfo source = document.AddQueryBackedTable(new ExcelQueryBackedTableOptions {
                ConnectionName = "StructuralQuery",
                WorksheetName = sheet.Name,
                TableName = "StructuralResults",
                ColumnNames = new[] { "A", "B" }
            });

            sheet.InsertColumns(2);
            Assert.Equal("A1:C1", sheet.GetTableRange(source.TableName));
            AssertQueryFieldSchema(sheet, new[] { "A", "Column2", "B" });

            sheet.DeleteColumns(2);
            Assert.Equal("A1:B1", sheet.GetTableRange(source.TableName));
            AssertQueryFieldSchema(sheet, new[] { "A", "B" });
            Assert.Empty(document.ValidateOpenXml());
        }
    }
}
