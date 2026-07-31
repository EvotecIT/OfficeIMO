using System;
using System.Linq;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_TableSchema_RenamesResizesAndRewritesStructuredReferences() {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellValue(1, 1, "Region");
            sheet.CellValue(1, 2, "Amount");
            sheet.CellValue(2, 1, "EU");
            sheet.CellValue(2, 2, 10);
            sheet.AddTable("A1:B2", true, "Sales", TableStyle.TableStyleMedium2);
            sheet.CellFormula(4, 1, "SUM(Sales[Amount])");

            ExcelTable renamed = sheet.Table("Sales").Rename("Ledger");
            renamed.SetSchema(new[] { "Region", "Net" }, "A1:B3");

            ExcelTableInfo table = Assert.Single(document.GetTables());
            Assert.Equal("Ledger", table.Name);
            Assert.Equal("A1:B3", table.Range);
            Assert.Equal(new[] { "Region", "Net" }, table.Columns.Select(column => column.Name));
            Assert.Equal("SUM(Ledger[Net])", sheet.GetFormulaCells().Single().Formula);
        }

        [Fact]
        public void Test_TableSchema_RejectsDuplicateColumnNamesWithoutMutation() {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellValue(1, 1, "A");
            sheet.CellValue(1, 2, "B");
            sheet.CellValue(2, 1, 1);
            sheet.CellValue(2, 2, 2);
            sheet.AddTable("A1:B2", true, "DataTable", TableStyle.TableStyleMedium2);

            Assert.Throws<ArgumentException>(() =>
                sheet.SetTableSchema("DataTable", new[] { "Same", "same" }));
            Assert.Equal(new[] { "A", "B" }, Assert.Single(document.GetTables()).Columns.Select(column => column.Name));
        }

        [Fact]
        public void Test_TableSchema_SwapsColumnNamesWithoutCascadingStructuredReferences() {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellValue(1, 1, "A");
            sheet.CellValue(1, 2, "B");
            sheet.CellValue(2, 1, 1);
            sheet.CellValue(2, 2, 2);
            sheet.AddTable("A1:B2", true, "DataTable", TableStyle.TableStyleMedium2);
            sheet.CellFormula(4, 1, "DataTable[A]+DataTable[B]");

            sheet.SetTableSchema("DataTable", new[] { "B", "A" });

            Assert.Equal("DataTable[B]+DataTable[A]", sheet.GetFormulaCells().Single().Formula);
        }
    }
}
