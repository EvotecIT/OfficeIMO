using System;
using System.Linq;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_TableSchema_RenamesResizesAndRewritesStructuredReferences() {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddSheet("Data");
            sheet.Cell(1, 1).SetValue("Region");
            sheet.Cell(1, 2).SetValue("Amount");
            sheet.Cell(2, 1).SetValue("EU");
            sheet.Cell(2, 2).SetValue(10);
            sheet.AddTable("A1:B2", true, "Sales", ExcelTableStyle.TableStyleMedium2);
            sheet.Cell(4, 1).SetFormula("SUM(Sales[Amount])");

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
            ExcelSheet sheet = document.AddSheet("Data");
            sheet.Cell(1, 1).SetValue("A");
            sheet.Cell(1, 2).SetValue("B");
            sheet.Cell(2, 1).SetValue(1);
            sheet.Cell(2, 2).SetValue(2);
            sheet.AddTable("A1:B2", true, "DataTable", ExcelTableStyle.TableStyleMedium2);

            Assert.Throws<ArgumentException>(() =>
                sheet.SetTableSchema("DataTable", new[] { "Same", "same" }));
            Assert.Equal(new[] { "A", "B" }, Assert.Single(document.GetTables()).Columns.Select(column => column.Name));
        }

        [Fact]
        public void Test_TableSchema_SwapsColumnNamesWithoutCascadingStructuredReferences() {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddSheet("Data");
            sheet.Cell(1, 1).SetValue("A");
            sheet.Cell(1, 2).SetValue("B");
            sheet.Cell(2, 1).SetValue(1);
            sheet.Cell(2, 2).SetValue(2);
            sheet.AddTable("A1:B2", true, "DataTable", ExcelTableStyle.TableStyleMedium2);
            sheet.Cell(4, 1).SetFormula("DataTable[A]+DataTable[B]");

            sheet.SetTableSchema("DataTable", new[] { "B", "A" });

            Assert.Equal("DataTable[B]+DataTable[A]", sheet.GetFormulaCells().Single().Formula);
        }

        [Fact]
        public void Test_TableSchema_RejectsResizeOverlappingAnotherTable() {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddSheet("Data");
            sheet.Cell(1, 1).SetValue("A");
            sheet.Cell(2, 1).SetValue(1);
            sheet.Cell(1, 3).SetValue("B");
            sheet.Cell(2, 3).SetValue(2);
            sheet.AddTable("A1:A2", true, "First", ExcelTableStyle.TableStyleMedium2);
            sheet.AddTable("C1:C2", true, "Second", ExcelTableStyle.TableStyleMedium2);

            Assert.Throws<InvalidOperationException>(() => sheet.SetTableSchema("First", new[] { "A", "New", "B" }, "A1:C2"));
            Assert.Equal("A1:A2", document.GetTables().Single(table => table.Name == "First").Range);
        }

        [Fact]
        public void Test_TableSchema_ShrinkInvalidatesRemovedStructuredColumn() {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddSheet("Data");
            sheet.Cell(1, 1).SetValue("A");
            sheet.Cell(1, 2).SetValue("B");
            sheet.Cell(2, 1).SetValue(1);
            sheet.Cell(2, 2).SetValue(2);
            sheet.AddTable("A1:B2", true, "Sales", ExcelTableStyle.TableStyleMedium2);
            sheet.Cell(4, 1).SetFormula("SUM(Sales[B])");

            sheet.ResizeTable("Sales", "A1:A2");

            Assert.Equal("SUM(#REF!)", Assert.Single(sheet.GetFormulaCells()).Formula);
        }

        [Fact]
        public void Test_TableSchema_RenamesRowContextStructuredReference() {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddSheet("Data");
            sheet.Cell(1, 1).SetValue("Amount");
            sheet.Cell(2, 1).SetValue(1);
            sheet.AddTable("A1:A2", true, "Sales", ExcelTableStyle.TableStyleMedium2);
            sheet.Cell(4, 1).SetFormula("Sales[@Amount]*2");

            sheet.SetTableSchema("Sales", new[] { "Net" });

            Assert.Equal("Sales[@Net]*2", Assert.Single(sheet.GetFormulaCells()).Formula);
        }

        [Fact]
        public void Test_TableSchema_ExpansionGeneratesUnusedColumnName() {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddSheet("Data");
            sheet.Cell(1, 1).SetValue("Column2");
            sheet.Cell(2, 1).SetValue(1);
            sheet.AddTable("A1:A2", true, "Sales", ExcelTableStyle.TableStyleMedium2);

            sheet.ResizeTable("Sales", "A1:B2");

            Assert.Equal(new[] { "Column2", "Column3" }, Assert.Single(document.GetTables()).Columns.Select(column => column.Name));
        }

        [Fact]
        public void Test_TableSchema_RewritesAndInvalidatesEscapedBracketColumnNames() {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddSheet("Data");
            sheet.Cell(1, 1).SetValue("Keep");
            sheet.Cell(1, 2).SetValue("Cost [ old");
            sheet.Cell(2, 1).SetValue(1);
            sheet.Cell(2, 2).SetValue(2);
            sheet.AddTable("A1:B2", true, "Table1", ExcelTableStyle.TableStyleMedium2);
            sheet.Cell(4, 1).SetFormula("SUM(Table1[Cost '[ old])");

            ExcelFormulaStructuredReferenceSyntax parsed = Assert.Single(
                ExcelFormulaSyntaxTree.Parse("Table1[Cost '[ old]").Nodes
                    .OfType<ExcelFormulaStructuredReferenceSyntax>());
            Assert.Equal("[Cost '[ old]", parsed.Selector);

            sheet.SetTableSchema("Table1", new[] { "Keep", "Cost ] new" });
            Assert.Equal("SUM(Table1[Cost '] new])", Assert.Single(sheet.GetFormulaCells()).Formula);

            sheet.ResizeTable("Table1", "A1:A2");
            Assert.Equal("SUM(#REF!)", Assert.Single(sheet.GetFormulaCells()).Formula);
        }
    }
}
