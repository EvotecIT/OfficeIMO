using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_TableResize_RemovesAndClampsSortConditionsOutsideShrunkenSchema() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellValue(1, 1, "A");
            sheet.CellValue(1, 2, "B");
            sheet.CellValue(1, 3, "C");
            sheet.AddTable("A1:C10", true, "Sales", OfficeIMO.Excel.TableStyle.TableStyleMedium2);
            Table table = Assert.Single(sheet.WorksheetPart.TableDefinitionParts).Table!;
            var sortState = new SortState(
                new SortCondition { Reference = "A2:B9" },
                new SortCondition { Reference = "C2:C9" }) {
                Reference = "A1:C10"
            };
            table.Append(sortState);

            sheet.ResizeTable("Sales", "A1:A10");

            Assert.Equal("A1:A10", sortState.Reference!.Value);
            SortCondition condition = Assert.Single(sortState.Elements<SortCondition>());
            Assert.Equal("A2:A9", condition.Reference!.Value);
        }

        [Fact]
        public void Test_ColumnPlan_RejectsFormulaReferenceThatWouldOverflowXfd() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellFormula(1, 1, "XFD1");

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
                sheet.PlanInsertColumns(1));

            Assert.Contains("beyond worksheet limits", exception.Message,
                System.StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Test_ColumnPlan_RejectsDefinedNameThatWouldOverflowXfd() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            document.WorkbookPartRoot.Workbook!.DefinedNames = new DefinedNames(
                new DefinedName { Name = "Edge", Text = "Data!$XFD$1" });

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
                sheet.PlanInsertColumns(1));

            Assert.Contains("beyond worksheet limits", exception.Message,
                System.StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Test_SparklineRemoval_PrunesEmptyExtensionStructuresAndRemainsValid() {
            using var output = new MemoryStream();
            using (var document = ExcelDocument.Create(output)) {
                ExcelSheet sheet = document.AddWorksheet("Data");
                sheet.AddSparklines("A1:C1", "D1");

                Assert.Equal(1, sheet.RemoveSparklines("D1"));
                Assert.Null(sheet.WorksheetPart.Worksheet!.GetFirstChild<ExtensionList>());
                document.Save();
            }

            output.Position = 0;
            using SpreadsheetDocument package = SpreadsheetDocument.Open(output, false);
            Assert.Empty(new OpenXmlValidator().Validate(package));
        }

        [Fact]
        public void Test_SparklineClear_PrunesEmptyExtensionStructures() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.AddSparklines("A1:C2", "D1:D2");

            Assert.Equal(2, sheet.ClearSparklines());

            Assert.Null(sheet.WorksheetPart.Worksheet!.GetFirstChild<ExtensionList>());
        }
    }
}
