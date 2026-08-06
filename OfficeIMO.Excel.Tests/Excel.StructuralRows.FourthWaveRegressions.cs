using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using Xm = DocumentFormat.OpenXml.Office.Excel;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_StructuralRows_RejectsPartialDeletionOfDataTableOwner() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellAt(2, 3).SetValue(10);
            sheet.CellAt(3, 3).SetValue(20);
            sheet.CellAt(4, 3).SetValue(30);

            Cell owner = sheet.WorksheetPart.Worksheet.Descendants<Cell>()
                .Single(cell => cell.CellReference?.Value == "C2");
            owner.CellFormula = new CellFormula {
                FormulaType = CellFormulaValues.DataTable,
                Reference = "C2:C4",
                R1 = "A1"
            };

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(
                () => sheet.DeleteRows(2));

            Assert.Contains("data-table", exception.Message);
            Assert.Equal("C2:C4", owner.CellFormula.Reference!.Value);
            Assert.Equal(20, sheet.CellAt(3, 3).GetValue<int>());
        }

        [Fact]
        public void Test_StructuralRows_RewritesCrossSheetOffice2010ConditionalFormattingFormulas() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet data = document.AddWorksheet("Data");
            ExcelSheet summary = document.AddWorksheet("Summary");
            data.CellAt(3, 1).SetValue(1);

            var formula = new Xm.Formula("Data!A3>0");
            var rule = new X14.ConditionalFormattingRule(formula) {
                Type = ConditionalFormatValues.Expression,
                Priority = 1,
                Id = "{9BBD0D84-319F-4BF5-8812-B855645AC843}"
            };
            var formatting = new X14.ConditionalFormatting(
                rule,
                new Xm.ReferenceSequence("C2"));
            summary.WorksheetPart.Worksheet.Append(
                new ExtensionList(
                    new Extension(new X14.ConditionalFormattings(formatting)) {
                        Uri = "{78C0D931-6437-407D-A8EE-F0AAD7539E65}"
                    }));

            data.InsertRows(3);

            Assert.Equal("Data!A4>0", formula.Text);
            Assert.Equal("C2", formatting.GetFirstChild<Xm.ReferenceSequence>()!.Text);
        }

        [Fact]
        public void Test_StructuralRows_RebasesCalculatedTableFormulaWithItsDataAnchor() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellAt(1, 1).SetValue("Value");
            sheet.CellAt(1, 2).SetValue("Result");
            sheet.CellAt(2, 1).SetValue(10);
            sheet.CellAt(3, 1).SetValue(20);
            sheet.AddTable(
                "A1:B3",
                hasHeader: true,
                name: "CalculatedData",
                OfficeIMO.Excel.ExcelTableStyle.TableStyleMedium2);

            Table table = Assert.Single(sheet.WorksheetPart.TableDefinitionParts).Table;
            TableColumn resultColumn = table.Descendants<TableColumn>()
                .Single(column => column.Name?.Value == "Result");
            var formula = new CalculatedColumnFormula("A2*2");
            resultColumn.Append(formula);

            sheet.DeleteRows(2);

            Assert.Equal("A2*2", formula.Text);
            Assert.Equal("A1:B2", table.Reference!.Value);
        }

        [Fact]
        public void Test_StructuralRows_NormalizesImplicitRowIndicesForShiftAndOverflow() {
            using (var document = ExcelDocument.Create(new MemoryStream())) {
                ExcelSheet sheet = document.AddWorksheet("Data");
                sheet.CellAt(1, 1).SetValue("Move");
                Row row = Assert.Single(sheet.WorksheetPart.Worksheet.Descendants<Row>());
                row.RowIndex = null;

                sheet.InsertRows(1);

                Assert.Equal(2U, row.RowIndex!.Value);
                Assert.Equal("A2", Assert.Single(row.Elements<Cell>()).CellReference!.Value);
                Assert.Equal("Move", sheet.CellAt(2, 1).GetValue<string>());
            }

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                ExcelSheet sheet = document.AddWorksheet("Data");
                sheet.CellAt(A1.MaxRows, 1).SetValue("Boundary");
                Row row = Assert.Single(sheet.WorksheetPart.Worksheet.Descendants<Row>());
                row.RowIndex = null;

                Assert.Throws<InvalidOperationException>(() => sheet.InsertRows(1));

                Assert.Null(row.RowIndex);
                Assert.Equal($"A{A1.MaxRows}", Assert.Single(row.Elements<Cell>()).CellReference!.Value);
            }
        }

        [Fact]
        public void Test_StructuralRows_RejectsR1C1ReferenceModeBeforeMutation() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellAt(3, 1).SetValue("Keep");
            document.WorkbookRoot.Append(new CalculationProperties {
                ReferenceMode = ReferenceModeValues.R1C1
            });

            InvalidOperationException insert = Assert.Throws<InvalidOperationException>(
                () => sheet.InsertRows(3));
            InvalidOperationException delete = Assert.Throws<InvalidOperationException>(
                () => sheet.DeleteRows(3));

            Assert.Contains("R1C1", insert.Message);
            Assert.Contains("R1C1", delete.Message);
            Assert.Equal("Keep", sheet.CellAt(3, 1).GetValue<string>());
        }

        [Fact]
        public void Test_StructuralRows_RemapsScenarioResultAndInputReferences() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellAt(5, 1).SetValue(1);
            sheet.CellAt(6, 1).SetValue(2);

            var input = new InputCells { CellReference = "A5", Val = "10" };
            var scenario = new Scenario(input) { Name = "Best", Count = 1U };
            var scenarios = new Scenarios(scenario) {
                SequenceOfReferences = new ListValue<StringValue> { InnerText = "A5:A6" }
            };
            sheet.WorksheetPart.Worksheet.Append(scenarios);

            sheet.InsertRows(5);

            Assert.Equal("A6:A7", scenarios.SequenceOfReferences!.InnerText);
            Assert.Equal("A6", input.CellReference!.Value);

            sheet.DeleteRows(5);

            Assert.Equal("A5:A6", scenarios.SequenceOfReferences!.InnerText);
            Assert.Equal("A5", input.CellReference!.Value);
            Assert.Equal(1U, scenario.Count!.Value);
        }

        [Fact]
        public void Test_StructuralRows_RemapsCustomSheetViewAutoFilters() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellAt(5, 1).SetValue("Header");
            sheet.CellAt(6, 1).SetValue("Value");

            var filter = new AutoFilter { Reference = "A5:B6" };
            var view = new CustomSheetView(filter) {
                Guid = "{2FC474E2-3AF8-43D5-93AC-5AF7D9A41923}"
            };
            sheet.WorksheetPart.Worksheet.Append(new CustomSheetViews(view));

            sheet.InsertRows(5);

            Assert.Equal("A6:B7", filter.Reference!.Value);

            sheet.DeleteRows(5);

            Assert.Equal("A5:B6", filter.Reference!.Value);
        }
    }
}
