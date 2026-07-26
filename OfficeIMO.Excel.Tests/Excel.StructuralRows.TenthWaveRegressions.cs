using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_StructuralRows_TracksTableFormulaTargetsIndependentlyFromTableMovement() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellAt(5, 1).SetValue("Value");
            sheet.CellAt(5, 2).SetValue("Result");
            sheet.CellAt(6, 1).SetValue(10);
            sheet.CellAt(7, 1).SetValue(20);
            sheet.AddTable(
                "A5:B7",
                hasHeader: true,
                name: "CalculatedData",
                OfficeIMO.Excel.TableStyle.TableStyleMedium2);

            Table table = Assert.Single(sheet.WorksheetPart.TableDefinitionParts).Table;
            TableColumn resultColumn = table.Descendants<TableColumn>()
                .Single(column => column.Name?.Value == "Result");
            var formula = new CalculatedColumnFormula("A1+A5");
            resultColumn.Append(formula);

            sheet.InsertRows(3);

            Assert.Equal("A6:B8", table.Reference!.Value);
            Assert.Equal("A1+A6", formula.Text);
        }

        [Fact]
        public void Test_StructuralRows_ShrinksPartiallyDeletedRelativeAnchoredRanges() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellAt(1, 2).SetValue(1);
            sheet.CellAt(2, 1).SetValue(2);
            sheet.CellAt(3, 1).SetValue(3);
            var validation = new DataValidation(
                new Formula1("SUM(A2:A3)>0"),
                new Formula2("SUM($A2:A$3)>0")) {
                SequenceOfReferences = new ListValue<StringValue> { InnerText = "B1" }
            };
            sheet.WorksheetPart.Worksheet.Append(
                new DataValidations(validation) { Count = 1U });

            sheet.DeleteRows(2);

            Assert.Equal("SUM(A2:A2)>0", validation.Formula1!.Text);
            Assert.Equal("SUM($A2:A$2)>0", validation.Formula2!.Text);
        }

        [Fact]
        public void Test_StructuralRows_RenumbersScenarioViewIndicesAfterRemoval() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellAt(2, 1).SetValue(1);
            sheet.CellAt(3, 1).SetValue(2);
            var removed = new Scenario(
                new InputCells { CellReference = "A2", Val = "1" }) {
                Name = "Removed",
                Count = 1U
            };
            var surviving = new Scenario(
                new InputCells { CellReference = "A3", Val = "2" }) {
                Name = "Surviving",
                Count = 1U
            };
            var scenarios = new Scenarios(removed, surviving) {
                Current = 1U,
                Show = 1U
            };
            sheet.WorksheetPart.Worksheet.Append(scenarios);

            sheet.DeleteRows(2);

            Assert.Same(surviving, Assert.Single(scenarios.Elements<Scenario>()));
            Assert.Equal(0U, scenarios.Current!.Value);
            Assert.Equal(0U, scenarios.Show!.Value);
            Assert.Equal("A2", Assert.Single(surviving.Elements<InputCells>()).CellReference!.Value);
            Assert.Empty(document.ValidateOpenXml());
        }

        [Fact]
        public void Test_StructuralRows_RejectsVmlOnlyFormControlsAtomically() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellAt(2, 1).SetValue("Keep");
            VmlDrawingPart vmlPart = sheet.WorksheetPart.AddNewPart<VmlDrawingPart>();
            string relationshipId = sheet.WorksheetPart.GetIdOfPart(vmlPart);
            sheet.WorksheetPart.Worksheet.Append(new LegacyDrawing { Id = relationshipId });
            XNamespace v = "urn:schemas-microsoft-com:vml";
            XNamespace x = "urn:schemas-microsoft-com:office:excel";
            var vml = new XDocument(
                new XElement(v + "xml",
                    new XElement(v + "shape",
                        new XElement(x + "ClientData",
                            new XAttribute("ObjectType", "Checkbox"),
                            new XElement(x + "Row", "1"),
                            new XElement(x + "Column", "0")))));
            using (Stream stream = vmlPart.GetStream(FileMode.Create, FileAccess.Write)) {
                vml.Save(stream);
            }

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(
                () => sheet.DeleteRows(2));

            Assert.Contains("form controls", exception.Message);
            Assert.Equal("Keep", sheet.CellAt(2, 1).GetValue<string>());
            Assert.Single(sheet.WorksheetPart.VmlDrawingParts);
        }

        [Fact]
        public void Test_StructuralRows_RemapsNormalAndCustomViewSelections() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellAt(5, 1).SetValue(1);
            var normalSelection = new Selection {
                ActiveCell = "A5",
                SequenceOfReferences = new ListValue<StringValue> { InnerText = "A5:B6" }
            };
            var sheetViews = new SheetViews(
                new SheetView(normalSelection) { WorkbookViewId = 0U });
            SheetData sheetData = sheet.WorksheetPart.Worksheet.GetFirstChild<SheetData>()!;
            sheet.WorksheetPart.Worksheet.InsertBefore(sheetViews, sheetData);

            var customSelection = new Selection {
                ActiveCell = "C5",
                SequenceOfReferences = new ListValue<StringValue> { InnerText = "C5:D6" }
            };
            var customView = new CustomSheetView(customSelection) {
                Guid = "{3CE8878D-FCA0-4DBA-8A9A-5B34EADBFDBD}"
            };
            sheet.WorksheetPart.Worksheet.Append(new CustomSheetViews(customView));

            sheet.InsertRows(3);

            Assert.Equal("A6", normalSelection.ActiveCell!.Value);
            Assert.Equal("A6:B7", normalSelection.SequenceOfReferences!.InnerText);
            Assert.Equal("C6", customSelection.ActiveCell!.Value);
            Assert.Equal("C6:D7", customSelection.SequenceOfReferences!.InnerText);

            sheet.DeleteRows(6);

            Assert.Equal("A6", normalSelection.ActiveCell!.Value);
            Assert.Equal("A6:B6", normalSelection.SequenceOfReferences!.InnerText);
            Assert.Equal("C6", customSelection.ActiveCell!.Value);
            Assert.Equal("C6:D6", customSelection.SequenceOfReferences!.InnerText);
            Assert.Empty(document.ValidateOpenXml());
        }

        [Fact]
        public void Test_StructuralRows_PreflightsSavedViewSelections() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellAt(1, 1).SetValue("Keep");
            var selection = new Selection {
                ActiveCell = $"A{A1.MaxRows}",
                SequenceOfReferences = new ListValue<StringValue> {
                    InnerText = $"A{A1.MaxRows}"
                }
            };
            var sheetViews = new SheetViews(
                new SheetView(selection) { WorkbookViewId = 0U });
            SheetData sheetData = sheet.WorksheetPart.Worksheet.GetFirstChild<SheetData>()!;
            sheet.WorksheetPart.Worksheet.InsertBefore(sheetViews, sheetData);

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(
                () => sheet.InsertRows(1));

            Assert.Contains("row limit", exception.Message);
            Assert.Equal($"A{A1.MaxRows}", selection.ActiveCell!.Value);
            Assert.Equal("Keep", sheet.CellAt(1, 1).GetValue<string>());
        }
    }
}
