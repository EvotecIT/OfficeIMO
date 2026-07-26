using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_StructuralRows_PreservesBackslashNamesInSharedFormulas() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            SheetData sheetData = sheet.WorksheetPart.Worksheet.GetFirstChild<SheetData>()!;
            var ownerRow = new Row { RowIndex = 2U };
            ownerRow.Append(new Cell {
                CellReference = "A2",
                CellFormula = new CellFormula(@"\A5+1") {
                    FormulaType = CellFormulaValues.Shared,
                    SharedIndex = 101U,
                    Reference = "A2:A3"
                }
            });
            var followerRow = new Row { RowIndex = 3U };
            followerRow.Append(new Cell {
                CellReference = "A3",
                CellFormula = new CellFormula {
                    FormulaType = CellFormulaValues.Shared,
                    SharedIndex = 101U
                }
            });
            sheetData.Append(ownerRow, followerRow);

            sheet.InsertRows(6);

            CellFormula[] formulas = sheet.WorksheetPart.Worksheet
                .Descendants<CellFormula>()
                .ToArray();
            Assert.Equal(new[] { @"\A5+1", @"\A5+1" }, formulas.Select(item => item.Text));
        }

        [Fact]
        public void Test_StructuralRows_InvalidatesEveryWrapperForEditedWorksheet() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet retained = document.AddWorksheet("Data");
            retained.CellAt(5, 1).SetValue("Five");
            retained.CellAt(6, 1).SetValue("Six");
            Assert.Equal("Five", retained.CellAt(5, 1).GetValue<string>());
            ExcelSheet editor = document["Data"];
            Assert.NotSame(retained, editor);

            editor.InsertRows(5);

            retained.CellAt(5, 1).SetValue("New");
            Assert.Equal("New", editor.CellAt(5, 1).GetValue<string>());
            Assert.Equal("Five", retained.CellAt(6, 1).GetValue<string>());
            Assert.Equal("Six", retained.CellAt(7, 1).GetValue<string>());
        }

        [Fact]
        public void Test_StructuralRows_RenumbersSelectionActiveCellId() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            var selection = new Selection {
                ActiveCell = "B5",
                ActiveCellId = 1U,
                SequenceOfReferences = new ListValue<StringValue> {
                    InnerText = "A2 B5"
                }
            };
            var views = new SheetViews(
                new SheetView(selection) { WorkbookViewId = 0U });
            SheetData sheetData = sheet.WorksheetPart.Worksheet.GetFirstChild<SheetData>()!;
            sheet.WorksheetPart.Worksheet.InsertBefore(views, sheetData);

            sheet.DeleteRows(2);

            Assert.Equal("B4", selection.ActiveCell!.Value);
            Assert.Equal("B4", selection.SequenceOfReferences!.InnerText);
            Assert.Equal(0U, selection.ActiveCellId!.Value);
        }

        [Fact]
        public void Test_StructuralRows_InvalidatesInvariantDynamicNameChartCaches() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellAt(1, 1).SetValue("Category");
            sheet.CellAt(1, 2).SetValue("Value");
            sheet.CellAt(2, 1).SetValue("One");
            sheet.CellAt(2, 2).SetValue(1);
            sheet.CellAt(3, 1).SetValue("Two");
            sheet.CellAt(3, 2).SetValue(2);
            document.WorkbookRoot.DefinedNames = new DefinedNames(
                new DefinedName("OFFSET(Data!$A$1,0,0,COUNTA(Data!$A:$A),1)") {
                    Name = "SeriesData"
                });
            sheet.AddChartFromRange("A1:B3", row: 5, column: 4);
            ChartPart chartPart = Assert.Single(sheet.WorksheetPart.DrawingsPart!.ChartParts);
            C.Formula formula = chartPart.ChartSpace.Descendants<C.Formula>()
                .First(item => item.Parent!.ChildElements.Any(element =>
                    element.LocalName.EndsWith("Cache", System.StringComparison.OrdinalIgnoreCase)));
            formula.Text = "SeriesData";
            Assert.Contains(
                formula.Parent!.ChildElements,
                element => element.LocalName.EndsWith("Cache", System.StringComparison.OrdinalIgnoreCase));

            sheet.DeleteRows(3);

            Assert.Equal("SeriesData", formula.Text);
            Assert.DoesNotContain(
                formula.Parent!.ChildElements,
                element => element.LocalName.EndsWith("Cache", System.StringComparison.OrdinalIgnoreCase));
        }
    }
}
