using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using C15 = DocumentFormat.OpenXml.Office2013.Drawing.Chart;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_StructuralRows_RewritesClassicChartExtensionFormulasOnOtherSheets() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet data = document.AddWorksheet("Data");
            ExcelSheet summary = document.AddWorksheet("Summary");
            data.CellAt(5, 1).SetValue(1);
            summary.CellAt(1, 1).SetValue("Category");
            summary.CellAt(1, 2).SetValue("Value");
            summary.CellAt(2, 1).SetValue("One");
            summary.CellAt(2, 2).SetValue(1);
            summary.AddChartFromRange("A1:B2", row: 5, column: 4);
            ChartPart chartPart = Assert.Single(summary.WorksheetPart.DrawingsPart!.ChartParts);
            var extensionFormula = new C15.Formula("Data!A5:A6");
            chartPart.ChartSpace.Append(
                new C.ExtensionList(
                    new C.Extension(
                        new C15.DataLabelsRange(extensionFormula)) {
                        Uri = "{02D57815-91ED-43cb-92C2-25804820EDAC}"
                    }));

            data.InsertRows(5);

            Assert.Equal("Data!A6:A7", extensionFormula.Text);
        }

        [Fact]
        public void Test_StructuralRows_RejectsCrossSheetFormControlLinksAtomically() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet data = document.AddWorksheet("Data");
            ExcelSheet summary = document.AddWorksheet("Summary");
            data.CellAt(5, 1).SetValue("Keep");
            VmlDrawingPart vmlPart = summary.WorksheetPart.AddNewPart<VmlDrawingPart>();
            string relationshipId = summary.WorksheetPart.GetIdOfPart(vmlPart);
            summary.WorksheetPart.Worksheet.Append(new LegacyDrawing { Id = relationshipId });
            XNamespace v = "urn:schemas-microsoft-com:vml";
            XNamespace x = "urn:schemas-microsoft-com:office:excel";
            var vml = new XDocument(
                new XElement(v + "xml",
                    new XElement(v + "shape",
                        new XElement(x + "ClientData",
                            new XAttribute("ObjectType", "Checkbox"),
                            new XElement(x + "FmlaLink", "Data!A5"),
                            new XElement(x + "Row", "0"),
                            new XElement(x + "Column", "0")))));
            using (Stream stream = vmlPart.GetStream(FileMode.Create, FileAccess.Write)) {
                vml.Save(stream);
            }

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(
                () => data.InsertRows(5));

            Assert.Contains("cross-sheet links", exception.Message);
            Assert.Equal("Keep", data.CellAt(5, 1).GetValue<string>());
        }

        [Fact]
        public void Test_StructuralRows_RemovesScenariosWhenResultRangeIsDeleted() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellAt(1, 2).SetValue(1);
            sheet.CellAt(5, 1).SetValue(2);
            var scenario = new Scenario(
                new InputCells { CellReference = "B1", Val = "1" }) {
                Name = "Survives",
                Count = 1U
            };
            sheet.WorksheetPart.Worksheet.Append(
                new Scenarios(scenario) {
                    SequenceOfReferences = new ListValue<StringValue> { InnerText = "A5" }
                });

            sheet.DeleteRows(5);

            Assert.Null(sheet.WorksheetPart.Worksheet.GetFirstChild<Scenarios>());
            Assert.Empty(document.ValidateOpenXml());
        }

        [Fact]
        public void Test_StructuralRows_MaterializesOverflowedSharedFormulaFollowersAsRefErrors() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet data = document.AddWorksheet("Data");
            ExcelSheet other = document.AddWorksheet("Other");
            other.CellAt(1, 1).SetValue("Delete");
            SheetData sheetData = data.WorksheetPart.Worksheet.GetFirstChild<SheetData>()!;
            var ownerRow = new Row { RowIndex = (uint)(A1.MaxRows - 1) };
            ownerRow.Append(
                new Cell {
                    CellReference = $"A{A1.MaxRows - 1}",
                    CellFormula = new CellFormula($"B{A1.MaxRows}") {
                        FormulaType = CellFormulaValues.Shared,
                        SharedIndex = 91U,
                        Reference = $"A{A1.MaxRows - 1}:A{A1.MaxRows}"
                    }
                });
            var followerRow = new Row { RowIndex = (uint)A1.MaxRows };
            followerRow.Append(
                new Cell {
                    CellReference = $"A{A1.MaxRows}",
                    CellFormula = new CellFormula {
                        FormulaType = CellFormulaValues.Shared,
                        SharedIndex = 91U
                    }
                });
            sheetData.Append(ownerRow, followerRow);

            other.DeleteRows(1);

            CellFormula follower = followerRow.GetFirstChild<Cell>()!.CellFormula!;
            Assert.Equal("#REF!", follower.Text);
            Assert.Null(follower.FormulaType);
            Assert.Null(follower.SharedIndex);
        }
    }
}
