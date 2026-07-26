using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_StructuralRows_PreflightsDataTableInputReferences() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellAt(1, 2).SetValue(1);
            Cell owner = sheet.WorksheetPart.Worksheet.Descendants<Cell>()
                .Single(cell => cell.CellReference?.Value == "B1");
            owner.CellFormula = new CellFormula {
                FormulaType = CellFormulaValues.DataTable,
                Reference = "B1:B2",
                R1 = $"A{A1.MaxRows}"
            };

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(
                () => sheet.InsertRows(1));

            Assert.Contains("row limit", exception.Message);
            Assert.Equal($"A{A1.MaxRows}", owner.CellFormula.R1!.Value);
            Assert.Equal("B1", owner.CellReference!.Value);
        }

        [Fact]
        public void Test_StructuralRows_RemovesEmptyMergeCollection() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellAt(2, 1).SetValue(1);
            sheet.MergeRange("A2:B2");

            sheet.DeleteRows(2);

            Assert.Null(sheet.WorksheetPart.Worksheet.GetFirstChild<MergeCells>());
            Assert.Empty(document.ValidateOpenXml());
        }

        [Fact]
        public void Test_StructuralRows_RejectsOleObjectsAtomically() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellAt(2, 1).SetValue(1);
            sheet.WorksheetPart.Worksheet.Append(
                new OleObjects(
                    new OleObject { ShapeId = 1025U }));

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(
                () => sheet.InsertRows(2));

            Assert.Contains("OLE", exception.Message);
            Assert.Equal(1, sheet.CellAt(2, 1).GetValue<int>());
        }

        [Fact]
        public void Test_StructuralRows_RejectsSingleCellXmlMappingsAtomically() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellAt(2, 1).SetValue(1);
            sheet.WorksheetPart.AddNewPart<SingleCellTablePart>();

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(
                () => sheet.DeleteRows(2));

            Assert.Contains("single-cell XML mappings", exception.Message);
            Assert.Equal(1, sheet.CellAt(2, 1).GetValue<int>());
        }

        [Fact]
        public void Test_StructuralRows_PreservesMovedCommentVmlShape() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellAt(5, 2).SetValue(1);
            sheet.SetComment(5, 2, "Keep formatting", author: "Tester");

            VmlDrawingPart vmlPart = Assert.Single(sheet.WorksheetPart.VmlDrawingParts);
            XDocument vml;
            using (Stream stream = vmlPart.GetStream()) {
                vml = XDocument.Load(stream);
            }
            XNamespace v = "urn:schemas-microsoft-com:vml";
            XNamespace x = "urn:schemas-microsoft-com:office:excel";
            XElement shape = Assert.Single(vml.Root!.Elements(v + "shape"));
            string shapeId = shape.Attribute("id")!.Value;
            const string customStyle =
                "position:absolute;margin-left:9pt;margin-top:7pt;width:160pt;height:81pt;z-index:4;visibility:visible";
            shape.SetAttributeValue("style", customStyle);
            shape.SetAttributeValue("fillcolor", "#123456");
            XElement clientData = shape.Element(x + "ClientData")!;
            clientData.SetElementValue(x + "Anchor", "1, 11, 3, 22, 4, 33, 6, 44");
            clientData.Add(new XElement(x + "Visible"));
            using (Stream stream = vmlPart.GetStream(FileMode.Create, FileAccess.Write)) {
                vml.Save(stream);
            }

            sheet.InsertRows(3, 2);

            XDocument shiftedVml;
            using (Stream stream = vmlPart.GetStream()) {
                shiftedVml = XDocument.Load(stream);
            }
            XElement shiftedShape = Assert.Single(shiftedVml.Root!.Elements(v + "shape"));
            XElement shiftedClientData = shiftedShape.Element(x + "ClientData")!;
            Assert.Equal(shapeId, shiftedShape.Attribute("id")!.Value);
            Assert.Equal(customStyle, shiftedShape.Attribute("style")!.Value);
            Assert.Equal("#123456", shiftedShape.Attribute("fillcolor")!.Value);
            Assert.Equal("6", shiftedClientData.Element(x + "Row")!.Value);
            Assert.Equal("1", shiftedClientData.Element(x + "Column")!.Value);
            Assert.Equal("1, 11, 5, 22, 4, 33, 8, 44", shiftedClientData.Element(x + "Anchor")!.Value);
            Assert.NotNull(shiftedClientData.Element(x + "Visible"));
            Assert.True(sheet.HasComment(7, 2));
        }
    }
}
