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
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using Xm = DocumentFormat.OpenXml.Office.Excel;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_StructuralRows_PreflightsUnqualifiedLocalChartFormulas() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellAt(1, 1).SetValue("Category");
            sheet.CellAt(1, 2).SetValue("Value");
            sheet.CellAt(2, 1).SetValue("One");
            sheet.CellAt(2, 2).SetValue(1);
            sheet.AddChartFromRange("A1:B2", row: 5, column: 4);
            ChartPart chartPart = Assert.Single(sheet.WorksheetPart.DrawingsPart!.ChartParts);
            C.Formula formula = chartPart.ChartSpace.Descendants<C.Formula>().First();
            formula.Text = $"A{A1.MaxRows}";

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(
                () => sheet.InsertRows(1));

            Assert.Contains("row limit", exception.Message);
            Assert.Equal($"A{A1.MaxRows}", formula.Text);
            Assert.Equal("Category", sheet.CellAt(1, 1).GetValue<string>());
        }

        [Fact]
        public void Test_StructuralRows_DoesNotPreflightStationaryOneCellToMarker() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellAt(1, 1).SetValue("Category");
            sheet.CellAt(1, 2).SetValue("Value");
            sheet.CellAt(2, 1).SetValue("One");
            sheet.CellAt(2, 2).SetValue(1);
            sheet.AddChartFromRange("A1:B2", row: 5, column: 4);

            Xdr.WorksheetDrawing drawing = sheet.WorksheetPart.DrawingsPart!.WorksheetDrawing!;
            Xdr.OneCellAnchor oneCellAnchor = Assert.Single(drawing.Elements<Xdr.OneCellAnchor>());
            var oneCellPlacement = new Xdr.TwoCellAnchor(
                new Xdr.FromMarker(
                    new Xdr.ColumnId("3"),
                    new Xdr.ColumnOffset("0"),
                    new Xdr.RowId("0"),
                    new Xdr.RowOffset("0")),
                new Xdr.ToMarker(
                    new Xdr.ColumnId("8"),
                    new Xdr.ColumnOffset("0"),
                    new Xdr.RowId((A1.MaxRows - 1).ToString()),
                    new Xdr.RowOffset("0")),
                oneCellAnchor.GetFirstChild<Xdr.GraphicFrame>()!.CloneNode(true),
                new Xdr.ClientData()) {
                EditAs = Xdr.EditAsValues.OneCell
            };
            oneCellAnchor.Remove();
            drawing.Append(oneCellPlacement);

            sheet.InsertRows(3);

            Assert.Equal("0", oneCellPlacement.FromMarker!.RowId!.Text);
            Assert.Equal((A1.MaxRows - 1).ToString(), oneCellPlacement.ToMarker!.RowId!.Text);
        }

        [Fact]
        public void Test_StructuralRows_RemapsStationaryCommentVmlBoxAnchor() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.SetComment(1, 1, "Keep", author: "Tester");
            VmlDrawingPart vmlPart = Assert.Single(sheet.WorksheetPart.VmlDrawingParts);
            XNamespace v = "urn:schemas-microsoft-com:vml";
            XNamespace x = "urn:schemas-microsoft-com:office:excel";
            XDocument vml;
            using (Stream stream = vmlPart.GetStream()) {
                vml = XDocument.Load(stream);
            }
            XElement clientData = Assert.Single(vml.Root!.Elements(v + "shape"))
                .Element(x + "ClientData")!;
            clientData.SetElementValue(x + "Anchor", "0, 0, 0, 0, 2, 0, 3, 0");
            using (Stream stream = vmlPart.GetStream(FileMode.Create, FileAccess.Write)) {
                vml.Save(stream);
            }

            sheet.InsertRows(3);

            using Stream shiftedStream = vmlPart.GetStream();
            XDocument shifted = XDocument.Load(shiftedStream);
            XElement shiftedClientData = Assert.Single(shifted.Root!.Elements(v + "shape"))
                .Element(x + "ClientData")!;
            Assert.Equal("0, 0, 0, 0, 2, 0, 4, 0", shiftedClientData.Element(x + "Anchor")!.Value);
            Assert.Equal("0", shiftedClientData.Element(x + "Row")!.Value);
            Assert.True(sheet.HasComment(1, 1));
        }

        [Fact]
        public void Test_StructuralRows_RemapsSparklineGroupDateAxisFormulaWorkbookWide() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet data = document.AddWorksheet("Data");
            ExcelSheet summary = document.AddWorksheet("Summary");
            data.CellAt(5, 1).SetValue(1);
            data.CellAt(6, 1).SetValue(2);
            summary.AddSparklines("'Data'!A5:A6", "A1");
            X14.SparklineGroup group = Assert.Single(
                summary.WorksheetPart.Worksheet.Descendants<X14.SparklineGroup>());
            group.Formula = new Xm.Formula("'Data'!A5:A6");

            data.InsertRows(5);

            Assert.Equal("'Data'!A6:A7", group.Formula!.Text);
        }

        [Fact]
        public void Test_StructuralRows_RejectsMacroSheetsAtomically() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellAt(2, 1).SetValue("Keep");
            document.WorkbookPartRoot.AddNewPart<MacroSheetPart>();

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(
                () => sheet.DeleteRows(2));

            Assert.Contains("macro sheets", exception.Message);
            Assert.Equal("Keep", sheet.CellAt(2, 1).GetValue<string>());
        }

        [Fact]
        public void Test_StructuralRows_RejectsLegacyRevisionTrackingAtomically() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellAt(2, 1).SetValue("Keep");
            document.WorkbookPartRoot.AddNewPart<WorkbookRevisionHeaderPart>();

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(
                () => sheet.InsertRows(2));

            Assert.Contains("revision tracking", exception.Message);
            Assert.Equal("Keep", sheet.CellAt(2, 1).GetValue<string>());
        }

        [Fact]
        public void Test_StructuralRows_RewritesSheetNamesContainingDoubleQuotes() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet data = document.AddWorksheet("Data\"Set");
            ExcelSheet summary = document.AddWorksheet("Summary");
            data.CellAt(5, 1).SetValue(1);
            summary.CellFormula(1, 1, "'Data\"Set'!A5");

            data.InsertRows(5);

            Cell formulaCell = summary.WorksheetPart.Worksheet.Descendants<Cell>()
                .Single(cell => cell.CellReference?.Value == "A1");
            Assert.Equal("'Data\"Set'!A6", formulaCell.CellFormula!.Text);
        }

        [Fact]
        public void Test_StructuralRows_RewritesClassicChartExtensionFormulas() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellAt(1, 1).SetValue("Category");
            sheet.CellAt(1, 2).SetValue("Value");
            sheet.CellAt(2, 1).SetValue("One");
            sheet.CellAt(2, 2).SetValue(1);
            sheet.AddChartFromRange("A1:B2", row: 5, column: 4);
            ChartPart chartPart = Assert.Single(sheet.WorksheetPart.DrawingsPart!.ChartParts);
            var extensionFormula = new C15.Formula("A5:A6");
            chartPart.ChartSpace.Append(
                new C.ExtensionList(
                    new C.Extension(
                        new C15.DataLabelsRange(extensionFormula)) {
                        Uri = "{02D57815-91ED-43cb-92C2-25804820EDAC}"
                    }));

            sheet.InsertRows(5);

            Assert.Equal("A6:A7", extensionFormula.Text);
        }
    }
}
