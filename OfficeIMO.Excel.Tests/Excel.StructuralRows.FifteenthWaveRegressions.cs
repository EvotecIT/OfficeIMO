using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;
using Cx = DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_StructuralRows_RemapsCustomSheetViewRowBreaks() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            var pageBreak = new Break { Id = 5U, ManualPageBreak = true };
            var rowBreaks = new RowBreaks(pageBreak) {
                Count = 1U,
                ManualBreakCount = 1U
            };
            sheet.WorksheetPart.Worksheet.Append(new CustomSheetViews(
                new CustomSheetView(rowBreaks) {
                    Guid = "{3A8A536C-D046-4D47-B259-BE24A7363D7F}"
                }));

            sheet.InsertRows(3, 2);

            Assert.Equal(7U, pageBreak.Id!.Value);
            Assert.Equal(1U, rowBreaks.Count!.Value);
            Assert.Equal(1U, rowBreaks.ManualBreakCount!.Value);
            Assert.Empty(document.ValidateOpenXml());
        }

        [Fact]
        public void Test_StructuralRows_PreflightsCustomSheetViewRowBreaks() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellAt(1, 1).SetValue("keep");
            var pageBreak = new Break { Id = (uint)A1.MaxRows, ManualPageBreak = true };
            sheet.WorksheetPart.Worksheet.Append(new CustomSheetViews(
                new CustomSheetView(
                    new RowBreaks(pageBreak) {
                        Count = 1U,
                        ManualBreakCount = 1U
                    }) {
                    Guid = "{366440C1-9E84-4299-A4ED-8D63C661BAC7}"
                }));

            Assert.Throws<InvalidOperationException>(() => sheet.InsertRows(1));

            Assert.Equal((uint)A1.MaxRows, pageBreak.Id!.Value);
            Assert.Equal("keep", sheet.CellAt(1, 1).GetValue<string>());
        }

        [Fact]
        public void Test_StructuralRows_RemapsOffice2010SortConditions() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellAt(1, 1).SetValue("Value");
            sheet.CellAt(2, 1).SetValue(1);
            sheet.CellAt(3, 1).SetValue(2);
            sheet.AutoFilterAdd("A1:A3");
            var worksheetCondition = new X14.SortCondition { Reference = "A2:A3" };
            sheet.WorksheetPart.Worksheet.GetFirstChild<AutoFilter>()!.Append(
                new SortState(worksheetCondition) { Reference = "A2:A3" });

            sheet.InsertRows(2);

            Assert.Equal("A3:A4", worksheetCondition.Reference!.Value);
        }

        [Fact]
        public void Test_StructuralRows_InvalidatesExtendedChartDimensionCaches() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellAt(2, 1).SetValue(1);
            sheet.CellAt(3, 1).SetValue(2);
            DrawingsPart drawingsPart = sheet.WorksheetPart.AddNewPart<DrawingsPart>();
            drawingsPart.WorksheetDrawing = new Xdr.WorksheetDrawing();
            ExtendedChartPart chartPart = drawingsPart.AddNewPart<ExtendedChartPart>();
            var formula = new Cx.Formula("A2:A3");
            var level = new Cx.NumericLevel(
                new Cx.NumericValue("1") { Idx = 0U },
                new Cx.NumericValue("2") { Idx = 1U }) {
                PtCount = 2U
            };
            var dimension = new Cx.NumericDimension(formula, level);
            chartPart.ChartSpace = new Cx.ChartSpace(dimension);

            sheet.InsertRows(2);

            Assert.Equal("A3:A4", formula.Text);
            Assert.Empty(dimension.Elements<Cx.NumericLevel>());
        }

        [Fact]
        public void Test_StructuralRows_RemapsDrawingAnchorsInsideAlternateContent() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = CreateDrawingSheet(document);
            Xdr.WorksheetDrawing drawing = sheet.WorksheetPart.DrawingsPart!.WorksheetDrawing!;
            Xdr.OneCellAnchor original = Assert.Single(drawing.Elements<Xdr.OneCellAnchor>());
            Xdr.OneCellAnchor fallbackAnchor = (Xdr.OneCellAnchor)original.CloneNode(true);
            original.Remove();
            var choice = new AlternateContentChoice(original) { Requires = "xdr" };
            choice.AddNamespaceDeclaration("xdr", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing");
            drawing.Append(new AlternateContent(
                choice,
                new AlternateContentFallback(fallbackAnchor)));

            sheet.InsertRows(5, 2);

            Assert.Equal("6", original.FromMarker!.RowId!.Text);
            Assert.Equal("6", fallbackAnchor.FromMarker!.RowId!.Text);
            Assert.Empty(document.ValidateOpenXml());
        }
    }
}
