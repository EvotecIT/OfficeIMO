using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_StructuralRows_RemovesEmptyDataValidationContainer() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellAt(2, 1).SetValue(1);
            sheet.WorksheetPart.Worksheet.Append(
                new DataValidations(
                    new DataValidation {
                        SequenceOfReferences = new ListValue<StringValue> { InnerText = "A2" }
                    }) {
                    Count = 1U
                });

            sheet.DeleteRows(2);

            Assert.Null(sheet.WorksheetPart.Worksheet.GetFirstChild<DataValidations>());
            Assert.Empty(document.ValidateOpenXml());
        }

        [Fact]
        public void Test_StructuralRows_RemovesEmptyHyperlinkContainer() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellAt(2, 1).SetValue("Link");
            sheet.SetHyperlink(2, 1, "https://example.com", style: false);

            sheet.DeleteRows(2);

            Assert.Null(sheet.WorksheetPart.Worksheet.GetFirstChild<Hyperlinks>());
            Assert.Empty(document.ValidateOpenXml());
        }

        [Fact]
        public void Test_StructuralRows_PreflightsCellWatchReferences() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            var watch = new CellWatch { CellReference = $"A{A1.MaxRows}" };
            sheet.WorksheetPart.Worksheet.Append(new CellWatches(watch));

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(
                () => sheet.InsertRows(1));

            Assert.Contains("row limit", exception.Message);
            Assert.Equal($"A{A1.MaxRows}", watch.CellReference!.Value);
        }

        [Fact]
        public void Test_StructuralRows_ClampsDeletedOneCellDrawingAnchor() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellAt(1, 1).SetValue("Category");
            sheet.CellAt(1, 2).SetValue("Value");
            sheet.CellAt(2, 1).SetValue("One");
            sheet.CellAt(2, 2).SetValue(1);
            sheet.AddChartFromRange("A1:B2", row: 5, column: 4);
            Xdr.OneCellAnchor anchor = Assert.Single(
                sheet.WorksheetPart.DrawingsPart!.WorksheetDrawing!.Elements<Xdr.OneCellAnchor>());

            sheet.DeleteRows(5);

            Xdr.OneCellAnchor shifted = Assert.Single(
                sheet.WorksheetPart.DrawingsPart!.WorksheetDrawing!.Elements<Xdr.OneCellAnchor>());
            Assert.Same(anchor, shifted);
            Assert.Equal("4", shifted.FromMarker!.RowId!.Text);
            Assert.Empty(document.ValidateOpenXml());
        }

        [Fact]
        public void Test_StructuralRows_InvalidatesNamedChartSourceCaches() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellAt(1, 1).SetValue("Category");
            sheet.CellAt(1, 2).SetValue("Value");
            sheet.CellAt(2, 1).SetValue("One");
            sheet.CellAt(2, 2).SetValue(1);
            sheet.CellAt(3, 1).SetValue("Two");
            sheet.CellAt(3, 2).SetValue(2);
            document.SetNamedRange("MySeries", "'Data'!$B$2:$B$3", save: false);
            sheet.AddChartFromRange("A1:B3", row: 5, column: 4);

            ChartPart chartPart = Assert.Single(sheet.WorksheetPart.DrawingsPart!.ChartParts);
            C.Formula namedFormula = chartPart.ChartSpace.Descendants<C.Formula>()
                .First(formula => formula.Parent!.ChildElements.Any(element =>
                    element.LocalName.EndsWith("Cache", System.StringComparison.OrdinalIgnoreCase)));
            namedFormula.Text = "MySeries";
            Assert.Contains(
                namedFormula.Parent!.ChildElements,
                element => element.LocalName.EndsWith("Cache", System.StringComparison.OrdinalIgnoreCase));

            sheet.InsertRows(2);

            WorkbookPart workbookPart = sheet.WorksheetPart.GetParentParts().OfType<WorkbookPart>().Single();
            DefinedName name = Assert.Single(workbookPart.Workbook.DefinedNames!.Elements<DefinedName>());
            Assert.Equal("'Data'!$B$3:$B$4", name.Text);
            Assert.Equal("MySeries", namedFormula.Text);
            Assert.DoesNotContain(
                namedFormula.Parent!.ChildElements,
                element => element.LocalName.EndsWith("Cache", System.StringComparison.OrdinalIgnoreCase));
        }
    }
}
