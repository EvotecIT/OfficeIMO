using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_StructuralRows_RejectsFormulaBackedPivotSourceHeaderDeletion() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = CreatePivotSheet(document);
            WorksheetSource source = Assert.Single(
                sheet.WorksheetPart.PivotTableParts).PivotTableCacheDefinitionPart!
                .PivotCacheDefinition!.CacheSource!.WorksheetSource!;
            source.Reference = null;
            source.Sheet = null;
            source.Name = "PivotSource";
            var definedName = new DefinedName("OFFSET(Data!$A$1,0,0,3,2)") {
                Name = "PivotSource"
            };
            document.WorkbookRoot.DefinedNames = new DefinedNames(definedName);

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(
                () => sheet.DeleteRows(1));

            Assert.Contains("header row", exception.Message);
            Assert.Equal("OFFSET(Data!$A$1,0,0,3,2)", definedName.Text);
            Assert.Equal("Region", sheet.CellAt(1, 1).GetValue<string>());
        }

        [Fact]
        public void Test_StructuralRows_RewritesSharedChartPartOnlyOnce() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet data = document.AddWorksheet("Data");
            data.CellAt(5, 1).SetValue(1);
            ExcelSheet firstOwner = CreateChartOwner(document, "First owner");
            ExcelSheet secondOwner = CreateChartOwner(document, "Second owner");

            ChartPart sharedChartPart = Assert.Single(firstOwner.WorksheetPart.DrawingsPart!.ChartParts);
            C.Formula formula = sharedChartPart.ChartSpace.Descendants<C.Formula>().First();
            formula.Text = "Data!A5";
            DrawingsPart secondDrawings = secondOwner.WorksheetPart.DrawingsPart!;
            ChartPart replacedPart = Assert.Single(secondDrawings.ChartParts);
            string relationshipId = secondDrawings.GetIdOfPart(replacedPart);
            secondDrawings.DeletePart(replacedPart);
            secondDrawings.AddPart(sharedChartPart, relationshipId);

            data.InsertRows(5);

            Assert.Equal("Data!A6", formula.Text);
            Assert.Same(sharedChartPart, Assert.Single(secondDrawings.ChartParts));
        }

        [Fact]
        public void Test_StructuralRows_ClampsCompletelyDeletedTwoCellAnchor() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = CreateDrawingSheet(document);
            Xdr.TwoCellAnchor anchor = ReplaceWithTwoCellAnchor(
                sheet,
                fromRow: 4,
                toRow: 10,
                toRowOffset: "0",
                Xdr.EditAsValues.TwoCell);

            sheet.DeleteRows(5, 6);

            Xdr.TwoCellAnchor retained = Assert.Single(
                sheet.WorksheetPart.DrawingsPart!.WorksheetDrawing!.Elements<Xdr.TwoCellAnchor>());
            Assert.Same(anchor, retained);
            Assert.Equal("4", retained.FromMarker!.RowId!.Text);
            Assert.Equal("4", retained.ToMarker!.RowId!.Text);
            Assert.Empty(document.ValidateOpenXml());
        }

        private static ExcelSheet CreateChartOwner(ExcelDocument document, string name) {
            ExcelSheet sheet = document.AddWorksheet(name);
            sheet.CellAt(1, 1).SetValue("Category");
            sheet.CellAt(1, 2).SetValue("Value");
            sheet.CellAt(2, 1).SetValue("One");
            sheet.CellAt(2, 2).SetValue(1);
            sheet.AddChartFromRange("A1:B2", row: 5, column: 4);
            return sheet;
        }
    }
}
