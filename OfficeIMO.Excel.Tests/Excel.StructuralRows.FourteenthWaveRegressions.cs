using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_StructuralRows_RejectsCompleteDeletionOfPivotConsolidationSource() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = CreatePivotSheet(document);
            sheet.CellAt(5, 1).SetValue(1);
            sheet.CellAt(6, 1).SetValue(2);
            PivotTableCacheDefinitionPart cachePart = Assert.Single(
                sheet.WorksheetPart.PivotTableParts).PivotTableCacheDefinitionPart!;
            var rangeSet = new RangeSet { Sheet = "Data", Reference = "A5:A6" };
            cachePart.PivotCacheDefinition!.CacheSource = new CacheSource(
                new Consolidation(
                    new RangeSets(rangeSet) { Count = 1U })) {
                Type = SourceValues.Consolidation
            };

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(
                () => sheet.DeleteRows(5, 2));

            Assert.Contains("complete consolidation source range", exception.Message);
            Assert.Equal("A5:A6", rangeSet.Reference!.Value);
            Assert.Equal(1, sheet.CellAt(5, 1).GetValue<int>());
            Assert.Equal(2, sheet.CellAt(6, 1).GetValue<int>());
        }

        [Fact]
        public void Test_StructuralRows_ShiftsOneCellEndMarkerByActualClampDelta() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = CreateDrawingSheet(document);
            Xdr.TwoCellAnchor anchor = ReplaceWithTwoCellAnchor(
                sheet,
                fromRow: 4,
                toRow: 7,
                toRowOffset: "0",
                Xdr.EditAsValues.OneCell);

            sheet.DeleteRows(4, 2);

            Assert.Equal("3", anchor.FromMarker!.RowId!.Text);
            Assert.Equal("6", anchor.ToMarker!.RowId!.Text);
        }

        [Fact]
        public void Test_StructuralRows_RemapsTwoCellEndpointInsideAffectedRow() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = CreateDrawingSheet(document);
            Xdr.TwoCellAnchor anchor = ReplaceWithTwoCellAnchor(
                sheet,
                fromRow: 0,
                toRow: 4,
                toRowOffset: "1",
                Xdr.EditAsValues.TwoCell);

            sheet.InsertRows(5);

            Assert.Equal("0", anchor.FromMarker!.RowId!.Text);
            Assert.Equal("5", anchor.ToMarker!.RowId!.Text);
            Assert.Equal("1", anchor.ToMarker.RowOffset!.Text);
        }

        private static ExcelSheet CreateDrawingSheet(ExcelDocument document) {
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellAt(1, 1).SetValue("Category");
            sheet.CellAt(1, 2).SetValue("Value");
            sheet.CellAt(2, 1).SetValue("One");
            sheet.CellAt(2, 2).SetValue(1);
            sheet.AddChartFromRange("A1:B2", row: 5, column: 4);
            return sheet;
        }

        private static Xdr.TwoCellAnchor ReplaceWithTwoCellAnchor(
            ExcelSheet sheet,
            int fromRow,
            int toRow,
            string toRowOffset,
            Xdr.EditAsValues editAs) {
            Xdr.WorksheetDrawing drawing = sheet.WorksheetPart.DrawingsPart!.WorksheetDrawing!;
            Xdr.OneCellAnchor oneCellAnchor = Assert.Single(drawing.Elements<Xdr.OneCellAnchor>());
            var anchor = new Xdr.TwoCellAnchor(
                new Xdr.FromMarker(
                    new Xdr.ColumnId("3"),
                    new Xdr.ColumnOffset("0"),
                    new Xdr.RowId(fromRow.ToString()),
                    new Xdr.RowOffset("0")),
                new Xdr.ToMarker(
                    new Xdr.ColumnId("8"),
                    new Xdr.ColumnOffset("0"),
                    new Xdr.RowId(toRow.ToString()),
                    new Xdr.RowOffset(toRowOffset)),
                oneCellAnchor.GetFirstChild<Xdr.GraphicFrame>()!.CloneNode(true),
                new Xdr.ClientData()) {
                EditAs = editAs
            };
            oneCellAnchor.Remove();
            drawing.Append(anchor);
            return anchor;
        }
    }
}
