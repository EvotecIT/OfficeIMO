using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_StructuralRows_RemapsSavedViewAndPaneTopLeftCells() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellAt(5, 1).SetValue(1);

            var normalPane = new Pane { TopLeftCell = "B5" };
            var normalView = new SheetView(normalPane) {
                WorkbookViewId = 0U,
                TopLeftCell = "A5"
            };
            var sheetViews = new SheetViews(normalView);
            SheetData sheetData = sheet.WorksheetPart.Worksheet.GetFirstChild<SheetData>()!;
            sheet.WorksheetPart.Worksheet.InsertBefore(sheetViews, sheetData);

            var customPane = new Pane { TopLeftCell = "D5" };
            var customView = new CustomSheetView(customPane) {
                Guid = "{F6130417-5307-4383-8797-A7314A4C5764}",
                TopLeftCell = "C5"
            };
            sheet.WorksheetPart.Worksheet.Append(new CustomSheetViews(customView));

            sheet.InsertRows(3);

            Assert.Equal("A6", normalView.TopLeftCell!.Value);
            Assert.Equal("B6", normalPane.TopLeftCell!.Value);
            Assert.Equal("C6", customView.TopLeftCell!.Value);
            Assert.Equal("D6", customPane.TopLeftCell!.Value);

            sheet.DeleteRows(6);

            Assert.Equal("A6", normalView.TopLeftCell!.Value);
            Assert.Equal("B6", normalPane.TopLeftCell!.Value);
            Assert.Equal("C6", customView.TopLeftCell!.Value);
            Assert.Equal("D6", customPane.TopLeftCell!.Value);
            Assert.Empty(document.ValidateOpenXml());
        }

        [Fact]
        public void Test_StructuralRows_PreflightsSavedViewTopLeftCells() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellAt(1, 1).SetValue("Keep");
            var view = new SheetView {
                WorkbookViewId = 0U,
                TopLeftCell = $"A{A1.MaxRows}"
            };
            var sheetViews = new SheetViews(view);
            SheetData sheetData = sheet.WorksheetPart.Worksheet.GetFirstChild<SheetData>()!;
            sheet.WorksheetPart.Worksheet.InsertBefore(sheetViews, sheetData);

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(
                () => sheet.InsertRows(1));

            Assert.Contains("row limit", exception.Message);
            Assert.Equal($"A{A1.MaxRows}", view.TopLeftCell!.Value);
            Assert.Equal("Keep", sheet.CellAt(1, 1).GetValue<string>());
        }

        [Fact]
        public void Test_StructuralRows_RewritesDrawingShapeTextLinksWorkbookWide() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet data = document.AddWorksheet("Data");
            ExcelSheet dashboard = document.AddWorksheet("Dashboard");
            data.CellAt(5, 1).SetValue("Moved");
            Xdr.Shape localShape = AddCellLinkedDrawingShape(data, 1U, "A5");
            Xdr.Shape remoteShape = AddCellLinkedDrawingShape(dashboard, 2U, "Data!A5");

            data.InsertRows(5);

            Assert.Equal("A6", localShape.TextLink!.Value);
            Assert.Equal("Data!A6", remoteShape.TextLink!.Value);
            Assert.Empty(document.ValidateOpenXml());
        }

        [Fact]
        public void Test_StructuralRows_PreflightsDrawingShapeTextLinksWorkbookWide() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet data = document.AddWorksheet("Data");
            ExcelSheet dashboard = document.AddWorksheet("Dashboard");
            data.CellAt(1, 1).SetValue("Keep");
            Xdr.Shape shape = AddCellLinkedDrawingShape(
                dashboard,
                1U,
                $"Data!A{A1.MaxRows}");

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(
                () => data.InsertRows(1));

            Assert.Contains("row limit", exception.Message);
            Assert.Equal($"Data!A{A1.MaxRows}", shape.TextLink!.Value);
            Assert.Equal("Keep", data.CellAt(1, 1).GetValue<string>());
        }

        [Fact]
        public void Test_StructuralRows_MutationPlanCountsOnlyChangedDrawingAnchors() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            AddCellLinkedDrawingShape(sheet, 1U, string.Empty);

            ExcelRowMutationPlan unaffected = sheet.PlanInsertRows(10);
            Assert.DoesNotContain(
                unaffected.Impacts,
                impact => impact.Category == "drawings");

            ExcelRowMutationPlan affected = sheet.PlanInsertRows(1);
            ExcelMutationImpact drawings = Assert.Single(
                affected.Impacts,
                impact => impact.Category == "drawings");
            Assert.Equal(1, drawings.ItemCount);
        }

        private static Xdr.Shape AddCellLinkedDrawingShape(
            ExcelSheet sheet,
            uint shapeId,
            string textLink) {
            DrawingsPart drawingsPart = sheet.WorksheetPart.DrawingsPart
                ?? sheet.WorksheetPart.AddNewPart<DrawingsPart>();
            drawingsPart.WorksheetDrawing ??= new Xdr.WorksheetDrawing();
            if (!sheet.WorksheetPart.Worksheet
                .Elements<DocumentFormat.OpenXml.Spreadsheet.Drawing>()
                .Any()) {
                sheet.WorksheetPart.Worksheet.Append(
                    new DocumentFormat.OpenXml.Spreadsheet.Drawing {
                        Id = sheet.WorksheetPart.GetIdOfPart(drawingsPart)
                    });
            }

            var shape = new Xdr.Shape(
                new Xdr.NonVisualShapeProperties(
                    new Xdr.NonVisualDrawingProperties {
                        Id = shapeId,
                        Name = $"Linked shape {shapeId}"
                    },
                    new Xdr.NonVisualShapeDrawingProperties()),
                new Xdr.ShapeProperties(
                    new A.PresetGeometry(new A.AdjustValueList()) {
                        Preset = A.ShapeTypeValues.Rectangle
                    })) {
                TextLink = textLink
            };
            drawingsPart.WorksheetDrawing.Append(
                new Xdr.OneCellAnchor(
                    new Xdr.FromMarker(
                        new Xdr.ColumnId("0"),
                        new Xdr.ColumnOffset("0"),
                        new Xdr.RowId("0"),
                        new Xdr.RowOffset("0")),
                    new Xdr.Extent { Cx = 9525L, Cy = 9525L },
                    shape,
                    new Xdr.ClientData()));
            drawingsPart.WorksheetDrawing.Save();
            sheet.WorksheetPart.Worksheet.Save();
            return shape;
        }
    }
}
