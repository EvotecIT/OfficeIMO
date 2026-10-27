using OfficeIMO.Drawing;
using OfficeIMO.Excel.Utilities;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Excel.Pdf {
    public static partial class ExcelPdfConverterExtensions {
        private static IReadOnlyList<ExcelWorksheetDrawingObjectInfo> ReadWorksheetDrawingObjects(
            ExcelSheet? sheet, ExcelPdfExportOptions options, string sheetName,
            ISet<string>? exportedCells, bool filterToExportedCells) {
            if (sheet == null || !options.UseWorksheetImages) return Array.Empty<ExcelWorksheetDrawingObjectInfo>();
            var result = new List<ExcelWorksheetDrawingObjectInfo>();
            foreach (var drawing in ExcelWorksheetDrawingObjectResolver.FindDrawingObjects(sheet.WorksheetPart)) {
                if (filterToExportedCells && (drawing.CellReference == null || exportedCells?.Contains(drawing.CellReference) != true)) continue;
                if (!drawing.IsRenderable || options.WorksheetLayout != ExcelPdfWorksheetLayoutMode.WorksheetCanvas) {
                    AddWarning(options, sheetName, "worksheet-drawing-object",
                        $"Worksheet drawing '{drawing.Name}' was not exported: {drawing.UnsupportedReason ?? "drawing shapes require worksheet-canvas layout"}.");
                    continue;
                }
                result.Add(drawing);
            }
            return result;
        }

        private static (double Width, double Height) GetWorksheetDrawingSize(
            WorksheetPdfExportPlan plan, ExcelWorksheetDrawingObjectInfo drawing) {
            double width = PixelsToPoints(drawing.WidthPixels);
            double height = PixelsToPoints(drawing.HeightPixels);
            if (drawing.ToColumn is int lastColumn && lastColumn >= drawing.Column) {
                width = PixelsToPoints(drawing.ToOffsetXPixels - drawing.OffsetXPixels);
                for (int column = drawing.Column; column < lastColumn; column++) width += GetWorksheetColumnWidthPoints(plan.Geometry, column);
            }
            if (drawing.ToRow is int lastRow && lastRow >= drawing.Row) {
                height = PixelsToPoints(drawing.ToOffsetYPixels - drawing.OffsetYPixels);
                for (int row = drawing.Row; row < lastRow; row++) height += GetWorksheetRowHeightPoints(plan.Geometry, row);
            }
            return (Math.Max(0.75D, width), Math.Max(0.75D, height));
        }

        private static void AddWorksheetSceneDrawingObjects(
            PdfCore.PdfPageCanvas canvas, WorksheetPdfExportPlan plan, TableChunk chunk,
            IReadOnlyList<double> columnWidths, IReadOnlyList<double> rowHeights,
            double sceneX, double sceneY, double scale) {
            foreach (var drawing in plan.DrawingObjects.OrderBy(drawing => drawing.Order)) {
                if (!TryGetWorksheetObjectOffset(plan, chunk, columnWidths, rowHeights, drawing.Row, drawing.Column, out double x, out double y)) continue;
                (double width, double height) = GetWorksheetDrawingSize(plan, drawing);
                ExcelVisualDrawingObject visual = drawing.ToVisual(0D, 0D, width / 0.75D, height / 0.75D, plan.SheetName);
                // DrawingML's shared visual model uses logical 96-DPI pixels. PDF uses points.
                var scene = ExcelRangeImageRenderer.CreateOfficeDrawing(visual, 0.75D * scale);
                canvas.Drawing(scene.Drawing,
                    sceneX + (x + PixelsToPoints(drawing.OffsetXPixels)) * scale - scene.OffsetX,
                    sceneY + (y + PixelsToPoints(drawing.OffsetYPixels)) * scale - scene.OffsetY,
                    scene.Drawing.Width, scene.Drawing.Height,
                    new PdfCore.PdfDrawingStyle { AlternativeText = string.IsNullOrWhiteSpace(drawing.Text) ? drawing.Name : drawing.Text });
            }
        }
    }
}
