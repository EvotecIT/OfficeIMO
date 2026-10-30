using OfficeIMO.Drawing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Excel.Pdf {
    public static partial class ExcelPdfConverterExtensions {
        private static void ApplyWorksheetConditionalIconStyle(
            PdfCore.PdfTableStyle style, SheetExportData data,
            IReadOnlyList<double> columnWidths, IReadOnlyList<double> rowHeights, double scale) {
            if (data.ConditionalFills == null || data.ConditionalFills.Icons.Count == 0) return;
            var paddings = style.CellPaddings == null
                ? new Dictionary<(int Row, int Column), PdfCore.PdfCellPadding>()
                : new Dictionary<(int Row, int Column), PdfCore.PdfCellPadding>(style.CellPaddings);
            foreach (var entry in data.ConditionalFills.Icons) {
                if (entry.Value.Kind == OfficeConditionalIconKind.None) continue;
                var key = entry.Key;
                if (!TryGetWorksheetConditionalIconBounds(data, columnWidths, rowHeights, key.Row, key.Column, out var cell)) continue;
                var bounds = OfficeConditionalIconRenderer.GetCellBounds(cell.Width / 0.75D, cell.Height / 0.75D, 0.75D * scale);
                PdfCore.PdfCellPadding padding = paddings.TryGetValue(key, out var existing) ? existing.Clone() : new PdfCore.PdfCellPadding();
                padding.Left = global::OfficeIMO.OfficeLength.Points(Math.Min(cell.Width * scale, bounds.X + bounds.Size + 3D * scale + (padding.Left?.ToPoints() ?? style.CellPaddingLeft?.ToPoints() ?? style.CellPaddingX.ToPoints())));
                paddings[key] = padding;
            }
            style.CellPaddings = paddings;
        }

        private static void AddWorksheetSceneConditionalIcons(
            PdfCore.PdfPageCanvas canvas, SheetExportData data,
            IReadOnlyList<double> columnWidths, IReadOnlyList<double> rowHeights,
            double sceneX, double sceneY, double scale) {
            if (data.ConditionalFills == null) return;
            var drawings = new Dictionary<OfficeConditionalIconKind, OfficeDrawing>();
            foreach (var entry in data.ConditionalFills.Icons) {
                if (entry.Value.Kind == OfficeConditionalIconKind.None) continue;
                if (!TryGetWorksheetConditionalIconBounds(data, columnWidths, rowHeights, entry.Key.Row, entry.Key.Column, out var cell)) continue;
                var bounds = OfficeConditionalIconRenderer.GetCellBounds(cell.Width / 0.75D, cell.Height / 0.75D, 0.75D * scale);
                if (bounds.Size <= 0D) continue;
                if (!drawings.TryGetValue(entry.Value.Kind, out OfficeDrawing? drawing)) {
                    drawing = OfficeConditionalIconRenderer.CreateDrawing(entry.Value.Kind);
                    drawings[entry.Value.Kind] = drawing;
                }
                double x = sceneX + cell.X * scale;
                double y = sceneY + cell.Y * scale;
                canvas.Clip(x, y, cell.Width * scale, cell.Height * scale, clipped =>
                    clipped.Artifact(artifact => artifact.Drawing(drawing, x + bounds.X, y + bounds.Y, bounds.Size, bounds.Size)));
            }
        }

        private static bool TryGetWorksheetConditionalIconBounds(
            SheetExportData data, IReadOnlyList<double> widths, IReadOnlyList<double> heights,
            int row, int column, out (double X, double Y, double Width, double Height) bounds) {
            bounds = default;
            if (row < 0 || row >= heights.Count || column < 0 || column >= widths.Count || data.MergedCells?.IsContinuation(row, column) == true) return false;
            MergeSpan? span = data.MergedCells?.GetSpan(row, column);
            double width = widths.Skip(column).Take(span?.ColumnSpan ?? 1).Sum();
            double height = heights.Skip(row).Take(span?.RowSpan ?? 1).Sum();
            if (width <= 0D || height <= 0D) return false;
            bounds = (widths.Take(column).Sum(), heights.Take(row).Sum(), width, height);
            return true;
        }

        // FlowTable intentionally retains its compact generic indicators; worksheet fidelity
        // uses the shared vector artwork above without changing the general PDF icon API.
        private static PdfCore.PdfCellIconKind GetLegacyConditionalIconShape(OfficeConditionalIconKind kind) => kind switch {
            OfficeConditionalIconKind.GreenUpArrow => PdfCore.PdfCellIconKind.TriangleUp,
            OfficeConditionalIconKind.RedDownArrow => PdfCore.PdfCellIconKind.TriangleDown,
            OfficeConditionalIconKind.YellowUpArrow => PdfCore.PdfCellIconKind.TriangleUp,
            OfficeConditionalIconKind.YellowDownArrow => PdfCore.PdfCellIconKind.TriangleDown,
            OfficeConditionalIconKind.YellowSideArrow => PdfCore.PdfCellIconKind.TriangleRight,
            OfficeConditionalIconKind.RedCross => PdfCore.PdfCellIconKind.Diamond,
            OfficeConditionalIconKind.YellowExclamation => PdfCore.PdfCellIconKind.TriangleUp,
            _ => PdfCore.PdfCellIconKind.Circle
        };

        private static global::OfficeIMO.OfficeColor GetLegacyConditionalIconColor(OfficeConditionalIconKind kind) => kind switch {
            OfficeConditionalIconKind.GreenCheck or OfficeConditionalIconKind.GreenCircle or OfficeConditionalIconKind.GreenUpArrow or OfficeConditionalIconKind.GreenFlag or OfficeConditionalIconKind.RatingFive or OfficeConditionalIconKind.QuarterFull => global::OfficeIMO.OfficeColor.FromRgb(99, 155, 71),
            OfficeConditionalIconKind.RedCross or OfficeConditionalIconKind.RedCircle or OfficeConditionalIconKind.RedDownArrow or OfficeConditionalIconKind.RedFlag or OfficeConditionalIconKind.RatingOne or OfficeConditionalIconKind.QuarterEmpty => global::OfficeIMO.OfficeColor.FromRgb(192, 80, 77),
            _ => global::OfficeIMO.OfficeColor.FromRgb(255, 192, 0)
        };
    }
}
