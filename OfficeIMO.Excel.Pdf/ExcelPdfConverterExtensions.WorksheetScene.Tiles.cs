using OfficeIMO.Drawing;

namespace OfficeIMO.Excel.Pdf {
    public static partial class ExcelPdfConverterExtensions {
        private readonly struct WorksheetSceneTile {
            internal WorksheetSceneTile(double x, double y, double width, double height, double sceneX = 0D, double sceneY = 0D) {
                OffsetX = x; OffsetY = y; Width = width; Height = height; SceneX = sceneX; SceneY = sceneY;
            }
            internal double OffsetX { get; }
            internal double OffsetY { get; }
            internal double Width { get; }
            internal double Height { get; }
            internal double SceneX { get; }
            internal double SceneY { get; }
        }

        /// <summary>Returns nonempty page windows in native worksheet print order.</summary>
        private static IReadOnlyList<WorksheetSceneTile> GetWorksheetSceneTiles(
            WorksheetPdfExportPlan plan,
            TableChunk chunk,
            ExcelPdfExportOptions options,
            double availableWidth,
            double availableHeight,
            double worksheetScale) {
            List<double> columnWidths = CreateWorksheetSceneColumnWidths(plan, chunk);
            List<double> rowHeights = CreateWorksheetSceneRowHeights(plan, chunk);
            WorksheetSceneBounds bounds = MeasureWorksheetObjects(plan, chunk, columnWidths, rowHeights);
            double tableWidth = Math.Max(1D, columnWidths.Sum());
            double tableHeight = Math.Max(1D, rowHeights.Sum());
            double width = Math.Max(tableWidth, bounds.Right) - bounds.Left;
            double height = Math.Max(tableHeight, bounds.Bottom) - bounds.Top;
            double scale = ResolveWorksheetSceneScale(plan.PageSetup, width, height, availableWidth, availableHeight, worksheetScale);
            var rects = new List<(double X, double Y, double Width, double Height)>();
            if (plan.HasTable && chunk.RowIndexes.Count > 0 && chunk.ColumnCount > 0) rects.Add((0D, 0D, tableWidth, tableHeight));
            foreach (WorksheetImageExportData image in plan.Images) {
                if (TryGetWorksheetObjectOffset(plan, chunk, columnWidths, rowHeights, image.RowIndex, image.ColumnIndex, out double x, out double y,
                        image.OffsetXPoints + image.WidthPoints, image.OffsetYPoints + image.HeightPoints)) {
                    rects.Add((x + image.OffsetXPoints, y + image.OffsetYPoints, image.WidthPoints, image.HeightPoints));
                }
            }

            foreach (WorksheetChartExportData chart in plan.Charts) {
                ExcelChartInfo snapshot = chart.Snapshot;
                if (TryGetWorksheetObjectOffset(plan, chunk, columnWidths, rowHeights, snapshot.RowIndex, snapshot.ColumnIndex, out double x, out double y,
                        (snapshot.OffsetX.ToPoints() + snapshot.Width.ToPoints()), (snapshot.OffsetY.ToPoints() + snapshot.Height.ToPoints()))) {
                    rects.Add((x + snapshot.OffsetX.ToPoints(), y + snapshot.OffsetY.ToPoints(),
                        snapshot.Width.ToPoints(), snapshot.Height.ToPoints()));
                }
            }

            foreach (var drawing in plan.DrawingObjects) {
                (double drawingWidth, double drawingHeight) = GetWorksheetDrawingSize(plan, drawing);
                if (TryGetWorksheetObjectOffset(plan, chunk, columnWidths, rowHeights, drawing.Row, drawing.Column, out double x, out double y,
                        PixelsToPoints(drawing.OffsetXPixels) + drawingWidth, PixelsToPoints(drawing.OffsetYPixels) + drawingHeight)) {
                    rects.Add((x + PixelsToPoints(drawing.OffsetXPixels), y + PixelsToPoints(drawing.OffsetYPixels), drawingWidth, drawingHeight));
                }
            }

            // Page windows are aligned to the chunk's first cell, so the cells (which the row
            // and column breaks already fit on one page) are never cut by a window edge.
            double tileWidth = availableWidth / scale;
            double tileHeight = availableHeight / scale;
            double right = Math.Max(tableWidth, bounds.Right);
            double bottom = Math.Max(tableHeight, bounds.Bottom);
            IReadOnlyList<double> columnBreaks = options.UseWorksheetPageBreaks && !IsFitToWidth(plan.PageSetup)
                ? GetUnmappedSceneBreakPositions(plan, chunk, columnWidths, rowHeights, columns: true)
                : Array.Empty<double>();
            IReadOnlyList<double> rowBreaks = options.UseWorksheetPageBreaks && !IsFitToHeight(plan.PageSetup)
                ? GetUnmappedSceneBreakPositions(plan, chunk, columnWidths, rowHeights, columns: false)
                : Array.Empty<double>();
            var horizontal = CreateSceneAxisWindows(bounds.Left, right, tileWidth, columnBreaks, scale);
            // A single column wider than the page prints once, cut at the page edge, as in a
            // spreadsheet application. Continuing it on another page would only show the cut
            // ends of words.
            if (horizontal.Count > 1 && chunk.ColumnCount == 1 && rects.Count == 1 && columnBreaks.Count == 0) {
                horizontal = new[] { horizontal[0] };
            }
            var vertical = CreateSceneAxisWindows(bounds.Top, bottom, tileHeight, rowBreaks, scale);
            var tiles = new List<WorksheetSceneTile>();
            void Add((double Start, double Length) x, (double Start, double Length) y) {
                bool hasContent = rects.Any(rect => rect.X < x.Start + x.Length && rect.X + rect.Width > x.Start &&
                    rect.Y < y.Start + y.Length && rect.Y + rect.Height > y.Start);
                if (hasContent) tiles.Add(new WorksheetSceneTile((x.Start - bounds.Left) * scale,
                    (y.Start - bounds.Top) * scale, x.Length * scale, y.Length * scale, x.Start, y.Start));
            }
            if (plan.PageSetup?.PageOrder == ExcelPageOrder.OverThenDown) {
                foreach (var y in vertical) foreach (var x in horizontal) Add(x, y);
            } else {
                foreach (var x in horizontal) foreach (var y in vertical) Add(x, y);
            }
            if (tiles.Count == 0) tiles.Add(new WorksheetSceneTile(0D, 0D, availableWidth, availableHeight));
            return tiles;
        }

        // Cell chunks already handle breaks within populated ranges. Media can occupy
        // leading/trailing worksheet bands that those cell-index chunks cannot describe.
        private static IReadOnlyList<double> GetUnmappedSceneBreakPositions(WorksheetPdfExportPlan plan,
            TableChunk chunk, IReadOnlyList<double> widths, IReadOnlyList<double> heights, bool columns) {
            int first = columns ? plan.Geometry.FirstColumn : plan.Geometry.FirstRow;
            int last = columns ? plan.Geometry.LastColumn : plan.Geometry.LastRow;
            var nativeBreaks = (columns ? plan.ManualColumnBreaks : plan.ManualRowBreaks)
                .Where(value => value > 0 && (!plan.HasTable || value < first || value >= last))
                .Where(value => !plan.ClipToExportRange || value >= first && value < last)
                .Distinct().OrderBy(value => value).ToArray();
            var result = new List<double>();
            string?[,]? references = plan.ExportData.CellReferences;
            int rows = references == null ? 0 : Math.Min(plan.ExportedRows, references.GetLength(0));
            int sourceRow = references != null && chunk.RowIndexes.Count > 0
                ? GetOriginalRowNumber(references, chunk.RowIndexes[Math.Min(chunk.HeaderRowCount, chunk.RowIndexes.Count - 1)])
                : plan.Geometry.FirstRow;
            int sourceColumn = references != null && chunk.ColumnCount > 0
                ? GetOriginalColumnNumber(references, chunk.StartColumn, rows) : plan.Geometry.FirstColumn;
            // Seed each leading/trailing band through the canonical object mapping.
            // Advance monotonically within that band rather than rescanning a distant
            // million-row anchor for every manual break.
            foreach (var band in nativeBreaks.GroupBy(value => !plan.HasTable || value < first ? 0 : 1)) {
                int previous = -1;
                double position = 0D;
                foreach (int nativeBreak in band) {
                    int next = checked(nativeBreak + 1);
                    if (previous < 0) {
                        if (!TryGetWorksheetObjectOffset(plan, chunk, widths, heights,
                            columns ? sourceRow : next, columns ? next : sourceColumn,
                            out double x, out double y, double.PositiveInfinity, double.PositiveInfinity)) continue;
                        position = columns ? x : y;
                    } else {
                        for (int index = previous; index < next; index++) position += columns
                            ? GetWorksheetColumnWidthPoints(plan.Geometry, index)
                            : GetWorksheetRowHeightPoints(plan.Geometry, index);
                    }
                    previous = next;
                    result.Add(position);
                }
            }
            return result;
        }

        private static IReadOnlyList<(double Start, double Length)> CreateSceneAxisWindows(
            double minimum, double maximum, double pageSize, IReadOnlyList<double> breaks, double scale) {
            double[] boundaries = breaks.Where(value => value > minimum + 1E-9 && value < maximum - 1E-9)
                .Distinct().OrderBy(value => value).ToArray();
            // A scene that fits has no artificial boundary at the chunk's cell
            // origin. Leading media can have negative coordinates relative to it.
            if (boundaries.Length == 0 && maximum - minimum <= pageSize + 1D / scale)
                return new[] { (minimum, pageSize) };
            double firstBoundary = boundaries.Length > 0 ? boundaries[0] : maximum;
            double start = firstBoundary - minimum <= pageSize + 1D / scale
                ? minimum : Math.Floor(minimum / pageSize + 1E-9) * pageSize;
            var windows = new List<(double Start, double Length)>();
            foreach (double boundary in boundaries) {
                while (start < boundary - 1E-9) {
                    double end = Math.Min(start + pageSize, boundary);
                    windows.Add((start, end - start));
                    start = end;
                }
            }
            // Keep the existing one-point tolerance for an automatic terminal sliver.
            while (start < maximum - 1D / scale) {
                windows.Add((start, pageSize));
                start += pageSize;
            }
            if (windows.Count == 0) windows.Add((start, pageSize));
            return windows;
        }
    }
}
