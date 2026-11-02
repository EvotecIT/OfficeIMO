namespace OfficeIMO.Excel.Pdf {
    public static partial class ExcelPdfConverterExtensions {
        private sealed class WorksheetScenePage {
            internal WorksheetScenePage(TableChunk chunk, WorksheetSceneTile tile) { Chunk = chunk; Tile = tile; }
            internal TableChunk Chunk { get; }
            internal WorksheetSceneTile Tile { get; }
        }

        /// <summary>Orders the complete page sequence after cell and media subdivision.</summary>
        private static IReadOnlyList<WorksheetScenePage> CreateWorksheetScenePages(WorksheetPdfExportPlan plan,
            IReadOnlyList<TableChunk> chunks, ExcelPdfExportOptions options,
            double availableWidth, double availableHeight, double worksheetScale) {
            var pages = new List<WorksheetScenePage>();
            bool hasMediaSubdivision = false;
            foreach (TableChunk chunk in chunks) {
                IReadOnlyList<WorksheetSceneTile> tiles = GetWorksheetSceneTiles(plan, chunk, options,
                    availableWidth, availableHeight, worksheetScale);
                hasMediaSubdivision |= tiles.Count > 1;
                foreach (WorksheetSceneTile tile in tiles) pages.Add(new WorksheetScenePage(chunk, tile));
            }
            // Existing cell ordering and single-chunk tile ordering are already
            // complete. Only their combination needs coordinates in a shared grid.
            if (!hasMediaSubdivision || chunks.Count <= 1) return pages;
            int BodyRow(TableChunk chunk) => chunk.HeaderRowCount < chunk.RowIndexes.Count
                ? chunk.RowIndexes[chunk.HeaderRowCount] : chunk.RowIndexes.LastOrDefault();
            IReadOnlyDictionary<int, double> rows = MeasureSceneAxisOffsets(chunks.Select(BodyRow),
                index => GetExportedRowHeightPoints(plan, index));
            IReadOnlyDictionary<int, double> columns = MeasureSceneAxisOffsets(chunks.Select(chunk => chunk.StartColumn),
                index => GetExportedColumnWidthPoints(plan, index));
            // Logical body bands determine order. Repeated titles affect paint
            // placement, but must never collapse two distinct source page bands.
            double X(WorksheetScenePage page) => columns[page.Chunk.StartColumn] + page.Tile.SceneX;
            double Y(WorksheetScenePage page) => rows[BodyRow(page.Chunk)] + page.Tile.SceneY;
            return plan.PageSetup?.PageOrder == ExcelPageOrder.OverThenDown
                ? pages.OrderBy(Y).ThenBy(X).ToArray()
                : pages.OrderBy(X).ThenBy(Y).ToArray();
        }

        // Measure only requested anchors, advancing through the exported grid once.
        // Repeated title rows/columns remain with painting, outside these anchors.
        private static IReadOnlyDictionary<int, double> MeasureSceneAxisOffsets(IEnumerable<int> anchors,
            Func<int, double> size) {
            var offsets = new Dictionary<int, double>();
            int previous = 0;
            double position = 0D;
            foreach (int anchor in anchors.Distinct().OrderBy(value => value)) {
                while (previous < anchor) position += size(previous++);
                offsets[anchor] = position;
            }
            return offsets;
        }
    }
}
