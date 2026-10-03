namespace OfficeIMO.Excel.Pdf {
    public static partial class ExcelPdfConverterExtensions {
        private static IReadOnlyList<TableChunk> CreateWorksheetSceneChunks(
            WorksheetPdfExportPlan plan, ExcelToPdfOptions options, int columns,
            double availableWidth, double availableHeight, double scale) {
            // A constrained fit axis controls its breaks; the unlimited axis retains manual breaks.
            IReadOnlyList<TableChunk> requested = CreateTableChunks(plan, options, columns,
                honorRowBreaks: !IsFitToHeight(plan.PageSetup), honorColumnBreaks: !IsFitToWidth(plan.PageSetup));
            var chunks = new List<TableChunk>();
            foreach (TableChunk requestedChunk in requested) {
                var columnSegments = SplitWorksheetColumns(plan, requestedChunk.StartColumn,
                    requestedChunk.ColumnCount, availableWidth / scale);
                var rowSegments = SplitWorksheetRows(plan, requestedChunk.RowIndexes,
                    requestedChunk.HeaderRowCount, availableHeight / scale);
                foreach ((int startColumn, int columnCount) in columnSegments) {
                    foreach (IReadOnlyList<int> rowIndexes in rowSegments) {
                        chunks.Add(new TableChunk(rowIndexes, GetChunkHeaderRowCount(plan, rowIndexes),
                            startColumn, columnCount, CreateChunkColumnIndexes(plan, startColumn, columnCount)));
                    }
                }
            }
            int BodyRow(TableChunk chunk) => chunk.RowIndexes.Skip(chunk.HeaderRowCount).FirstOrDefault();
            return plan.PageSetup?.PageOrder == ExcelPageOrder.OverThenDown
                ? chunks.OrderBy(BodyRow).ThenBy(chunk => chunk.StartColumn).ToArray()
                : chunks.OrderBy(chunk => chunk.StartColumn).ThenBy(BodyRow).ToArray();
        }

        private static double ResolveWorksheetPlanScale(WorksheetPdfExportPlan plan, int columns, double availableWidth, double availableHeight) {
            if (!ExcelPageSetupGeometry.HasFitToPageScale(plan.PageSetup)) return GetWorksheetAuthoredScale(plan.PageSetup);
            double scale = 1D;
            if (IsFitToWidth(plan.PageSetup)) {
                scale = FindWorksheetFitScale(scale, candidate => SplitWorksheetColumns(plan, 0, columns, availableWidth / candidate).Count,
                    plan.PageSetup!.FitToWidth!.Value);
            }
            if (IsFitToHeight(plan.PageSetup)) {
                var rows = Enumerable.Range(0, plan.ExportedRows).ToArray();
                scale = FindWorksheetFitScale(scale, candidate => SplitWorksheetRows(plan, rows,
                    plan.ExportData.HeaderRowCount, availableHeight / candidate).Count, plan.PageSetup!.FitToHeight!.Value);
            }
            return scale;
        }

        private static double FindWorksheetFitScale(double upper, Func<double, int> pageCount, uint targetPages) {
            if (pageCount(upper) <= targetPages) return upper;
            double lower = 0.05D;
            for (int iteration = 0; iteration < 20; iteration++) {
                double candidate = (lower + upper) / 2D;
                if (pageCount(candidate) <= targetPages) lower = candidate;
                else upper = candidate;
            }
            return lower;
        }

        private static IReadOnlyList<(int Start, int Count)> SplitWorksheetColumns(
            WorksheetPdfExportPlan plan, int startColumn, int columnCount, double maximumWidth) {
            if (columnCount <= 0) return new[] { (startColumn, columnCount) };
            var result = new List<(int Start, int Count)>();
            int segmentStart = startColumn;
            int segmentCount = 0;
            double segmentWidth = 0D;
            for (int column = startColumn; column < startColumn + columnCount; column++) {
                double width = GetExportedColumnWidthPoints(plan, column);
                double titleWidth = plan.ExportData.PrintTitleColumnIndexes.Where(title => title < segmentStart)
                    .Sum(title => GetExportedColumnWidthPoints(plan, title));
                if (segmentCount > 0 && segmentWidth + width + titleWidth > maximumWidth) {
                    result.Add((segmentStart, segmentCount));
                    segmentStart = column;
                    segmentCount = 0;
                    segmentWidth = 0D;
                }
                segmentWidth += width;
                segmentCount++;
            }
            if (segmentCount > 0) result.Add((segmentStart, segmentCount));
            return result;
        }

        private static IReadOnlyList<IReadOnlyList<int>> SplitWorksheetRows(
            WorksheetPdfExportPlan plan, IReadOnlyList<int> rowIndexes, int headerRowCount, double maximumHeight) {
            if (rowIndexes.Count == 0) return new[] { rowIndexes };
            int headerRows = Math.Min(headerRowCount, rowIndexes.Count);
            var initialHeaderIndexes = rowIndexes.Take(headerRows).ToArray();
            if (headerRows == rowIndexes.Count) return new[] { rowIndexes };
            var result = new List<IReadOnlyList<int>>();
            var currentBody = new List<int>();
            IReadOnlyList<int> headerIndexes = Array.Empty<int>();
            double bodyCapacity = maximumHeight;
            double currentHeight = 0D;
            for (int index = headerRows; index < rowIndexes.Count; index++) {
                int row = rowIndexes[index];
                double height = GetExportedRowHeightPoints(plan, row);
                if (currentBody.Count > 0 && currentHeight + height > bodyCapacity) {
                    result.Add(headerIndexes.Concat(currentBody).ToList());
                    currentBody.Clear();
                    currentHeight = 0D;
                }
                if (currentBody.Count == 0) {
                    // Automatic page breaks can newly reach part or all of a middle title
                    // block, so repeat its visited rows and reserve their height on this page.
                    headerIndexes = initialHeaderIndexes.Concat(plan.ExportData.RepeatingRowIndexes.Where(title => title < row))
                        .Distinct().OrderBy(title => title).ToArray();
                    bodyCapacity = Math.Max(1D, maximumHeight - headerIndexes.Sum(title => GetExportedRowHeightPoints(plan, title)));
                }
                currentBody.Add(row);
                currentHeight += height;
            }
            if (currentBody.Count > 0 || result.Count == 0) result.Add(headerIndexes.Concat(currentBody).ToList());
            return result;
        }
    }
}
