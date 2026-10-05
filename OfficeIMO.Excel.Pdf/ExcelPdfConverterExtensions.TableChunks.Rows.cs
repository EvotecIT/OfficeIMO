namespace OfficeIMO.Excel.Pdf {
    public static partial class ExcelPdfConverterExtensions {
        // Map authored row numbers after hidden-row filtering and any title-row prepend.
        private static IReadOnlyList<int> GetPrintTitleRowIndexes(ExcelSheet? sheet, string?[,]? references, ExcelToPdfOptions options) {
            ExcelPrintTitles? titles = sheet?.GetPrintTitles();
            if (!options.UseWorksheetPrintTitleRows || titles?.HasRows != true || references == null) return Array.Empty<int>();
            var indexes = new List<int>();
            for (int row = 0; row < references.GetLength(0); row++) {
                int source = GetOriginalRowNumber(references, row);
                if (source >= titles.FirstRow!.Value && source <= titles.LastRow!.Value) indexes.Add(row);
            }
            return indexes;
        }

        // A title block can start inside the body. Prepend only rows already reached;
        // future title rows stay at their natural position, including a break within the block.
        private static IReadOnlyList<int> CreateChunkRowIndexes(WorksheetPdfExportPlan plan, TableAxisChunk rowChunk) =>
            plan.ExportData.RepeatingRowIndexes.Where(row => row < rowChunk.Start)
                .Concat(Enumerable.Range(rowChunk.Start, rowChunk.Count)).ToArray();

        private static int GetChunkHeaderRowCount(WorksheetPdfExportPlan plan, IReadOnlyList<int> rows) {
            // PDF tables can repeat a leading header prefix. Natural middle rows become
            // that prefix only on later chunks, after earlier body rows have been traversed.
            var repeating = new HashSet<int>(plan.ExportData.RepeatingRowIndexes);
            int count = 0;
            while (count < rows.Count && repeating.Contains(rows[count])) count++;
            return count;
        }
    }
}
