using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Excel.Pdf {
    public static partial class ExcelPdfConverterExtensions {
        // Pagination retains increasing source columns. Only a page's physical grid is
        // reversed, so widths, references and cell decorations travel with their cells.
        private static SheetExportData ReverseWorksheetPageColumns(SheetExportData source) {
            int rows = source.Values.GetLength(0);
            int columns = source.Values.GetLength(1);
            object?[,] values = ReverseColumns(source.Values)!;
            ExcelCellStyleInfo?[,]? styles = ReverseColumns(source.Styles);
            ExcelHyperlinkInfo?[,]? hyperlinks = ReverseColumns(source.Hyperlinks);
            string?[,]? references = ReverseColumns(source.CellReferences);
            MergeLayoutData? merged = source.MergedCells == null ? null : new MergeLayoutData(rows, columns);
            Dictionary<(int Row, int Column), string>? fills = source.ConditionalFills == null ? null : ReverseCellVisuals(source.ConditionalFills.FillColors, columns);
            Dictionary<(int Row, int Column), ConditionalDataBarCell>? bars = source.ConditionalFills == null ? null : ReverseCellVisuals(source.ConditionalFills.DataBars, columns);
            Dictionary<(int Row, int Column), ConditionalIconCell>? icons = source.ConditionalFills == null ? null : ReverseCellVisuals(source.ConditionalFills.Icons, columns);
            if (merged != null) {
                for (int row = 0; row < rows; row++) {
                    for (int column = 0; column < columns; column++) {
                        MergeRegion? region = source.MergedCells!.GetRegion(row, column);
                        if (region == null || region.Row != row || region.Column != column) continue;
                        MergeSpan span = region.Span;
                        int physicalColumn = columns - column - span.ColumnSpan;
                        PdfCore.PdfTableCellViewport? viewport = span.Viewport;
                        if (viewport != null) viewport = new PdfCore.PdfTableCellViewport(
                            viewport.ContentWidth, viewport.ContentHeight, viewport.Width, viewport.Height,
                            viewport.ContentWidth - viewport.Width - viewport.OffsetX, viewport.OffsetY);
                        merged.SetSpan(row, physicalColumn, span.RowSpan, span.ColumnSpan, viewport, span.ImageAnchorReference);
                        values[row, physicalColumn] = source.Values[row, column];
                        if (styles != null) styles[row, physicalColumn] = source.Styles![row, column];
                        if (hyperlinks != null) hyperlinks[row, physicalColumn] = source.Hyperlinks![row, column];
                        if (references != null) references[row, physicalColumn] = source.CellReferences![row, column];
                        if (source.ConditionalFills != null) {
                            CopyMergeAnchorVisual(source.ConditionalFills.FillColors, fills!, region, row, physicalColumn);
                            CopyMergeAnchorVisual(source.ConditionalFills.DataBars, bars!, region, row, physicalColumn);
                            CopyMergeAnchorVisual(source.ConditionalFills.Icons, icons!, region, row, physicalColumn);
                        }
                    }
                }
            }
            ColumnLayoutData? widths = source.ColumnWidths == null ? null : new ColumnLayoutData(
                source.ColumnWidths.WidthWeights.AsEnumerable().Reverse().ToList(), source.ColumnWidths.ApproximateWidthPoints);
            ConditionalFillData? conditional = source.ConditionalFills == null ? null : new ConditionalFillData(fills!, bars!, icons!);
            return new SheetExportData(values, styles, hyperlinks, references, merged, widths, source.RowHeights,
                source.HeaderRowCount, source.FirstBodyRowNumber,
                source.StructuredTables.Select(table => table.WithRightToLeftLayout()).ToArray(), conditional);
        }

        private static T[,]? ReverseColumns<T>(T[,]? source) {
            if (source == null) return null;
            int rows = source.GetLength(0);
            int columns = source.GetLength(1);
            var reversed = new T[rows, columns];
            for (int row = 0; row < rows; row++)
                for (int column = 0; column < columns; column++) reversed[row, columns - column - 1] = source[row, column];
            return reversed;
        }

        private static Dictionary<(int Row, int Column), T> ReverseCellVisuals<T>(
            IReadOnlyDictionary<(int Row, int Column), T> source, int columns) =>
            source.ToDictionary(entry => (entry.Key.Row, columns - entry.Key.Column - 1), entry => entry.Value);

        // Mirror placement intervals only. Glyphs, charts and image pixels retain their orientation.
        private static double GetWorksheetSceneObjectX(WorksheetPdfExportPlan plan, double tableWidth, double logicalX, double width) =>
            plan.RightToLeft ? tableWidth - logicalX - width : logicalX;
    }
}
