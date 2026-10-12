namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    /// <summary>Retains a merged text cell's cursor separately from the physical rows that carry it.</summary>
    private sealed class TableSpanFlow {
        public List<TableSpanCellFlow> Cells { get; } = new();
        public List<TableSpanCellFlow> ActiveCells { get; } = new();

        // A delayed opaque fill can overlap a border drawn for an earlier neighbor.
        public bool HasDeferredFills(PdfTableStyle style) =>
            style.CellFills?.Keys.Any(key => Contains(key.Item1, key.Item2)) == true;
        private readonly Dictionary<(int Row, int Column), TableSpanCellFlow> anchors = new();
        private readonly HashSet<int> coveredRows = new();
        private int nextCell;

        public TableSpanFlow(TableBlock table, PdfTableStyle style, int columns, int headerCount, int footerStart) {
            for (int row = headerCount; row < footerStart; row++) {
                foreach (TableCellLayout cell in GetTableCellLayouts(table, row, columns)) {
                    if (cell.RowSpan > 1 && cell.Viewport == null && cell.TextRotation == 0 &&
                        cell.Images.Count == 0 && cell.CheckBoxes.Count == 0 && cell.FormFields.Count == 0 &&
                        style.CellDataBars?.ContainsKey((row, cell.Column)) != true &&
                        style.CellIcons?.ContainsKey((row, cell.Column)) != true &&
                        !Enumerable.Range(row, cell.RowSpan).Any(spanRow => GetTableRowFixedHeight(style, spanRow).HasValue ||
                            !GetTableRowAllowBreakAcrossPages(style, spanRow))) {
                        var flow = new TableSpanCellFlow(row, cell);
                        Cells.Add(flow);
                        anchors.Add((row, cell.Column), flow);
                        for (int spanRow = row; spanRow < flow.EndRow; spanRow++) coveredRows.Add(spanRow);
                    }
                }
            }
        }

        public bool Contains(int row, int column) => anchors.ContainsKey((row, column));
        public int GetConsumedLines(int row, int column) => anchors.TryGetValue((row, column), out var span) ? span.ConsumedLines : 0;
        public bool CoversRow(int row) => coveredRows.Contains(row);

        public void AdmitRow(int row, double top, double bottom, double rowGap) {
            ActiveCells.RemoveAll(cell => cell.EndRow <= row);
            while (nextCell < Cells.Count && Cells[nextCell].Row <= row) {
                TableSpanCellFlow cell = Cells[nextCell++];
                if (cell.EndRow > row) ActiveCells.Add(cell);
            }
            foreach (TableSpanCellFlow cell in ActiveCells) {
                if (cell.Height <= .001D) {
                    // A gap between physical rows can fall outside both page fragments.
                    // Count it once in logical alignment, never between pieces of one row.
                    if (cell.ProcessedHeight > 0D && cell.LastFragmentRow < row)
                        cell.ProcessedHeight += (row - cell.LastFragmentRow) * rowGap;
                    cell.Top = top;
                    cell.FirstFragmentRow = row;
                    cell.FragmentRowHeights.Clear();
                    cell.FragmentRowHeights.Add(top - bottom);
                } else if (cell.LastFragmentRow == row) {
                    cell.FragmentRowHeights[cell.FragmentRowHeights.Count - 1] += cell.Top - bottom - cell.Height;
                } else {
                    cell.FragmentRowHeights[cell.FragmentRowHeights.Count - 1] += cell.Top - cell.Height - top;
                    cell.FragmentRowHeights.Add(top - bottom);
                }
                cell.LastFragmentRow = row;
                cell.Height = cell.Top - bottom;
            }
        }

        public void BindStructureRow(LayoutResult.Page? page, PageStructElement? rowElement) {
            foreach (TableSpanCellFlow cell in ActiveCells) {
                if (ReferenceEquals(cell.StructurePage, page)) continue;
                cell.StructurePage = page;
                cell.RowStructure = rowElement;
            }
        }
    }

    private sealed class TableSpanCellFlow {
        public TableSpanCellFlow(int row, TableCellLayout cell) { Row = row; Cell = cell; }
        public int Row { get; }
        public TableCellLayout Cell { get; }
        public int EndRow => Row + Cell.RowSpan;
        public int ConsumedLines { get; set; }
        public double ProcessedHeight { get; set; }
        public double? AlignmentLeadingHeight { get; set; }
        public double Top { get; set; }
        public double Height { get; set; }
        public int FirstFragmentRow { get; set; }
        public int LastFragmentRow { get; set; }
        public List<double> FragmentRowHeights { get; } = new();
        public LayoutResult.Page? StructurePage { get; set; }
        public PageStructElement? RowStructure { get; set; }
        public bool DestinationEmitted { get; set; }
    }
}
