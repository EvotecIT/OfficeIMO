namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private static PdfTableStyle PrepareDeferredPairedTableBorders(DeferredTableBatch batch, PdfTableStyle sourceStyle) {
        TableBlock table = batch.Table;
        int headerRows = table.Style!.HeaderRowCount;
        int footerStart = table.Rows.Count - table.Style.FooterRowCount;
        var rows = new List<PdfTableCell[]>();
        var sourceIndexes = new List<int>();
        var localIndexes = new Dictionary<int, int>();
        for (int row = 0; row <= table.Rows.Count; row++) {
            if (row == headerRows && batch.PreviousBodyRow is { } previous) AddNeighbour(previous);
            if (row == footerStart && batch.NextBodyRow is { } next) AddNeighbour(next);
            if (row == table.Rows.Count) break;
            localIndexes[rows.Count] = row;
            rows.Add(table.Cells[row].ToArray());
            sourceIndexes.Add(batch.SourceRowIndexes[row]);
        }
        // One preceding and one following row keep classification and incoming
        // clearance independent of batch size without materializing the table.
        PdfTableStyle contextStyle = sourceStyle.Clone();
        contextStyle.CellBorders = RemapSource(sourceStyle.CellBorders);
        contextStyle.CellPaddings = RemapSource(sourceStyle.CellPaddings);
        var contextTable = new TableBlock(rows, table.Align, contextStyle);
        PdfTableStyle context = PreparePairedTableBorders(contextTable, contextStyle);
        PdfTableStyle prepared = table.Style.Clone();
        prepared.CellBorders = SelectLocal(context.CellBorders);
        prepared.CellPaddings = SelectLocal(context.CellPaddings);
        return prepared;

        void AddNeighbour(DeferredTableBlock.IndexedTableRow neighbour) {
            rows.Add(neighbour.Cells);
            sourceIndexes.Add(neighbour.SourceIndex);
        }
        Dictionary<(int Row, int Column), T>? RemapSource<T>(Dictionary<(int Row, int Column), T>? values) {
            if (values == null) return null;
            var sourceToContext = sourceIndexes.Select((source, context) => (source, context)).ToDictionary(item => item.source, item => item.context);
            return values.Where(entry => sourceToContext.ContainsKey(entry.Key.Row))
                .ToDictionary(entry => (sourceToContext[entry.Key.Row], entry.Key.Column), entry => entry.Value);
        }
        Dictionary<(int Row, int Column), T>? SelectLocal<T>(Dictionary<(int Row, int Column), T>? values) =>
            values?.Where(entry => localIndexes.ContainsKey(entry.Key.Row))
                .ToDictionary(entry => (localIndexes[entry.Key.Row], entry.Key.Column), entry => entry.Value);
    }
}
