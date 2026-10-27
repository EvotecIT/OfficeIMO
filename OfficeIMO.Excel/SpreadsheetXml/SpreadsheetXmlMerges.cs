using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using static OfficeIMO.Excel.SpreadsheetXml.SpreadsheetXmlContext;

namespace OfficeIMO.Excel.SpreadsheetXml;

internal static class SpreadsheetXmlMerges {
    /// <summary>Materializes sparse merge/annotation anchors and rejects merges that would discard stored data.</summary>
    internal static Dictionary<string, (int across, int down)> Build(WorksheetPart part, IEnumerable<string> annotations,
        SpreadsheetXmlContext context, out IReadOnlyList<Row> result) {
        var annotationSet = new HashSet<string>(annotations, StringComparer.OrdinalIgnoreCase);
        var rows = new SortedDictionary<int, Row>();
        int previousRow = 0;
        foreach (Row source in part.Worksheet!.GetFirstChild<SheetData>()?.Elements<Row>() ?? Enumerable.Empty<Row>()) {
            context.Options.CancellationToken.ThrowIfCancellationRequested();
            int index = checked((int)(source.RowIndex?.Value ?? (uint)(previousRow + 1))); Coordinate(index, 1);
            if (index <= previousRow) throw new InvalidDataException("Worksheet row indexes must increase.");
            previousRow = index;
            foreach (Cell cell in source.Elements<Cell>()) context.Cell();
            var row = (Row)source.CloneNode(true); row.RowIndex = (uint)index;
            int column = 0;
            foreach (Cell cell in row.Elements<Cell>()) {
                int next = column + 1;
                if (cell.CellReference != null) {
                    ExcelReference reference = ExcelReference.Parse(cell.CellReference.Value!);
                    if (reference.Kind != ExcelReferenceKind.Cell || reference.Start.Row != index) throw new InvalidDataException("Cell address disagrees with its row.");
                    next = reference.Start.Column;
                }
                if (next <= column) throw new InvalidDataException("Worksheet cell indexes must increase.");
                Coordinate(index, next); column = next; cell.CellReference = ExcelA1.CellReference(index, column);
            }
            rows.Add(index, row);
        }
        var merges = new List<(int row, int column, int endRow, int endColumn)>();
        var map = new Dictionary<string, (int across, int down)>(StringComparer.OrdinalIgnoreCase);
        foreach (MergeCell merge in part.Worksheet.GetFirstChild<MergeCells>()?.Elements<MergeCell>() ?? Enumerable.Empty<MergeCell>()) {
            context.Options.CancellationToken.ThrowIfCancellationRequested();
            if (merges.Count >= context.Options.MaxCells) throw new InvalidDataException("Too many XML Spreadsheet merges.");
            if (!ExcelA1.TryParseRange(merge.Reference?.Value ?? "", out int row, out int column, out int endRow, out int endColumn)) throw new InvalidDataException("Invalid merged-cell reference.");
            Coordinate(row, column); Coordinate(endRow, endColumn);
            if (endRow < row || endColumn < column) throw new InvalidDataException("Invalid merged-cell order.");
            string reference = ExcelA1.CellReference(row, column);
            if (map.ContainsKey(reference)) throw new InvalidDataException("Duplicate merge anchor.");
            map.Add(reference, (endColumn - column, endRow - row));
            merges.Add((row, column, endRow, endColumn)); Anchor(reference);
        }
        foreach (string reference in annotationSet) Anchor(reference);
        var ordered = merges.OrderBy(item => item.row).ThenBy(item => item.column).ToArray();
        var active = new (int row, int column, int endRow)[257]; int mergeIndex = 0;
        foreach (var pair in rows) {
            context.Options.CancellationToken.ThrowIfCancellationRequested();
            int rowIndex = pair.Key;
            while (mergeIndex < ordered.Length && ordered[mergeIndex].row <= rowIndex) {
                var merge = ordered[mergeIndex++];
                for (int column = merge.column; column <= merge.endColumn; column++) {
                    if (active[column].endRow >= merge.row) throw new InvalidDataException("Overlapping merged-cell ranges.");
                    active[column] = (merge.row, merge.column, merge.endRow);
                }
            }
            foreach (Cell cell in pair.Value.Elements<Cell>().ToArray()) {
                int column = ExcelReference.Parse(cell.CellReference!.Value!).Start.Column;
                var merge = active[column];
                if (merge.endRow < rowIndex || merge.row == rowIndex && merge.column == column) continue;
                if (cell.CellFormula != null || cell.CellValue != null || cell.InlineString != null || annotationSet.Contains(cell.CellReference!.Value!))
                    throw new InvalidDataException("XML Spreadsheet cannot preserve data or annotations in covered merged cells.");
                if ((cell.StyleIndex?.Value ?? 0) != 0) context.Omitted("covered merged-cell formatting");
                cell.Remove();
            }
        }
        result = rows.Values.ToArray(); return map;

        void Anchor(string reference) {
            if (!ExcelReference.TryParse(reference, out var address) || address?.Kind != ExcelReferenceKind.Cell) throw new InvalidDataException("Invalid annotation address.");
            int row = address.Start.Row, column = address.Start.Column; Coordinate(row, column);
            if (!rows.TryGetValue(row, out Row? target)) rows.Add(row, target = new Row { RowIndex = (uint)row });
            var cells = target.Elements<Cell>().ToArray();
            if (cells.Any(cell => cell.CellReference?.Value == reference)) return;
            var before = cells.FirstOrDefault(cell => ExcelReference.Parse(cell.CellReference!.Value!).Start.Column > column);
            context.Cell();
            target.InsertBefore(new Cell { CellReference = ExcelA1.CellReference(row, column) }, before);
        }
    }
}
