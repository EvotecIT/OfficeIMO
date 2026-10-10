using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Internal;

internal sealed partial class PublisherPageProjector {
    private void ProjectTable(PublisherSourceShape native, uint? storyId, PublisherEscherShape shape,
        FrameRectangle bounds, uint pageId, OfficeDrawing drawing) {
        PublisherSourceTable? table = native.Table;
        if (table == null) return;
        List<IReadOnlyList<OfficeRichTextParagraph>>? paragraphs = null;
        bool mapped = storyId.HasValue && _text.Stories.ContainsKey(storyId.Value)
            && _text.CellParagraphs.TryGetValue(storyId.Value, out paragraphs) && paragraphs.Count == table.Cells.Count;
        if (!mapped) {
            _context.Add("PUB_TABLE_TEXT_MAPPING_UNRESOLVED", "Recovered table text remains available in TextStories, but cannot be assigned to native cells.",
                OfficeConversionLossKind.Omission, PublisherEscherReader.ShapeLocation(shape.Id));
        }
        _context.AccountTableModel(checked(1 + table.Columns.Length + table.Rows.Length + table.Cells.Count));
        double[] xs = TrackOffsets(table.Columns), ys = TrackOffsets(table.Rows);
        var cells = new List<PublisherTableCell>(table.Cells.Count);
        for (int i = 0; i < table.Cells.Count; i++) {
            _context.Record();
            PublisherSourceCell cell = table.Cells[i];
            double x = xs[cell.FirstColumn], y = ys[cell.FirstRow];
            double width = xs[cell.LastColumn + 1] - x, height = ys[cell.LastRow + 1] - y;
            IReadOnlyList<OfficeRichTextParagraph> text = mapped ? paragraphs![i] : Array.Empty<OfficeRichTextParagraph>();
            long characters = Math.Max(0, text.Count - 1);
            foreach (OfficeRichTextParagraph paragraph in text) {
                _context.Token.ThrowIfCancellationRequested();
                characters += paragraph.Runs.Sum(run => (long)run.Text.Length);
            }
            _context.AccountTableText(characters);
            cells.Add(new PublisherTableCell(cell.FirstRow, cell.FirstColumn,
                cell.LastRow - cell.FirstRow + 1, cell.LastColumn - cell.FirstColumn + 1, x, y, width, height, text));
            if (text.Count != 0) ProjectText(text, drawing, shape, x, y, width, height);
        }
        _tables.Add(shape.Id, new PublisherTable(shape.Id, storyId, pageId, bounds.X, bounds.Y, bounds.Width, bounds.Height,
            OfficeTransform.Translate(bounds.X, bounds.Y).Then(PageTransform(shape, bounds)),
            table.Columns, table.Rows, cells, mapped));
        if (!mapped) return;
        _usedStories.Add(storyId!.Value);
        _context.Add("PUB_TEXT_LAYOUT_APPROXIMATED", "Native text stories use the shared drawing paragraph layout. Font availability, line breaking, fitting, and overflow may differ from Publisher.",
            OfficeConversionLossKind.Approximation, "Quill");
        _context.Add("PUB_TABLE_CELL_STYLE_UNASSESSED", "Native cell fills, individual borders, and cell-specific padding require additional qualification. The recovered native grid and styled cell text are projected.",
            OfficeConversionLossKind.Unassessed, PublisherEscherReader.ShapeLocation(shape.Id));
    }
    private static double[] TrackOffsets(double[] sizes) {
        var offsets = new double[sizes.Length + 1];
        for (int i = 0; i < sizes.Length; i++) offsets[i + 1] = offsets[i] + sizes[i];
        return offsets;
    }
}
