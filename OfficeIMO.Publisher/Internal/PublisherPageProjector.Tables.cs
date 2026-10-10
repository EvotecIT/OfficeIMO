using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Internal;

internal sealed partial class PublisherPageProjector {
    private void ProjectTable(PublisherSourceShape native, uint storyId, PublisherEscherShape shape, OfficeDrawing drawing) {
        PublisherSourceTable? table = native.Table;
        if (table == null || !_text.CellParagraphs.TryGetValue(storyId, out var paragraphs) || paragraphs.Count != table.Cells.Count) {
            _context.Add("PUB_TABLE_TEXT_MAPPING_UNRESOLVED", "Recovered table text remains available in TextStories, but cannot be assigned to native cells.",
                OfficeConversionLossKind.Omission, PublisherEscherReader.ShapeLocation(shape.Id));
            return;
        }
        double[] xs = TrackOffsets(table.Columns), ys = TrackOffsets(table.Rows);
        for (int i = 0; i < table.Cells.Count; i++) {
            _context.Record();
            PublisherSourceCell cell = table.Cells[i];
            double x = xs[cell.FirstColumn], y = ys[cell.FirstRow];
            if (paragraphs[i].Count == 0) continue;
            ProjectText(paragraphs[i], drawing, shape, x, y, xs[cell.LastColumn + 1] - x, ys[cell.LastRow + 1] - y);
        }
        _context.Add("PUB_TABLE_CELL_STYLE_UNASSESSED", "Native cell fills, individual borders, and cell-specific padding require additional qualification. The recovered native grid and styled cell text are projected.",
            OfficeConversionLossKind.Unassessed, PublisherEscherReader.ShapeLocation(shape.Id));
    }
    private static double[] TrackOffsets(double[] sizes) {
        var offsets = new double[sizes.Length + 1];
        for (int i = 0; i < sizes.Length; i++) offsets[i + 1] = offsets[i] + sizes[i];
        return offsets;
    }
}
