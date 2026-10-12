using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Internal;

internal sealed partial class PublisherPageProjector {
    private void ProjectTable(PublisherSourceShape native, uint? storyId, PublisherEscherShape shape,
        FrameRectangle bounds, uint pageId, OfficeDrawing? drawing) {
        if (_tables.TryGetValue(shape.Id, out PublisherTable? recovered)) {
            if (drawing != null && recovered.HasTextMapping) {
                foreach (PublisherTableCell cell in recovered.Cells)
                    if (cell.Paragraphs.Count != 0) ProjectText(cell.Paragraphs, drawing, shape, cell.X, cell.Y, cell.Width, cell.Height);
                _usedStories.Add(storyId!.Value);
                _context.Add("PUB_TEXT_LAYOUT_APPROXIMATED", "Native text stories use the shared drawing paragraph layout. Font availability, line breaking, fitting, and overflow may differ from Publisher.",
                    OfficeConversionLossKind.Approximation, "Quill");
            }
            return;
        }
        PublisherSourceTable? table = native.Table;
        if (table == null) return;
        List<IReadOnlyList<OfficeRichTextParagraph>>? paragraphs = null;
        bool hasStory = storyId.HasValue && _text.Stories.ContainsKey(storyId.Value);
        bool mapped = hasStory && _text.CellParagraphs.TryGetValue(storyId!.Value, out paragraphs) && paragraphs.Count == table.Cells.Count;
        if (!hasStory) {
            _context.Add("PUB_TEXT_REFERENCE_UNRESOLVED", "A native table has no available text story. Its grid is retained without cell text.",
                OfficeConversionLossKind.Omission, PublisherEscherReader.ShapeLocation(shape.Id));
        } else if (!mapped) {
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
                if (paragraph.Label != null) characters += paragraph.Label.Run.Text.Length
                    + (paragraph.Label.FollowedBy == OfficeTextParagraphLabelFollowedBy.Nothing ? 0 : 1);
            }
            _context.AccountTableText(characters);
            cells.Add(new PublisherTableCell(cell.FirstRow, cell.FirstColumn,
                cell.LastRow - cell.FirstRow + 1, cell.LastColumn - cell.FirstColumn + 1, x, y, width, height, text));
        }
        bool hasPlacement = shape.Bounds.HasValue && bounds.Width > 0 && bounds.Height > 0
            && FiniteRectangle(bounds.X, bounds.Y, bounds.Width, bounds.Height)
            && shape.LayoutTransform.TryInvert(out _) && shape.GroupTransform.TryInvert(out _);
        if (!hasPlacement)
            _context.Add("PUB_TABLE_PLACEMENT_UNRESOLVED", "The native table grid and mapped cell text are retained without available page placement.",
                OfficeConversionLossKind.Omission, PublisherEscherReader.ShapeLocation(shape.Id));
        _tables.Add(shape.Id, new PublisherTable(shape.Id, storyId, pageId, bounds.X, bounds.Y, bounds.Width, bounds.Height,
            hasPlacement ? OfficeTransform.Translate(bounds.X, bounds.Y).Then(PageTransform(shape, bounds)) : OfficeTransform.Identity,
            table.Columns, table.Rows, cells, mapped, hasPlacement));
        _context.Add("PUB_TABLE_CELL_STYLE_UNASSESSED", "Native cell fills, individual borders, and cell-specific padding require additional qualification. The table model retains the native grid and any mapped styled cell text.",
            OfficeConversionLossKind.Unassessed, PublisherEscherReader.ShapeLocation(shape.Id));

    }
    private static double[] TrackOffsets(double[] sizes) {
        var offsets = new double[sizes.Length + 1];
        for (int i = 0; i < sizes.Length; i++) offsets[i + 1] = offsets[i] + sizes[i];
        return offsets;
    }
}
