namespace OfficeIMO.Publisher.Internal;

internal sealed partial class PublisherPublicationState {
    /// <summary>Changes authored cell ownership and its complete story partition together; native topology remains refused.</summary>
    internal PublisherPublicationState MergeTableCells(uint id, int row, int column, int rowCount, int columnCount,
        PublisherReadOptions options, CancellationToken token) {
        PublisherTableState table = Table(id);
        if (!table.IsAuthored) throw new NotSupportedException("Merging imported native table cells is not qualified.");
        if (row < 0 || row >= table.Rows.Count) throw new ArgumentOutOfRangeException(nameof(row));
        if (column < 0 || column >= table.Columns.Count) throw new ArgumentOutOfRangeException(nameof(column));
        if (rowCount <= 0 || rowCount > table.Rows.Count - row) throw new ArgumentOutOfRangeException(nameof(rowCount));
        if (columnCount <= 0 || columnCount > table.Columns.Count - column) throw new ArgumentOutOfRangeException(nameof(columnCount));
        string? limitation = CellTextEditLimitation(table);
        if (limitation != null) throw new NotSupportedException(limitation);

        var selected = new List<int>();
        for (int i = 0; i < table.Cells.Count; i++) {
            token.ThrowIfCancellationRequested();
            PublisherTableCellSlot cell = table.Cells[i];
            if (cell.Row >= row + rowCount || cell.Column >= column + columnCount
                || cell.Row + cell.RowSpan <= row || cell.Column + cell.ColumnSpan <= column) continue;
            if (cell.Row < row || cell.Column < column || cell.Row + cell.RowSpan > row + rowCount
                || cell.Column + cell.ColumnSpan > column + columnCount)
                throw new ArgumentException("The merge range must contain complete cells; it cannot split an existing merged cell.");
            selected.Add(i);
        }
        if (selected.Count == 1) return this;
        if (table.BorderSegments.Count != 0)
            throw new NotSupportedException("Merging a table with shared borders requires a separately qualified border-topology edit.");
        uint storyId = table.StoryId!.Value;
        PublisherTableStory original = Text.TableStories[storyId];
        var selectedSet = new HashSet<int>(selected);
        var paragraphs = selected.SelectMany(index => original.Segments[index].Paragraphs).ToArray();
        PublisherTextStory merged = PublisherStoryEditor.Replace(original.Segments[selected[0]], paragraphs, options, token);
        var cells = new List<PublisherTableCellSlot>();
        var segments = new List<PublisherTextStory>();
        for (int i = 0; i < table.Cells.Count; i++) {
            if (i == selected[0]) {
                cells.Add(new PublisherTableCellSlot(row, column, rowCount, columnCount));
                segments.Add(merged);
            } else if (!selectedSet.Contains(i)) {
                cells.Add(table.Cells[i]);
                segments.Add(original.Segments[i]);
            }
        }
        var partition = new PublisherTableStory(segments);
        PublisherTextStory complete = partition.Complete(Text.Stories[storyId], options, token);
        var tables = Tables.ToDictionary(item => item.Key, item => item.Value);
        tables[id] = table.WithCells(cells);
        return new PublisherPublicationState(Pages, ObjectOwners, Pictures, Images, Text.WithTableStory(complete, partition),
            TableStoryReferences, _sourceImageIds, NextImageId, NextObjectId, Shapes, tables);
    }
}
