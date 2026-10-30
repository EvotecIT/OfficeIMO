using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Internal;

internal sealed partial class PublisherPublicationState {
    internal PublisherTableState Table(uint id) => Tables.TryGetValue(id, out PublisherTableState? table) ? table
        : throw new ArgumentOutOfRangeException(nameof(id), "The publication has no such table.");

    internal string? CellTextEditLimitation(PublisherTableState table) {
        if (!table.HasTextMapping(Text)) return "The table has no complete resolved cell-to-story mapping.";
        uint storyId = table.StoryId!.Value;
        if (!TableStoryReferences.TryGetValue(table.Id, out uint? declaredStory) || declaredStory != storyId)
            return "The table's complete story ownership is unresolved.";
        if (!Text.TableStories[storyId].Matches(Text.Stories[storyId]))
            return "The native cell boundaries cannot be safely represented by the complete styled story.";
        if (Text.Frames.Values.Any(frame => frame.StoryId == storyId))
            return "The table story is also referenced by an ordinary text frame.";
        if (TableStoryReferences.Any(reference => reference.Value == storyId
            && (!Tables.TryGetValue(reference.Key, out PublisherTableState? other) || other.StoryId != storyId || !other.HasTextMapping(Text))))
            return "Another table references this story with a conflicting or unresolved cell grid or page ownership.";
        return null;
    }

    internal PublisherPublicationState ReplaceTableCell(uint tableId, int row, int column,
        IReadOnlyList<OfficeRichTextParagraph> paragraphs, PublisherReadOptions options, CancellationToken token) {
        PublisherTableState table = Table(tableId);
        string? limitation = CellTextEditLimitation(table);
        if (limitation != null) throw new NotSupportedException(limitation);
        if (row < 0 || row >= table.Rows.Count) throw new ArgumentOutOfRangeException(nameof(row));
        if (column < 0 || column >= table.Columns.Count) throw new ArgumentOutOfRangeException(nameof(column));
        int index = -1;
        for (int i = 0; i < table.Cells.Count; i++) {
            token.ThrowIfCancellationRequested();
            if (table.Cells[i].Row == row && table.Cells[i].Column == column) { index = i; break; }
        }
        if (index < 0) throw new ArgumentException("Use the first row and column of a recovered cell; covered merged positions and unresolved grid positions cannot be edited.", nameof(column));
        uint storyId = table.StoryId!.Value;
        PublisherTableStory partition = Text.TableStories[storyId].Replace(index, paragraphs, options, token);
        PublisherTextStory story = partition.Complete(Text.Stories[storyId], options, token);
        if (Text.Stories.Values.Sum(item => PublisherStoryEditor.CharacterCount(item.Id == storyId ? story : item)) > options.Limits.MaxTextCharacters)
            throw new ArgumentException("Publication story text exceeds the configured character limit.", nameof(paragraphs));
        return WithText(Text.WithTableStory(story, partition));
    }

    internal PublisherPublicationState SetTableBounds(uint id, double x, double y, double width, double height) {
        PublisherTableState table = Table(id);
        if (table.BoundsEditLimitation != null) throw new NotSupportedException(table.BoundsEditLimitation);
        ValidateTableBounds(x, y, width, height);
        Pages.Single(page => page.Id == table.PageId).FrameEditRange.Validate(x, y, width, height);
        var tables = Tables.ToDictionary(item => item.Key, item => item.Value);
        tables[id] = table.WithBounds(x, y, width, height);
        return new PublisherPublicationState(Pages, ObjectOwners, Pictures, Images, Text, TableStoryReferences, _sourceImageIds, NextImageId, NextObjectId, Shapes, tables);
    }

    internal PublisherPublicationState AddTable(uint pageId, double x, double y, double[] columns, double[] rows,
        PublisherReadOptions options, CancellationToken token, out uint tableId) {
        PublisherPageState? page = Pages.SingleOrDefault(item => item.Id == pageId);
        if (page == null) throw new ArgumentOutOfRangeException(nameof(pageId), "The publication has no such page.");
        if (!page.SupportsPlacementEditing) throw new NotSupportedException("This page's coordinate registration is not qualified for table creation.");
        double width = PublisherTableState.ValidateTracks(columns, nameof(columns)), height = PublisherTableState.ValidateTracks(rows, nameof(rows));
        ValidateTableBounds(x, y, width, height); page.FrameEditRange.Validate(x, y, width, height);
        long cellCount = (long)columns.Length * rows.Length;
        if (cellCount > options.Limits.MaxItems || 1L + columns.Length + rows.Length + cellCount > options.Limits.MaxItems)
            throw new InvalidDataException("Publisher table grid exceeds the configured item limit.");
        if (cellCount > 4096 || cellCount > options.Limits.MaxTextCharacters)
            throw new ArgumentException("The empty table story exceeds the paragraph or character limit.", nameof(rows));
        if (Text.Stories.Count >= options.Limits.MaxItems) throw new InvalidDataException("Publisher story limit exceeded.");
        // A table and its complete story consume two identities from the same
        // allocator. Nothing is published until both reservations and projection succeed.
        if (NextObjectId == 0 || NextObjectId == uint.MaxValue)
            throw new InvalidOperationException("The publication has insufficient identifiers for a table and its story.");
        tableId = NextObjectId; uint storyId = tableId + 1;
        var cells = new List<PublisherTableCellSlot>((int)cellCount);
        var segments = new List<PublisherTextStory>((int)cellCount);
        for (int row = 0; row < rows.Length; row++) {
            for (int column = 0; column < columns.Length; column++) {
                token.ThrowIfCancellationRequested();
                cells.Add(new PublisherTableCellSlot(row, column, 1, 1));
                segments.Add(new PublisherTextStory(storyId, "\n", new[] { new OfficeRichTextParagraph(Array.Empty<OfficeRichTextRun>()) }, true));
            }
        }
        var partition = new PublisherTableStory(segments);
        var original = new PublisherTextStory(storyId, string.Empty, Array.Empty<OfficeRichTextParagraph>(), true);
        PublisherTextStory story = partition.Complete(original, options, token);
        if (Text.Stories.Values.Sum(PublisherStoryEditor.CharacterCount) + PublisherStoryEditor.CharacterCount(story) > options.Limits.MaxTextCharacters)
            throw new ArgumentException("The table story exceeds the publication character limit.", nameof(rows));
        var tables = Tables.ToDictionary(item => item.Key, item => item.Value);
        tables.Add(tableId, new PublisherTableState(tableId, storyId, pageId, x, y, width, height,
            OfficeTransform.Translate(x, y), true, false, columns, rows, cells, OfficeTextPadding.Empty,
            OfficeShape.Rectangle(width, height), null, true));
        var owners = ObjectOwners.ToDictionary(item => item.Key, item => item.Value); owners.Add(tableId, pageId);
        var references = TableStoryReferences.ToDictionary(item => item.Key, item => item.Value); references.Add(tableId, storyId);
        uint createdId = tableId;
        return new PublisherPublicationState(Pages.Select(item => item.Id == pageId ? item.WithObject(createdId) : item),
            owners, Pictures, Images, Text.WithTableStory(story, partition), references, _sourceImageIds, NextImageId,
            storyId == uint.MaxValue ? 0 : storyId + 1, Shapes, tables);
    }

    private static void ValidateTableBounds(double x, double y, double width, double height) {
        if (!FiniteTable(x)) throw new ArgumentOutOfRangeException(nameof(x));
        if (!FiniteTable(y)) throw new ArgumentOutOfRangeException(nameof(y));
        if (!FiniteTable(width) || width <= 0) throw new ArgumentOutOfRangeException(nameof(width));
        if (!FiniteTable(height) || height <= 0) throw new ArgumentOutOfRangeException(nameof(height));
    }
    private static bool FiniteTable(double value) => !double.IsNaN(value) && !double.IsInfinity(value);
}
