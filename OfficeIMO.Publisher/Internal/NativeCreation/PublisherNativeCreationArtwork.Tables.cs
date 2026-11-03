using OfficeIMO.Publisher.Internal;
using static OfficeIMO.Publisher.Internal.PublisherNativeCreationGraph;

namespace OfficeIMO.Publisher.Internal;

internal static partial class PublisherNativeCreationArtwork {
    /// <summary>Consumes the shared source-free table utility encoder for the internal authoring specimen.</summary>
    private static PublisherTableUtilityDrawing? TableDrawing(Graph graph, PublisherEscherDrawingAllocation allocation, PublisherParseContext context) {
        context.Token.ThrowIfCancellationRequested();
        var inputs = new List<PublisherTableUtilityInput>();
        foreach (ObjectBinding item in graph.Objects) {
            context.Token.ThrowIfCancellationRequested();
            if (item.Source.Table is not { } table) continue;
            bool unmerged = true, hasFill = false;
            foreach (PublisherTableCell cell in table.Cells) {
                context.Token.ThrowIfCancellationRequested();
                unmerged &= cell.RowSpan == 1 && cell.ColumnSpan == 1;
                hasFill |= cell.FillColor.HasValue;
            }
            if (!unmerged && !hasFill && table.BorderSegments.Count == 0) continue;
            var slots = new PublisherTableCellSlot[table.Cells.Count];
            var fills = new OfficeIMO.Drawing.OfficeColor?[table.Cells.Count];
            for (int index = 0; index < table.Cells.Count; index++) {
                context.Token.ThrowIfCancellationRequested();
                PublisherTableCell cell = table.Cells[index];
                slots[index] = new(cell.RowIndex, cell.ColumnIndex, cell.RowSpan, cell.ColumnSpan);
                fills[index] = cell.FillColor;
            }
            inputs.Add(new(item.Id, table.RowHeights.Count, table.ColumnWidths.Count, table.Width, table.Height,
                slots, fills, table.BorderSegments));
        }
        if (inputs.Count == 0) return null;
        int available = context.Options.Limits.MaxItems - graph.Objects.Count - 2;
        if (!allocation.StartDrawing(2, 1, context.Token))
            throw new NotSupportedException("The auxiliary drawing requires fresh allocation history.");
        return PublisherTableUtilityDrawingEncoding.Write(allocation, 3, inputs, available, context.Token)
            ?? throw new NotSupportedException("Native table utilities require opaque solid paint, exact positive signed-EMU borders and representable fresh identifiers within the item limit.");
    }
}
