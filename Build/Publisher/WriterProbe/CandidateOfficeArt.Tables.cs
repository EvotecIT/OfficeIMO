using OfficeIMO.Publisher.Internal;
using static OfficeIMO.Publisher.WriterProbe.NativePublicationGraph;

namespace OfficeIMO.Publisher.WriterProbe;

internal static partial class CandidateOfficeArt {
    /// <summary>Consumes the shared source-free table utility encoder for the internal authoring specimen.</summary>
    private static PublisherTableUtilityDrawing? TableDrawing(Graph graph, PublisherEscherDrawingAllocation allocation) {
        ObjectBinding[] tables = graph.Objects.Where(item => item.Source.Table is { } table
            && (table.Cells.All(cell => cell.RowSpan == 1 && cell.ColumnSpan == 1)
                || table.Cells.Any(cell => cell.FillColor.HasValue) || table.BorderSegments.Count != 0)).ToArray();
        if (tables.Length == 0) return null;
        var inputs = new List<PublisherTableUtilityInput>();
        foreach (ObjectBinding item in tables) {
            PublisherTable table = item.Source.Table!;
            inputs.Add(new(item.Id, table.RowHeights.Count, table.ColumnWidths.Count, table.Width, table.Height,
                table.Cells.Select(cell => new PublisherTableCellSlot(cell.RowIndex, cell.ColumnIndex, cell.RowSpan, cell.ColumnSpan)).ToArray(),
                table.Cells.Select(cell => cell.FillColor).ToArray(), table.BorderSegments));
        }
        int available = new PublisherReadOptions().Limits.MaxItems - graph.Objects.Count - 2;
        if (!allocation.StartDrawing(2, 1, CancellationToken.None))
            throw new NotSupportedException("The auxiliary drawing requires fresh allocation history.");
        return PublisherTableUtilityDrawingEncoding.Write(allocation, 3, inputs, available, CancellationToken.None)
            ?? throw new NotSupportedException("Native table utilities require opaque solid paint, exact positive signed-EMU borders and representable fresh identifiers within the item limit.");
    }
}
