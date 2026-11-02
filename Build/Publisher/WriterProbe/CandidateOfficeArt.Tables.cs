using OfficeIMO.Publisher.Internal;
using static OfficeIMO.Publisher.WriterProbe.NativePublicationGraph;

namespace OfficeIMO.Publisher.WriterProbe;

internal static partial class CandidateOfficeArt {
    /// <summary>Consumes the shared source-free table utility encoder for the internal authoring specimen.</summary>
    private static (byte[] Bytes, uint Count, uint LastSpid)? TableDrawing(Graph graph) {
        ObjectBinding[] tables = graph.Objects.Where(item => item.Source.Table is { } table
            && (table.Cells.All(cell => cell.RowSpan == 1 && cell.ColumnSpan == 1)
                || table.Cells.Any(cell => cell.FillColor.HasValue) || table.BorderSegments.Count != 0)).ToArray();
        if (tables.Length == 0) return null;
        var inputs = new List<PublisherTableUtilityInput>();
        foreach (ObjectBinding item in tables) {
            PublisherTable table = item.Source.Table!;
            if (table.Cells.Any(cell => cell.RowSpan != 1 || cell.ColumnSpan != 1))
                throw new NotSupportedException("Decorated merged-table native creation requires separate qualification.");
            inputs.Add(new(item.Id, table.RowHeights.Count, table.ColumnWidths.Count, table.Width, table.Height,
                table.Cells.Select(cell => cell.FillColor).ToArray(), table.BorderSegments));
        }
        PublisherTableUtilityDrawing drawing = PublisherTableUtilityDrawingEncoding.Write(3, inputs, 1024, CancellationToken.None)
            ?? throw new NotSupportedException("Native table utilities require opaque solid paint and exact positive signed-EMU borders within one SPID cluster.");
        return (drawing.Bytes, drawing.Count, drawing.LastSpid);
    }
}
