namespace OfficeIMO.Publisher.Internal;

/// <summary>A table-owned utility shape: a cell fill by definition index, or an edge between grid boundaries.</summary>
internal sealed class PublisherTableArtwork {
    internal PublisherTableArtwork(uint tableId, IReadOnlyDictionary<ushort, uint> anchor, PublisherEscherShape shape) {
        TableId = tableId; Anchor = anchor; Shape = shape;
    }
    internal uint TableId { get; }
    internal IReadOnlyDictionary<ushort, uint> Anchor { get; }
    internal PublisherEscherShape Shape { get; }
    internal uint Value(ushort key) => Anchor.TryGetValue(key, out uint value) ? value : 0;
}
