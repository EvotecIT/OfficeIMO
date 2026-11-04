namespace OfficeIMO.Publisher.Internal;

/// <summary>Complete source-bound utility removals and additions at original container-body offsets.</summary>
internal sealed class PublisherEscherShapeEdits {
    internal HashSet<int> Removed { get; } = new();
    internal Dictionary<int, byte[]> Appended { get; } = new();
    internal PublisherNativeStreamEdit? AllocationRecord { get; init; }
    internal PublisherNativeStreamEdit? DrawingInsertion { get; init; }
    internal bool IsEmpty => Removed.Count == 0 && Appended.Count == 0 && AllocationRecord == null && DrawingInsertion == null;
}
