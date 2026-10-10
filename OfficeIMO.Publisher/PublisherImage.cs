namespace OfficeIMO.Publisher;

/// <summary>An embedded image recovered without resolving links or activating embedded objects.</summary>
public sealed class PublisherImage {
    private readonly byte[] _bytes;
    internal PublisherImage(int id, string contentType, byte[] bytes) { Id = id; ContentType = contentType; _bytes = bytes; }
    /// <summary>One-based image-store identifier used by native drawing properties.</summary>
    public int Id { get; }
    /// <summary>MIME type of the recovered payload, including Publisher-specific image envelopes.</summary>
    public string ContentType { get; }
    /// <summary>Length of the recovered image payload.</summary>
    public int ByteCount => _bytes.Length;
    /// <summary>Returns a detached copy of the recovered payload.</summary>
    public byte[] GetBytes() => (byte[])_bytes.Clone();
    internal byte[] Bytes => _bytes;
}
