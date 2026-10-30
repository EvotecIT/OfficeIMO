namespace OfficeIMO.Publisher;

/// <summary>An embedded image asset, recovered or supplied by picture creation or a supported edit. Links and embedded objects remain inert.</summary>
public sealed class PublisherImage {
    private readonly byte[] _bytes;
    internal PublisherImage(int id, string contentType, byte[] bytes) { Id = id; ContentType = contentType; _bytes = bytes; }
    /// <summary>Document-scoped asset identifier. Loaded assets retain their one-based native image-store identifiers; codecs own native binding when saving is supported.</summary>
    public int Id { get; }
    /// <summary>MIME type of the current encoded payload, including recovered Publisher-specific image envelopes.</summary>
    public string ContentType { get; }
    /// <summary>Length of the encoded image payload.</summary>
    public int ByteCount => _bytes.Length;
    /// <summary>Returns a detached copy of the encoded payload.</summary>
    public byte[] GetBytes() => (byte[])_bytes.Clone();
    internal byte[] Bytes => _bytes;
}
