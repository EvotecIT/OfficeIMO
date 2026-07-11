namespace OfficeIMO.Epub;

/// <summary>Manifest resource discovered in an EPUB package.</summary>
public sealed class EpubResource {
    /// <summary>OPF manifest item identifier.</summary>
    public string Id { get; set; } = string.Empty;

    /// <summary>Normalized archive path.</summary>
    public string Path { get; set; } = string.Empty;

    /// <summary>Declared media type.</summary>
    public string? MediaType { get; set; }

    /// <summary>Space-separated OPF properties.</summary>
    public string? Properties { get; set; }

    /// <summary>Uncompressed resource length.</summary>
    public long LengthBytes { get; set; }

    /// <summary>Optional bounded resource payload requested by the caller.</summary>
    public byte[]? Data { get; set; }
}
