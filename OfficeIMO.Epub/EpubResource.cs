namespace OfficeIMO.Epub;

/// <summary>Manifest resource discovered in an EPUB package.</summary>
public sealed class EpubResource {
    private byte[]? _data;

    /// <summary>OPF manifest item identifier.</summary>
    public string Id { get; internal set; } = string.Empty;

    /// <summary>Normalized archive path or absolute remote URI.</summary>
    public string Path { get; internal set; } = string.Empty;

    /// <summary>Original manifest href.</summary>
    public string? Href { get; internal set; }

    /// <summary>Whether this manifest item references an absolute remote resource.</summary>
    public bool IsRemote { get; internal set; }

    /// <summary>Absolute remote URI when <see cref="IsRemote"/> is true.</summary>
    public string? RemoteUri { get; internal set; }

    /// <summary>Declared media type.</summary>
    public string? MediaType { get; internal set; }

    /// <summary>Space-separated OPF properties.</summary>
    public string? Properties { get; internal set; }

    /// <summary>Uncompressed resource length.</summary>
    public long LengthBytes { get; internal set; }

    /// <summary>Encryption declaration for this resource, when present.</summary>
    public EpubEncryptionInfo? Encryption { get; internal set; }

    /// <summary>Optional bounded resource payload requested by the caller.</summary>
    public byte[]? Data {
        get => _data == null ? null : (byte[])_data.Clone();
        internal set => _data = value == null ? null : (byte[])value.Clone();
    }
}
