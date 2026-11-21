namespace OfficeIMO.Toml;

/// <summary>
/// Settings for <see cref="TomlReader"/>. The defaults follow TOML 1.1 and bound resource use. An instance can be shared
/// by readers on several threads as long as it is not modified while they run.
/// </summary>
public sealed class TomlReaderOptions {
    private int _maxDepth = 256;
    private int _maxStringLength = 16 * 1024 * 1024;

    /// <summary>Gets or sets whether <see cref="TomlReader.Read"/> stops on comments. Defaults to <see langword="false"/>.</summary>
    public bool IncludeComments { get; set; }

    /// <summary>Gets or sets the maximum nesting of arrays and inline tables. Defaults to 256.</summary>
    /// <exception cref="ArgumentOutOfRangeException">The value is less than 1.</exception>
    public int MaxDepth {
        get => _maxDepth;
        set => _maxDepth = value >= 1 ? value : throw new ArgumentOutOfRangeException(nameof(value), value, "MaxDepth must be at least 1.");
    }

    /// <summary>
    /// Gets or sets the maximum length of one string or key token's source text in UTF-16 code units, including its
    /// quotes. Defaults to 16 MiB of characters.
    /// </summary>
    /// <exception cref="ArgumentOutOfRangeException">The value is less than 1.</exception>
    public int MaxStringLength {
        get => _maxStringLength;
        set => _maxStringLength = value >= 1 ? value : throw new ArgumentOutOfRangeException(nameof(value), value, "MaxStringLength must be at least 1.");
    }

    /// <summary>
    /// Gets or sets the maximum input size, or <see langword="null"/> (the default) for no limit. Stream input counts
    /// bytes; text and <see cref="TextReader"/> input count UTF-16 code units.
    /// </summary>
    /// <exception cref="ArgumentOutOfRangeException">The value is negative.</exception>
    public long? MaxInputBytes {
        get => _maxInputBytes;
        set => _maxInputBytes = value is null or >= 0 ? value : throw new ArgumentOutOfRangeException(nameof(value), value, "MaxInputBytes cannot be negative.");
    }

    private long? _maxInputBytes;

    internal TomlReaderOptions Clone() => (TomlReaderOptions)MemberwiseClone();
}
