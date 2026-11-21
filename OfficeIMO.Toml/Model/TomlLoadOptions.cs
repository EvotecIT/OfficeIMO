namespace OfficeIMO.Toml;

/// <summary>
/// Settings for loading a <see cref="TomlDocument"/>. The defaults follow TOML 1.1 and bound resource use; use
/// <see cref="UntrustedDefaults"/> for files from outside. An instance can be shared by loads on several threads as long
/// as it is not modified while they run.
/// </summary>
public sealed class TomlLoadOptions {
    private readonly TomlReaderOptions _reader = new TomlReaderOptions();
    private int _maxNodes = 1_000_000;

    /// <summary>Gets the default options. The returned instance is shared and must not be modified.</summary>
    internal static TomlLoadOptions Default { get; } = new TomlLoadOptions();

    /// <summary>
    /// Gets new options with tighter limits for input from outside: 16 MiB of input, nesting depth 64, 200,000 nodes and
    /// 1 MiB per string. Each call returns a new instance that can be adjusted.
    /// </summary>
    public static TomlLoadOptions UntrustedDefaults => new TomlLoadOptions {
        MaxInputBytes = 16L * 1024 * 1024,
        MaxDepth = 64,
        MaxNodes = 200_000,
        MaxStringLength = 1024 * 1024
    };

    internal TomlReaderOptions ReaderOptions => _reader;

    /// <summary>
    /// Gets or sets the maximum nesting of tables, arrays and inline tables, counting each key of a dotted key or header.
    /// Defaults to 256.
    /// </summary>
    /// <exception cref="ArgumentOutOfRangeException">The value is less than 1.</exception>
    public int MaxDepth {
        get => _reader.MaxDepth;
        set => _reader.MaxDepth = value;
    }

    /// <summary>Gets or sets the maximum number of tables, arrays and values in the document. Defaults to 1,000,000.</summary>
    /// <exception cref="ArgumentOutOfRangeException">The value is less than 1.</exception>
    public int MaxNodes {
        get => _maxNodes;
        set => _maxNodes = value >= 1 ? value : throw new ArgumentOutOfRangeException(nameof(value), value, "MaxNodes must be at least 1.");
    }

    /// <summary>
    /// Gets or sets the maximum length of one string or key's source text in UTF-16 code units, including its quotes.
    /// Defaults to 16 MiB of characters.
    /// </summary>
    /// <exception cref="ArgumentOutOfRangeException">The value is less than 1.</exception>
    public int MaxStringLength {
        get => _reader.MaxStringLength;
        set => _reader.MaxStringLength = value;
    }

    /// <summary>
    /// Gets or sets the maximum input size, or <see langword="null"/> (the default) for no limit. File and stream input
    /// counts bytes; <see cref="TomlDocument.Parse(string, TomlLoadOptions?)"/> counts UTF-16 code units.
    /// </summary>
    /// <exception cref="ArgumentOutOfRangeException">The value is negative.</exception>
    public long? MaxInputBytes {
        get => _reader.MaxInputBytes;
        set => _reader.MaxInputBytes = value;
    }
}
