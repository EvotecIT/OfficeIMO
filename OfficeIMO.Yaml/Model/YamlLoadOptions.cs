namespace OfficeIMO.Yaml;

/// <summary>How <see cref="YamlDocument"/> treats a mapping that contains two equal keys.</summary>
public enum YamlDuplicateKeyPolicy : byte {
    /// <summary>The load fails with <see cref="YamlErrorCodes.DuplicateKey"/>, as the YAML specification requires. This is the default.</summary>
    Error = 0,
    /// <summary>The load succeeds; key lookups return the first of the equal keys, and each later one is reported as a diagnostic.</summary>
    FirstWins,
    /// <summary>The load succeeds; key lookups return the last of the equal keys, and each later one is reported as a diagnostic.</summary>
    LastWins
}

/// <summary>
/// Settings for loading a <see cref="YamlDocument"/>. The defaults follow YAML 1.2 and bound resource use for untrusted
/// input. An instance can be shared by loads on several threads as long as it is not modified while they run.
/// </summary>
public sealed class YamlLoadOptions {
    private readonly YamlReaderOptions _reader = new YamlReaderOptions();

    /// <summary>Gets the default options. The returned instance is shared and must not be modified.</summary>
    internal static YamlLoadOptions Default { get; } = new YamlLoadOptions();

    /// <summary>
    /// Gets new options for files from outside: input up to 16 MiB, nesting up to 64, 200,000 nodes per document, alias
    /// expansion up to 10,000 nodes and scalars up to 1 MiB of characters. Each call returns a new instance.
    /// </summary>
    public static YamlLoadOptions UntrustedDefaults => new YamlLoadOptions {
        MaxInputBytes = 16 * 1024 * 1024,
        MaxDepth = 64,
        MaxNodes = 200_000,
        MaxAliasExpansion = 10_000,
        MaxScalarLength = 1024 * 1024
    };

    /// <summary>Returns a copy, so a caller's options can be adjusted without changing them.</summary>
    internal YamlLoadOptions Clone() => new YamlLoadOptions {
        Schema = Schema,
        DuplicateKeys = DuplicateKeys,
        MaxDepth = MaxDepth,
        MaxNodes = MaxNodes,
        MaxAliasExpansion = MaxAliasExpansion,
        MaxScalarLength = MaxScalarLength,
        MaxInputBytes = MaxInputBytes
    };

    /// <summary>Gets the reader settings the load runs with. Comments are never requested; the model reads them from the source.</summary>
    internal YamlReaderOptions ReaderOptions => _reader;

    /// <summary>
    /// Gets or sets the schema for typed reads. When <see langword="null"/> (the default), each document uses
    /// <see cref="YamlSchema.Core"/>, or <see cref="YamlSchema.Yaml11"/> when it declares <c>%YAML 1.1</c>.
    /// Set it to force one schema for ecosystems that assume YAML 1.1 without declaring it.
    /// </summary>
    public YamlSchema? Schema {
        get => _reader.Schema;
        set => _reader.Schema = value;
    }

    /// <summary>Gets or sets how a mapping with equal keys loads. Defaults to <see cref="YamlDuplicateKeyPolicy.Error"/>.</summary>
    public YamlDuplicateKeyPolicy DuplicateKeys { get; set; }

    /// <summary>Gets or sets the maximum nesting of collections and flow brackets. Defaults to 256.</summary>
    public int MaxDepth {
        get => _reader.MaxDepth;
        set => _reader.MaxDepth = value;
    }

    /// <summary>
    /// Gets or sets the maximum number of nodes (scalars, collections and aliases) in one document of the stream. Defaults
    /// to 1,000,000. A multi-document stream such as a large Kubernetes manifest list is bounded per document; bound the
    /// whole input with <see cref="MaxInputBytes"/>.
    /// </summary>
    public int MaxNodes {
        get => _reader.MaxNodes;
        set => _reader.MaxNodes = value;
    }

    /// <summary>
    /// Gets or sets the maximum number of nodes one document's aliases stand for when resolved, counting nested aliases
    /// at every use. Defaults to 100,000. The load measures it without expanding anything, so a "billion laughs" input
    /// fails in time proportional to its size, and every resolved view of a loaded document stays within the bound.
    /// </summary>
    public int MaxAliasExpansion {
        get => _reader.MaxAliasExpansion;
        set => _reader.MaxAliasExpansion = value;
    }

    /// <summary>
    /// Gets or sets the maximum length of one scalar's source text in UTF-16 code units, including quotes and block
    /// scalar indentation. Defaults to 16 MiB of characters.
    /// </summary>
    public int MaxScalarLength {
        get => _reader.MaxScalarLength;
        set => _reader.MaxScalarLength = value;
    }

    /// <summary>
    /// Gets or sets the maximum input size, or <see langword="null"/> (the default) for no limit. File and stream input
    /// counts bytes; <see cref="YamlDocument.Parse(string, YamlLoadOptions?, CancellationToken)"/> counts UTF-16 code units.
    /// </summary>
    public long? MaxInputBytes {
        get => _reader.MaxInputBytes;
        set => _reader.MaxInputBytes = value;
    }
}
