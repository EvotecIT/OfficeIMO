namespace OfficeIMO.Yaml;

/// <summary>
/// Settings for <see cref="YamlReader"/>. The defaults read YAML 1.2 as real files use it (see <see cref="Strict"/>) and
/// bound resource use for untrusted input.
/// An instance can be shared by readers on several threads as long as it is not modified while they run.
/// </summary>
public sealed class YamlReaderOptions {
    private int _maxDepth = 256;
    private int _maxNodes = 1_000_000;
    private int _maxAliasExpansion = 100_000;
    private int _maxScalarLength = 16 * 1024 * 1024;
    private long? _maxInputBytes;

    /// <summary>Gets the default options. The returned instance is shared and must not be modified.</summary>
    internal static YamlReaderOptions Default { get; } = new YamlReaderOptions();

    /// <summary>
    /// Gets new options for input from outside, with the limits of <see cref="YamlLoadOptions.UntrustedDefaults"/>: input
    /// up to 16 MiB, nesting up to 64, alias expansion up to 10,000 nodes and scalars up to 1 MiB of characters. Each
    /// call returns a new instance.
    /// </summary>
    public static YamlReaderOptions UntrustedDefaults => new YamlReaderOptions {
        MaxInputBytes = 16 * 1024 * 1024,
        MaxDepth = 64,
        MaxNodes = 200_000,
        MaxAliasExpansion = 10_000,
        MaxScalarLength = 1024 * 1024
    };

    /// <summary>
    /// Gets or sets the schema used by the typed accessors. When <see langword="null"/> (the default), each document
    /// uses <see cref="YamlSchema.Core"/>, or <see cref="YamlSchema.Yaml11"/> when it declares <c>%YAML 1.1</c>.
    /// Set it to force one schema for ecosystems that assume YAML 1.1 without declaring it.
    /// </summary>
    public YamlSchema? Schema { get; set; }

    /// <summary>
    /// Gets or sets whether the reader follows the YAML 1.2.2 specification exactly. Off by default, the reader also
    /// accepts the indentation that libyaml and SharpYaml accept and real files rely on: lines inside a flow
    /// collection (such as a closing <c>]</c> under its key) and quoted-scalar continuation lines indented no deeper than
    /// the enclosing block collection. Each accepted occurrence is reported in <see cref="YamlReader.Diagnostics"/>.
    /// When on, both forms are syntax errors, as the specification requires.
    /// </summary>
    public bool Strict { get; set; }

    /// <summary>
    /// Gets or sets whether the reader reports <see cref="YamlTokenType.Comment"/> events, in source order.
    /// Off by default; when off, comments cost nothing beyond skipping them.
    /// </summary>
    public bool IncludeComments { get; set; }

    /// <summary>
    /// Gets or sets the maximum nesting of collections. Block and flow collections count the same, so <c>a: {b: 1}</c>
    /// and a block mapping inside a block mapping both nest two deep. Defaults to 256.
    /// </summary>
    public int MaxDepth {
        get => _maxDepth;
        set => _maxDepth = RequirePositive(value, nameof(MaxDepth));
    }

    /// <summary>
    /// Gets or sets the maximum number of nodes one document of a loaded model may hold. Defaults to 1,000,000.
    /// The forward-only reader does not enforce it, so streaming consumers can read unbounded multi-document streams.
    /// </summary>
    public int MaxNodes {
        get => _maxNodes;
        set => _maxNodes = RequirePositive(value, nameof(MaxNodes));
    }

    /// <summary>
    /// Gets or sets the maximum number of nodes visited while resolving aliases, which stops "billion laughs" inputs in
    /// resolved views and deserialization. Defaults to 100,000. The reader reports aliases without expanding them.
    /// </summary>
    public int MaxAliasExpansion {
        get => _maxAliasExpansion;
        set => _maxAliasExpansion = RequirePositive(value, nameof(MaxAliasExpansion));
    }

    /// <summary>
    /// Gets or sets the maximum length of one scalar's source text in UTF-16 code units, including quotes and block
    /// scalar indentation. Defaults to 16 MiB of characters.
    /// </summary>
    public int MaxScalarLength {
        get => _maxScalarLength;
        set => _maxScalarLength = RequirePositive(value, nameof(MaxScalarLength));
    }

    /// <summary>
    /// Gets or sets the maximum input size, or <see langword="null"/> (the default) for no limit. Stream input counts
    /// bytes; string, memory and <see cref="TextReader"/> input count UTF-16 code units.
    /// </summary>
    public long? MaxInputBytes {
        get => _maxInputBytes;
        set {
            if (value is <= 0) throw new ArgumentOutOfRangeException(nameof(MaxInputBytes), value, "The limit must be positive.");
            _maxInputBytes = value;
        }
    }

    private static int RequirePositive(int value, string name) {
        if (value <= 0) throw new ArgumentOutOfRangeException(name, value, "The limit must be positive.");
        return value;
    }
}
