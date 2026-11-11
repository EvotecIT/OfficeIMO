namespace OfficeIMO.Yaml;

/// <summary>
/// How <see cref="YamlDocument"/> treats a mapping that contains two equal keys. Without an explicit policy, loads use
/// <see cref="LastWins"/>, and strict loads (<see cref="YamlLoadOptions.Strict"/>) use <see cref="Error"/>.
/// </summary>
public enum YamlDuplicateKeyPolicy : byte {
    /// <summary>The load fails with <see cref="YamlErrorCodes.DuplicateKey"/>, as the YAML specification requires. Strict loads use it by default.</summary>
    Error = 0,
    /// <summary>The load succeeds; key lookups return the first of the equal keys, and each later one is reported as a warning diagnostic.</summary>
    FirstWins,
    /// <summary>
    /// The load succeeds; key lookups, paths and the serializer use the last of the equal keys, as PyYAML, go-yaml v2 and
    /// SharpYaml do, and each later one is reported as a warning diagnostic. Non-strict loads use it by default.
    /// </summary>
    LastWins
}

/// <summary>
/// Settings for loading a <see cref="YamlDocument"/>. The defaults read YAML 1.2 as real files use it (see
/// <see cref="Strict"/>) and bound resource use for untrusted input. An instance can be shared by loads on several threads as long as it is not modified while they run.
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
        Strict = Strict,
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

    /// <summary>
    /// Gets or sets whether loading follows the YAML 1.2.2 specification exactly. Off by default, loading accepts what
    /// real files rely on and libyaml and SharpYaml accept: flow-collection lines (such as a closing <c>]</c>
    /// under its key) and quoted-scalar continuation lines indented no deeper than the enclosing block collection, and
    /// <c>&lt;&lt;</c> merge keys under the YAML 1.2 core schema as under YAML 1.1, so repeated <c>&lt;&lt;</c> keys
    /// (common in docker-compose and GitLab CI files) are one merge list rather than duplicate keys, and duplicate keys
    /// unless <see cref="DuplicateKeys"/> says otherwise (the last one wins). Each accepted
    /// indentation and each document part that uses merge keys without a <c>%YAML 1.1</c> directive is reported in
    /// <see cref="YamlDocument.Diagnostics"/>. When on, both indentation forms are syntax errors, <c>&lt;&lt;</c> is an
    /// ordinary key unless the YAML 1.1 schema applies, and a duplicate key fails the load unless <see cref="DuplicateKeys"/>
    /// is set. Saving never writes the accepted indentation for new or edited text.
    /// </summary>
    public bool Strict {
        get => _reader.Strict;
        set => _reader.Strict = value;
    }

    /// <summary>
    /// Gets or sets how a mapping with equal keys loads. When <see langword="null"/> (the default), the last key wins and
    /// each repeat is a <see cref="OfficeDiagnosticCodes.YamlDuplicateKey"/> warning in <see cref="YamlDocument.Diagnostics"/>,
    /// or, with <see cref="Strict"/>, the load fails with <see cref="YamlErrorCodes.DuplicateKey"/> as the specification requires.
    /// Either way the document keeps every entry, so saving reproduces the file.
    /// </summary>
    public YamlDuplicateKeyPolicy? DuplicateKeys { get; set; }

    /// <summary>The policy the load applies: <see cref="DuplicateKeys"/>, or the default that <see cref="Strict"/> selects.</summary>
    internal YamlDuplicateKeyPolicy EffectiveDuplicateKeys => DuplicateKeys ?? (Strict ? YamlDuplicateKeyPolicy.Error : YamlDuplicateKeyPolicy.LastWins);

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
