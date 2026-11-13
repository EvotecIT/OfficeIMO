using System.Collections.Concurrent;

namespace OfficeIMO.Toml;

/// <summary>Supplies <see cref="TomlTypeInfo"/> metadata for the types it knows. <see cref="TomlSerializerContext"/> implements it.</summary>
public interface ITomlTypeInfoResolver {
    /// <summary>Returns metadata for <paramref name="type"/> bound to <paramref name="options"/>, or null for a type the resolver does not know.</summary>
    TomlTypeInfo? GetTypeInfo(Type type, TomlSerializerOptions options);
}

/// <summary>
/// Settings for <see cref="TomlSerializer"/>. An instance caches the metadata it creates, so reuse one per configuration.
/// It becomes read-only when first used; changing it afterwards throws. The naming policy and ignore condition are the
/// types <c>YamlSerializerOptions</c> uses too.
/// </summary>
public sealed class TomlSerializerOptions {
    private readonly ConcurrentDictionary<Type, TomlTypeInfo> _cache = new ConcurrentDictionary<Type, TomlTypeInfo>();
    private OfficeNamingPolicy? _namingPolicy;
    private OfficeIgnoreCondition _defaultIgnoreCondition;
    private TomlNullHandling _nullHandling;
    private TomlLoadOptions? _loadOptions;
    private TomlWriterOptions? _writerOptions;
    private ITomlTypeInfoResolver? _resolver;
    private volatile ITomlTypeInfoResolver? _reflectionResolver;
    private volatile bool _readOnly;

    /// <summary>Creates options with the defaults.</summary>
    public TomlSerializerOptions() {
    }

    /// <summary>Creates a writable copy of <paramref name="options"/>, including its resolver.</summary>
    public TomlSerializerOptions(TomlSerializerOptions options) {
        if (options is null) throw new ArgumentNullException(nameof(options));
        _namingPolicy = options._namingPolicy;
        _defaultIgnoreCondition = options._defaultIgnoreCondition;
        _nullHandling = options._nullHandling;
        _loadOptions = options._loadOptions;
        _writerOptions = options._writerOptions;
        _resolver = options._resolver;
    }

    /// <summary>Gets the shared default options, which are read-only. Serializer overloads without options use them.</summary>
    public static TomlSerializerOptions Default { get; } = new TomlSerializerOptions { _readOnly = true };

    /// <summary>
    /// Gets or sets the policy that turns property names into keys, or null (the default) to use the names as declared.
    /// <see cref="OfficePropertyNameAttribute"/> wins over it. <see cref="OfficeNamingPolicy.SnakeCaseLower"/> matches
    /// <c>pyproject.toml</c> and <see cref="OfficeNamingPolicy.KebabCaseLower"/> matches <c>Cargo.toml</c>.
    /// </summary>
    public OfficeNamingPolicy? PropertyNamingPolicy {
        get => _namingPolicy;
        set {
            CheckWritable();
            _namingPolicy = value;
        }
    }

    /// <summary>
    /// Gets or sets when properties are left out of written TOML: <see cref="OfficeIgnoreCondition.Never"/> (the default),
    /// <see cref="OfficeIgnoreCondition.WhenWritingNull"/> or <see cref="OfficeIgnoreCondition.WhenWritingDefault"/>. TOML
    /// has no null, so a null value is never written; <see cref="NullHandling"/> says what happens instead.
    /// </summary>
    public OfficeIgnoreCondition DefaultIgnoreCondition {
        get => _defaultIgnoreCondition;
        set {
            CheckWritable();
            if (value == OfficeIgnoreCondition.Always) throw new ArgumentException("Always applies to single properties through OfficeIgnoreAttribute.", nameof(value));
            _defaultIgnoreCondition = value;
        }
    }

    /// <summary>
    /// Gets or sets what writing does with a null property or dictionary value, which TOML cannot hold:
    /// <see cref="TomlNullHandling.Omit"/> (the default) leaves the key out, and <see cref="TomlNullHandling.Error"/>
    /// fails. A property that a <see cref="OfficeIgnoreCondition"/> skips when null is left out either way. A null array
    /// item always fails, because leaving it out would shift the items after it.
    /// </summary>
    public TomlNullHandling NullHandling {
        get => _nullHandling;
        set {
            CheckWritable();
            _nullHandling = value;
        }
    }

    /// <summary>
    /// Gets or sets the limits for overloads that read text, streams or text readers, or null for the defaults. Use
    /// <see cref="TomlLoadOptions.UntrustedDefaults"/> for input from outside.
    /// </summary>
    public TomlLoadOptions? LoadOptions {
        get => _loadOptions;
        set {
            CheckWritable();
            _loadOptions = value;
        }
    }

    /// <summary>Gets or sets the line break of written TOML, or null for the <see cref="TomlWriter"/> default (<c>\n</c>).</summary>
    public TomlWriterOptions? WriterOptions {
        get => _writerOptions;
        set {
            CheckWritable();
            _writerOptions = value;
        }
    }

    /// <summary>
    /// Gets or sets where metadata comes from: a generated <see cref="TomlSerializerContext"/>, or null. With null,
    /// <see cref="TomlSerializer"/> overloads that allow reflection create metadata by reflection.
    /// </summary>
    public ITomlTypeInfoResolver? TypeInfoResolver {
        get => _resolver;
        set {
            CheckWritable();
            _resolver = value;
        }
    }

    /// <summary>Gets whether the options are read-only, which they become on first use.</summary>
    public bool IsReadOnly => _readOnly;

    /// <summary>
    /// Returns the metadata for <paramref name="type"/> from the built-in types or <see cref="TypeInfoResolver"/>, creating
    /// it once. It does not fall back to reflection.
    /// </summary>
    /// <exception cref="NotSupportedException">No metadata source knows the type.</exception>
    public TomlTypeInfo GetTypeInfo(Type type) {
        if (type is null) throw new ArgumentNullException(nameof(type));
        if (_cache.TryGetValue(type, out TomlTypeInfo? info)) return info;
        _readOnly = true;
        info = TomlMetadataServices.TryCreateBuiltIn(type, this);
        if (info is null) {
            ITomlTypeInfoResolver? resolver = _resolver ?? _reflectionResolver;
            info = resolver?.GetTypeInfo(type, this);
            if (info is null) throw NoMetadata(type, resolver);
            if (!ReferenceEquals(info.Options, this) || info.Type != type) {
                throw new InvalidOperationException("The resolver returned metadata for a different type or options instance than requested for " + TomlSerializationErrors.TypeName(type) + ".");
            }
        }
        return _cache.GetOrAdd(type, info);
    }

    /// <summary>Lets overloads that allow reflection use it when no resolver is set; reflection never replaces a set resolver.</summary>
    internal void UseReflectionFallback(ITomlTypeInfoResolver resolver) {
        if (_resolver is null && _reflectionResolver is null) _reflectionResolver = resolver;
    }

    internal TomlLoadOptions? EffectiveLoadOptions => _loadOptions;

    internal string NewLine => _writerOptions?.NewLine ?? "\n";

    private static NotSupportedException NoMetadata(Type type, ITomlTypeInfoResolver? resolver) {
        string name = TomlSerializationErrors.TypeName(type);
        if (resolver is TomlSerializerContext context) {
            return new NotSupportedException("The type " + name + " is not part of " + context.GetType().Name
                + ". Add [TomlSerializable(typeof(" + name + "))] to the context.");
        }
        if (resolver is null) {
            return new NotSupportedException("No metadata is available for " + name + ". Pass generated metadata from a TomlSerializerContext ("
                + "[TomlSerializable(typeof(" + name + "))]), or set TomlSerializerOptions.TypeInfoResolver to the context.");
        }
        return new NotSupportedException("The metadata resolver does not support " + name + ".");
    }

    private void CheckWritable() {
        if (_readOnly) throw new InvalidOperationException("The options are read-only because they are in use; create a copy with new TomlSerializerOptions(options) to change them.");
    }
}
