using System.Collections.Concurrent;

namespace OfficeIMO.Yaml;

/// <summary>Supplies <see cref="YamlTypeInfo"/> metadata for the types it knows. <see cref="YamlSerializerContext"/> implements it.</summary>
public interface IYamlTypeInfoResolver {
    /// <summary>Returns metadata for <paramref name="type"/> bound to <paramref name="options"/>, or null for a type the resolver does not know.</summary>
    YamlTypeInfo? GetTypeInfo(Type type, YamlSerializerOptions options);
}

/// <summary>
/// Settings for <see cref="YamlSerializer"/>. An instance caches the metadata it creates, so reuse one per configuration.
/// It becomes read-only when first used; changing it afterwards throws.
/// </summary>
public sealed class YamlSerializerOptions {
    private readonly ConcurrentDictionary<Type, YamlTypeInfo> _cache = new ConcurrentDictionary<Type, YamlTypeInfo>();
    private YamlNamingPolicy? _namingPolicy;
    private YamlIgnoreCondition _defaultIgnoreCondition;
    private YamlReferenceHandling _referenceHandling;
    private YamlReaderOptions? _readerOptions;
    private YamlWriterOptions? _writerOptions;
    private IYamlTypeInfoResolver? _resolver;
    private volatile IYamlTypeInfoResolver? _reflectionResolver;
    private volatile bool _readOnly;

    /// <summary>Creates options with the defaults.</summary>
    public YamlSerializerOptions() {
    }

    /// <summary>Creates a writable copy of <paramref name="options"/>, including its resolver.</summary>
    public YamlSerializerOptions(YamlSerializerOptions options) {
        if (options is null) throw new ArgumentNullException(nameof(options));
        _namingPolicy = options._namingPolicy;
        _defaultIgnoreCondition = options._defaultIgnoreCondition;
        _referenceHandling = options._referenceHandling;
        _readerOptions = options._readerOptions;
        _writerOptions = options._writerOptions;
        _resolver = options._resolver;
    }

    /// <summary>
    /// Gets the shared default options, which are read-only. Serializer overloads without options use them.
    /// </summary>
    public static YamlSerializerOptions Default { get; } = new YamlSerializerOptions { _readOnly = true };

    /// <summary>
    /// Gets or sets the policy that turns property names into mapping keys, or null (the default) to use the names as
    /// declared. <see cref="YamlPropertyNameAttribute"/> wins over it.
    /// </summary>
    public YamlNamingPolicy? PropertyNamingPolicy {
        get => _namingPolicy;
        set {
            CheckWritable();
            _namingPolicy = value;
        }
    }

    /// <summary>
    /// Gets or sets when properties are left out of written YAML: <see cref="YamlIgnoreCondition.Never"/> (the default),
    /// <see cref="YamlIgnoreCondition.WhenWritingNull"/> or <see cref="YamlIgnoreCondition.WhenWritingDefault"/>.
    /// </summary>
    public YamlIgnoreCondition DefaultIgnoreCondition {
        get => _defaultIgnoreCondition;
        set {
            CheckWritable();
            if (value == YamlIgnoreCondition.Always) throw new ArgumentException("Always applies to single properties through YamlIgnoreAttribute.", nameof(value));
            _defaultIgnoreCondition = value;
        }
    }

    /// <summary>Gets or sets how shared references and cycles are written and how aliases are read. Defaults to <see cref="YamlReferenceHandling.None"/>.</summary>
    public YamlReferenceHandling ReferenceHandling {
        get => _referenceHandling;
        set {
            CheckWritable();
            _referenceHandling = value;
        }
    }

    /// <summary>
    /// Gets or sets the reader settings for overloads that read text, streams or text readers: the schema, strictness and
    /// the limits, including <see cref="YamlReaderOptions.MaxAliasExpansion"/> for aliases. Null uses the defaults; use
    /// <see cref="YamlReaderOptions.UntrustedDefaults"/> for input from outside. Overloads that take a
    /// <see cref="YamlReader"/> use that reader's settings.
    /// </summary>
    public YamlReaderOptions? ReaderOptions {
        get => _readerOptions;
        set {
            CheckWritable();
            _readerOptions = value;
        }
    }

    /// <summary>Gets or sets the layout of written YAML, or null for the <see cref="YamlWriter"/> defaults.</summary>
    public YamlWriterOptions? WriterOptions {
        get => _writerOptions;
        set {
            CheckWritable();
            _writerOptions = value;
        }
    }

    /// <summary>
    /// Gets or sets where metadata comes from: a generated <see cref="YamlSerializerContext"/>, or null. With null,
    /// <see cref="YamlSerializer"/> overloads that allow reflection create metadata by reflection.
    /// </summary>
    public IYamlTypeInfoResolver? TypeInfoResolver {
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
    public YamlTypeInfo GetTypeInfo(Type type) {
        if (type is null) throw new ArgumentNullException(nameof(type));
        if (_cache.TryGetValue(type, out YamlTypeInfo? info)) return info;
        _readOnly = true;
        info = YamlMetadataServices.TryCreateBuiltIn(type, this);
        if (info is null) {
            IYamlTypeInfoResolver? resolver = _resolver ?? _reflectionResolver;
            info = resolver?.GetTypeInfo(type, this);
            if (info is null) throw NoMetadata(type, resolver);
            if (!ReferenceEquals(info.Options, this) || info.Type != type) {
                throw new InvalidOperationException("The resolver returned metadata for a different type or options instance than requested for " + YamlSerializationErrors.TypeName(type) + ".");
            }
        }
        return _cache.GetOrAdd(type, info);
    }

    /// <summary>Gets typed metadata, for the serializer entry points.</summary>
    internal YamlTypeInfo<T> GetTypeInfo<T>() => (YamlTypeInfo<T>)GetTypeInfo(typeof(T));

    /// <summary>Gets whether a reflection fallback is in place, for an error message that names the missing source.</summary>
    internal bool UsesReflection => _resolver is null && _reflectionResolver is not null;

    /// <summary>Lets overloads that allow reflection use it when no resolver is set; reflection never replaces a set resolver.</summary>
    internal void UseReflectionFallback(IYamlTypeInfoResolver resolver) {
        if (_resolver is null && _reflectionResolver is null) _reflectionResolver = resolver;
    }

    internal YamlReaderOptions EffectiveReaderOptions => _readerOptions ?? YamlReaderOptions.Default;

    internal YamlWriterOptions EffectiveWriterOptions => _writerOptions ?? YamlWriterOptions.Default;

    private static NotSupportedException NoMetadata(Type type, IYamlTypeInfoResolver? resolver) {
        string name = YamlSerializationErrors.TypeName(type);
        if (resolver is YamlSerializerContext context) {
            return new NotSupportedException("The type " + name + " is not part of " + context.GetType().Name
                + ". Add [YamlSerializable(typeof(" + name + "))] to the context.");
        }
        if (resolver is null) {
            return new NotSupportedException("No metadata is available for " + name + ". Pass generated metadata from a YamlSerializerContext ("
                + "[YamlSerializable(typeof(" + name + "))]), or set YamlSerializerOptions.TypeInfoResolver to the context.");
        }
        return new NotSupportedException("The metadata resolver does not support " + name + ".");
    }

    private void CheckWritable() {
        if (_readOnly) throw new InvalidOperationException("The options are read-only because they are in use; create a copy with new YamlSerializerOptions(options) to change them.");
    }
}
