namespace OfficeIMO.Yaml;

/// <summary>
/// The base of generated serialization metadata. Declare a partial class that derives from it and lists its root types
/// with <see cref="YamlSerializableAttribute"/>; the source generator fills in metadata for those types and every type
/// they reach, so serialization needs no reflection and works with trimming and NativeAOT.
/// </summary>
/// <example>
/// <code>
/// [YamlSerializable(typeof(AppConfig))]
/// internal partial class AppYamlContext : YamlSerializerContext { }
///
/// string yaml = YamlSerializer.Serialize(config, AppYamlContext.Default.AppConfig);
/// AppConfig? copy = YamlSerializer.Deserialize(yaml, AppYamlContext.Default.AppConfig);
/// </code>
/// </example>
public abstract class YamlSerializerContext : IYamlTypeInfoResolver {
    private YamlSerializerOptions? _options;

    /// <summary>Creates the context and binds it to <paramref name="options"/>, or to new default options when null.</summary>
    /// <exception cref="InvalidOperationException">The options are read-only or already use another resolver.</exception>
    protected YamlSerializerContext(YamlSerializerOptions? options) {
        if (options is not null) {
            if (options.TypeInfoResolver is not null && !ReferenceEquals(options.TypeInfoResolver, this)) {
                throw new InvalidOperationException("The options already use another metadata resolver.");
            }
            options.TypeInfoResolver = this;
            _options = options;
        }
    }

    /// <summary>Gets the options the context's metadata is bound to.</summary>
    public YamlSerializerOptions Options {
        get {
            YamlSerializerOptions? options = _options;
            if (options is null) {
                Interlocked.CompareExchange(ref _options, new YamlSerializerOptions { TypeInfoResolver = this }, null);
                options = _options!;
            }
            return options;
        }
    }

    /// <summary>Returns the metadata for <paramref name="type"/> bound to <see cref="Options"/>, or null when the context does not include it.</summary>
    public YamlTypeInfo? GetTypeInfo(Type type) {
        if (type is null) throw new ArgumentNullException(nameof(type));
        try {
            return Options.GetTypeInfo(type);
        } catch (NotSupportedException) {
            return null;
        }
    }

    /// <summary>Creates metadata for <paramref name="type"/> bound to <paramref name="options"/>; generated contexts implement it.</summary>
    protected abstract YamlTypeInfo? CreateTypeInfo(Type type, YamlSerializerOptions options);

    YamlTypeInfo? IYamlTypeInfoResolver.GetTypeInfo(Type type, YamlSerializerOptions options) => CreateTypeInfo(type, options);
}
