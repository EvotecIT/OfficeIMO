using System.ComponentModel;
using System.Numerics;

namespace OfficeIMO.Yaml;

/// <summary>
/// Creates <see cref="YamlTypeInfo{T}"/> instances. The source generator calls these methods from the code it emits for a
/// <see cref="YamlSerializerContext"/>, and reflection-based metadata uses the same methods, so both produce the same YAML.
/// </summary>
[EditorBrowsable(EditorBrowsableState.Never)]
public static class YamlMetadataServices {
    /// <summary>Creates metadata for an object written as a mapping of its properties.</summary>
    public static YamlTypeInfo<T> CreateObjectInfo<T>(YamlSerializerOptions options, YamlObjectInfoValues<T> values) {
        if (options is null) throw new ArgumentNullException(nameof(options));
        if (values is null) throw new ArgumentNullException(nameof(values));
        var objectInfo = new YamlTypeInfo<T>(options, new YamlObjectConverter<T>(values));
        YamlDerivedTypeAttribute[]? derived = values.DerivedTypes;
        if (derived is null || derived.Length == 0) return objectInfo;
        return new YamlTypeInfo<T>(options, new YamlPolymorphicConverter<T>(objectInfo, derived, values.DiscriminatorPropertyName), objectInfo);
    }

    /// <summary>Creates metadata for one property of an object.</summary>
    public static YamlPropertyInfo CreatePropertyInfo<T>(YamlSerializerOptions options, YamlPropertyInfoValues<T> values) {
        if (options is null) throw new ArgumentNullException(nameof(options));
        if (values is null) throw new ArgumentNullException(nameof(values));
        if (values.IgnoreCondition == YamlIgnoreCondition.Always) throw new ArgumentException("An ignored property has no metadata.", nameof(values));
        return new YamlPropertyInfo<T>(values);
    }

    /// <summary>Creates metadata for an enum written by member name.</summary>
    public static YamlTypeInfo<T> CreateEnumInfo<T>(YamlSerializerOptions options) where T : struct, Enum {
        if (options is null) throw new ArgumentNullException(nameof(options));
        return new YamlTypeInfo<T>(options, new YamlEnumConverter<T>());
    }

    /// <summary>Creates metadata for a nullable value type.</summary>
    public static YamlTypeInfo<T?> CreateNullableInfo<T>(YamlSerializerOptions options) where T : struct {
        if (options is null) throw new ArgumentNullException(nameof(options));
        return new YamlTypeInfo<T?>(options, new YamlNullableConverter<T>());
    }

    /// <summary>
    /// Creates metadata for a collection written as a sequence. Items are read into a list, and
    /// <paramref name="finish"/> turns the list into the collection.
    /// </summary>
    public static YamlTypeInfo<TCollection> CreateCollectionInfo<TCollection, TElement>(YamlSerializerOptions options, Func<List<TElement>, TCollection> finish)
        where TCollection : IEnumerable<TElement> {
        if (options is null) throw new ArgumentNullException(nameof(options));
        if (finish is null) throw new ArgumentNullException(nameof(finish));
        return new YamlTypeInfo<TCollection>(options, new YamlCollectionConverter<TCollection, TElement>(finish));
    }

    /// <summary>
    /// Creates metadata for a dictionary written as a mapping. Entries are read into a dictionary, and
    /// <paramref name="finish"/> turns it into the target type.
    /// </summary>
    public static YamlTypeInfo<TDictionary> CreateDictionaryInfo<TDictionary, TKey, TValue>(YamlSerializerOptions options, Func<Dictionary<TKey, TValue>, TDictionary> finish)
        where TDictionary : IEnumerable<KeyValuePair<TKey, TValue>>
        where TKey : notnull {
        if (options is null) throw new ArgumentNullException(nameof(options));
        if (finish is null) throw new ArgumentNullException(nameof(finish));
        return new YamlTypeInfo<TDictionary>(options, new YamlDictionaryConverter<TDictionary, TKey, TValue>(finish));
    }

    /// <summary>Returns built-in metadata for scalars, <see cref="object"/> and the untyped collections, or null.</summary>
    internal static YamlTypeInfo? TryCreateBuiltIn(Type type, YamlSerializerOptions options) {
        if (type == typeof(string)) return new YamlTypeInfo<string>(options, YamlStringConverter.Instance);
        if (type == typeof(object)) return new YamlTypeInfo<object>(options, new YamlUntypedConverter());
        if (type == typeof(Dictionary<string, object?>)) return CreateDictionaryInfo<Dictionary<string, object?>, string, object?>(options, static d => d);
        if (type == typeof(List<object?>)) return CreateCollectionInfo<List<object?>, object?>(options, static l => l);
        if (type == typeof(object[])) return CreateCollectionInfo<object?[], object?>(options, static l => l.ToArray());
        if (type == typeof(byte[])) return new YamlTypeInfo<byte[]>(options, YamlByteArrayConverter.Instance);
        if (type == typeof(Uri)) return new YamlTypeInfo<Uri>(options, YamlUriConverter.Instance);
        if (!type.IsValueType) return null;
        Type? underlying = Nullable.GetUnderlyingType(type);
        if (underlying is not null) {
            return underlying == typeof(bool) ? CreateNullableInfo<bool>(options)
                : underlying == typeof(int) ? CreateNullableInfo<int>(options)
                : underlying == typeof(long) ? CreateNullableInfo<long>(options)
                : underlying == typeof(double) ? CreateNullableInfo<double>(options)
                : underlying == typeof(decimal) ? CreateNullableInfo<decimal>(options)
                : underlying == typeof(float) ? CreateNullableInfo<float>(options)
                : underlying == typeof(byte) ? CreateNullableInfo<byte>(options)
                : underlying == typeof(sbyte) ? CreateNullableInfo<sbyte>(options)
                : underlying == typeof(short) ? CreateNullableInfo<short>(options)
                : underlying == typeof(ushort) ? CreateNullableInfo<ushort>(options)
                : underlying == typeof(uint) ? CreateNullableInfo<uint>(options)
                : underlying == typeof(ulong) ? CreateNullableInfo<ulong>(options)
                : underlying == typeof(char) ? CreateNullableInfo<char>(options)
                : underlying == typeof(BigInteger) ? CreateNullableInfo<BigInteger>(options)
                : underlying == typeof(DateTime) ? CreateNullableInfo<DateTime>(options)
                : underlying == typeof(DateTimeOffset) ? CreateNullableInfo<DateTimeOffset>(options)
                : underlying == typeof(TimeSpan) ? CreateNullableInfo<TimeSpan>(options)
                : underlying == typeof(Guid) ? CreateNullableInfo<Guid>(options)
#if NET6_0_OR_GREATER
                : underlying == typeof(DateOnly) ? CreateNullableInfo<DateOnly>(options)
                : underlying == typeof(TimeOnly) ? CreateNullableInfo<TimeOnly>(options)
#endif
#if NET6_0_OR_GREATER
                : (YamlTypeInfo?)null;
#else
                : YamlFormattedValueConverter<DateTime>.TryCreate(underlying, options) is null ? null
                : (YamlTypeInfo)Activator.CreateInstance(typeof(YamlTypeInfo<>).MakeGenericType(type), System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic,
                    null, new object?[] { options, Activator.CreateInstance(typeof(YamlNullableConverter<>).MakeGenericType(underlying)), null }, null)!;
#endif
        }
        if (type == typeof(bool)) return new YamlTypeInfo<bool>(options, YamlBooleanConverter.Instance);
        if (type == typeof(int)) return new YamlTypeInfo<int>(options, YamlInt32Converter.Instance);
        if (type == typeof(long)) return new YamlTypeInfo<long>(options, YamlInt64Converter.Instance);
        if (type == typeof(double)) return new YamlTypeInfo<double>(options, YamlDoubleConverter.Instance);
        if (type == typeof(decimal)) return new YamlTypeInfo<decimal>(options, YamlDecimalConverter.Instance);
        if (type == typeof(float)) return new YamlTypeInfo<float>(options, YamlSingleConverter.Instance);
        if (type == typeof(byte)) return new YamlTypeInfo<byte>(options, YamlByteConverter.Instance);
        if (type == typeof(sbyte)) return new YamlTypeInfo<sbyte>(options, YamlSByteConverter.Instance);
        if (type == typeof(short)) return new YamlTypeInfo<short>(options, YamlInt16Converter.Instance);
        if (type == typeof(ushort)) return new YamlTypeInfo<ushort>(options, YamlUInt16Converter.Instance);
        if (type == typeof(uint)) return new YamlTypeInfo<uint>(options, YamlUInt32Converter.Instance);
        if (type == typeof(ulong)) return new YamlTypeInfo<ulong>(options, YamlUInt64Converter.Instance);
        if (type == typeof(char)) return new YamlTypeInfo<char>(options, YamlCharConverter.Instance);
        if (type == typeof(BigInteger)) return new YamlTypeInfo<BigInteger>(options, YamlBigIntegerConverter.Instance);
        if (type == typeof(DateTime)) return new YamlTypeInfo<DateTime>(options, YamlDateTimeConverter.Instance);
        if (type == typeof(DateTimeOffset)) return new YamlTypeInfo<DateTimeOffset>(options, YamlDateTimeOffsetConverter.Instance);
        if (type == typeof(TimeSpan)) return new YamlTypeInfo<TimeSpan>(options, YamlTimeSpanConverter.Instance);
        if (type == typeof(Guid)) return new YamlTypeInfo<Guid>(options, YamlGuidConverter.Instance);
#if NET6_0_OR_GREATER
        if (type == typeof(DateOnly)) return new YamlTypeInfo<DateOnly>(options, YamlDateOnlyConverter.Instance);
        if (type == typeof(TimeOnly)) return new YamlTypeInfo<TimeOnly>(options, YamlTimeOnlyConverter.Instance);
        return null;
#else
        return YamlFormattedValueConverter<DateTime>.TryCreate(type, options);
#endif
    }
}

/// <summary>The constructor and members of an object type, for <see cref="YamlMetadataServices.CreateObjectInfo{T}"/>.</summary>
/// <typeparam name="T">The object type.</typeparam>
[EditorBrowsable(EditorBrowsableState.Never)]
public sealed class YamlObjectInfoValues<T> {
    /// <summary>Gets or sets the function that creates an instance with the parameterless constructor.</summary>
    public Func<T>? ObjectCreator { get; set; }

    /// <summary>
    /// Gets or sets the function that creates an instance from constructor arguments, in parameter order. When set it is
    /// used instead of <see cref="ObjectCreator"/>.
    /// </summary>
    public Func<object?[], T>? ObjectWithParameterizedConstructorCreator { get; set; }

    /// <summary>Gets or sets the function that creates the property metadata, in writing order.</summary>
    public Func<YamlSerializerOptions, YamlPropertyInfo[]>? PropertyMetadataInitializer { get; set; }

    /// <summary>Gets or sets the function that describes the parameters of the constructor <see cref="ObjectWithParameterizedConstructorCreator"/> calls.</summary>
    public Func<YamlParameterInfoValues[]>? ConstructorParameterMetadataInitializer { get; set; }

    /// <summary>Gets or sets the registered derived types, which make the type polymorphic.</summary>
    public YamlDerivedTypeAttribute[]? DerivedTypes { get; set; }

    /// <summary>Gets or sets the mapping key that holds the discriminator, or null to use tags.</summary>
    public string? DiscriminatorPropertyName { get; set; }
}

/// <summary>One property of an object type, for <see cref="YamlMetadataServices.CreatePropertyInfo{T}"/>.</summary>
/// <typeparam name="T">The property type.</typeparam>
[EditorBrowsable(EditorBrowsableState.Never)]
public sealed class YamlPropertyInfoValues<T> {
    /// <summary>Gets or sets the .NET property name.</summary>
    public string Name { get; set; } = string.Empty;

    /// <summary>Gets or sets the name from <see cref="YamlPropertyNameAttribute"/>, or null.</summary>
    public string? YamlName { get; set; }

    /// <summary>Gets or sets the getter, or null for a property that cannot be read.</summary>
    public Func<object, T>? Getter { get; set; }

    /// <summary>Gets or sets the setter, or null for a property that is not set after construction.</summary>
    public Action<object, T>? Setter { get; set; }

    /// <summary>Gets or sets whether a mapping must contain the property.</summary>
    public bool IsRequired { get; set; }

    /// <summary>Gets or sets the condition from <see cref="YamlIgnoreAttribute"/>, or null for the options' default.</summary>
    public YamlIgnoreCondition? IgnoreCondition { get; set; }
}

/// <summary>One parameter of the constructor used to create an object.</summary>
[EditorBrowsable(EditorBrowsableState.Never)]
public sealed class YamlParameterInfoValues {
    /// <summary>Gets or sets the parameter name, which binds to the property of the same name ignoring case.</summary>
    public string Name { get; set; } = string.Empty;

    /// <summary>Gets or sets the parameter type, which equals the bound property's type.</summary>
    public Type ParameterType { get; set; } = typeof(object);

    /// <summary>Gets or sets the zero-based position.</summary>
    public int Position { get; set; }

    /// <summary>Gets or sets the value passed when the mapping has no key for the parameter: its default value, or the type's default.</summary>
    public object? DefaultValue { get; set; }
}
