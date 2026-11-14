using System.ComponentModel;
using System.Numerics;

namespace OfficeIMO.Toml;

/// <summary>
/// Creates <see cref="TomlTypeInfo{T}"/> instances. The source generator calls these methods from the code it emits for a
/// <see cref="TomlSerializerContext"/>, and reflection-based metadata uses the same methods, so both produce the same TOML.
/// </summary>
[EditorBrowsable(EditorBrowsableState.Never)]
public static class TomlMetadataServices {
    /// <summary>Creates metadata for an object written as a table of its properties.</summary>
    public static TomlTypeInfo<T> CreateObjectInfo<T>(TomlSerializerOptions options, TomlObjectInfoValues<T> values) {
        if (options is null) throw new ArgumentNullException(nameof(options));
        if (values is null) throw new ArgumentNullException(nameof(values));
        var objectInfo = new TomlTypeInfo<T>(options, new TomlObjectConverter<T>(values));
        OfficeDerivedTypeAttribute[]? derived = values.DerivedTypes;
        if (derived is null || derived.Length == 0) return objectInfo;
        return new TomlTypeInfo<T>(options, new TomlPolymorphicConverter<T>(objectInfo, derived, values.DiscriminatorPropertyName), objectInfo);
    }

    /// <summary>Creates metadata for one property of an object.</summary>
    public static TomlPropertyInfo CreatePropertyInfo<T>(TomlSerializerOptions options, TomlPropertyInfoValues<T> values) {
        if (options is null) throw new ArgumentNullException(nameof(options));
        if (values is null) throw new ArgumentNullException(nameof(values));
        if (values.IgnoreCondition == OfficeIgnoreCondition.Always) throw new ArgumentException("An ignored property has no metadata.", nameof(values));
        return new TomlPropertyInfo<T>(values);
    }

    /// <summary>Creates metadata for an enum written by member name.</summary>
    public static TomlTypeInfo<T> CreateEnumInfo<T>(TomlSerializerOptions options) where T : struct, Enum {
        if (options is null) throw new ArgumentNullException(nameof(options));
        return new TomlTypeInfo<T>(options, new TomlEnumConverter<T>());
    }

    /// <summary>Creates metadata for a nullable value type.</summary>
    public static TomlTypeInfo<T?> CreateNullableInfo<T>(TomlSerializerOptions options) where T : struct {
        if (options is null) throw new ArgumentNullException(nameof(options));
        return new TomlTypeInfo<T?>(options, new TomlNullableConverter<T>());
    }

    /// <summary>
    /// Creates metadata for a collection written as an array or an array of tables. Items are read into a list, and
    /// <paramref name="finish"/> turns the list into the collection.
    /// </summary>
    public static TomlTypeInfo<TCollection> CreateCollectionInfo<TCollection, TElement>(TomlSerializerOptions options, Func<List<TElement>, TCollection> finish)
        where TCollection : IEnumerable<TElement> {
        if (options is null) throw new ArgumentNullException(nameof(options));
        if (finish is null) throw new ArgumentNullException(nameof(finish));
        return new TomlTypeInfo<TCollection>(options, new TomlCollectionConverter<TCollection, TElement>(finish));
    }

    /// <summary>
    /// Creates metadata for a dictionary written as a table. Entries are read into a dictionary, and
    /// <paramref name="finish"/> turns it into the target type.
    /// </summary>
    public static TomlTypeInfo<TDictionary> CreateDictionaryInfo<TDictionary, TKey, TValue>(TomlSerializerOptions options, Func<Dictionary<TKey, TValue>, TDictionary> finish)
        where TDictionary : IEnumerable<KeyValuePair<TKey, TValue>>
        where TKey : notnull {
        if (options is null) throw new ArgumentNullException(nameof(options));
        if (finish is null) throw new ArgumentNullException(nameof(finish));
        return new TomlTypeInfo<TDictionary>(options, new TomlDictionaryConverter<TDictionary, TKey, TValue>(finish));
    }

    private static readonly TomlIntegerConverter<int> Int32 = new TomlIntegerConverter<int>(int.MinValue, int.MaxValue, static n => (int)n, static v => v);
    private static readonly TomlIntegerConverter<long> Int64 = new TomlIntegerConverter<long>(long.MinValue, long.MaxValue, static n => n, static v => v);
    private static readonly TomlIntegerConverter<short> Int16 = new TomlIntegerConverter<short>(short.MinValue, short.MaxValue, static n => (short)n, static v => v);
    private static readonly TomlIntegerConverter<sbyte> SByte = new TomlIntegerConverter<sbyte>(sbyte.MinValue, sbyte.MaxValue, static n => (sbyte)n, static v => v);
    private static readonly TomlIntegerConverter<byte> Byte = new TomlIntegerConverter<byte>(byte.MinValue, byte.MaxValue, static n => (byte)n, static v => v);
    private static readonly TomlIntegerConverter<ushort> UInt16 = new TomlIntegerConverter<ushort>(ushort.MinValue, ushort.MaxValue, static n => (ushort)n, static v => v);
    private static readonly TomlIntegerConverter<uint> UInt32 = new TomlIntegerConverter<uint>(uint.MinValue, uint.MaxValue, static n => (uint)n, static v => v);

    /// <summary>Returns built-in metadata for scalars, <see cref="object"/> and the untyped collections, or null.</summary>
    internal static TomlTypeInfo? TryCreateBuiltIn(Type type, TomlSerializerOptions options) {
        if (type == typeof(string)) return new TomlTypeInfo<string>(options, TomlStringConverter.Instance);
        if (type == typeof(object)) return new TomlTypeInfo<object>(options, new TomlUntypedConverter());
        if (type == typeof(Dictionary<string, object?>)) return CreateDictionaryInfo<Dictionary<string, object?>, string, object?>(options, static d => d);
        if (type == typeof(List<object?>)) return CreateCollectionInfo<List<object?>, object?>(options, static l => l);
        if (type == typeof(object[])) return CreateCollectionInfo<object?[], object?>(options, static l => l.ToArray());
        if (type == typeof(byte[])) return new TomlTypeInfo<byte[]>(options, TomlTextConverters.Base64);
        if (type == typeof(Uri)) return new TomlTypeInfo<Uri>(options, TomlTextConverters.UriText);
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
                : underlying == typeof(TomlDateTime) ? CreateNullableInfo<TomlDateTime>(options)
#if NET6_0_OR_GREATER
                : underlying == typeof(DateOnly) ? CreateNullableInfo<DateOnly>(options)
                : underlying == typeof(TimeOnly) ? CreateNullableInfo<TimeOnly>(options)
#endif
                : (TomlTypeInfo?)null;
        }
        if (type == typeof(bool)) return new TomlTypeInfo<bool>(options, TomlBooleanConverter.Instance);
        if (type == typeof(int)) return new TomlTypeInfo<int>(options, Int32);
        if (type == typeof(long)) return new TomlTypeInfo<long>(options, Int64);
        if (type == typeof(double)) return new TomlTypeInfo<double>(options, TomlDoubleConverter.Instance);
        if (type == typeof(decimal)) return new TomlTypeInfo<decimal>(options, TomlDecimalConverter.Instance);
        if (type == typeof(float)) return new TomlTypeInfo<float>(options, TomlSingleConverter.Instance);
        if (type == typeof(byte)) return new TomlTypeInfo<byte>(options, Byte);
        if (type == typeof(sbyte)) return new TomlTypeInfo<sbyte>(options, SByte);
        if (type == typeof(short)) return new TomlTypeInfo<short>(options, Int16);
        if (type == typeof(ushort)) return new TomlTypeInfo<ushort>(options, UInt16);
        if (type == typeof(uint)) return new TomlTypeInfo<uint>(options, UInt32);
        if (type == typeof(ulong)) return new TomlTypeInfo<ulong>(options, TomlUInt64Converter.Instance);
        if (type == typeof(char)) return new TomlTypeInfo<char>(options, TomlCharConverter.Instance);
        if (type == typeof(BigInteger)) return new TomlTypeInfo<BigInteger>(options, TomlBigIntegerConverter.Instance);
        if (type == typeof(DateTime)) return new TomlTypeInfo<DateTime>(options, TomlDateTimeConverters.Local);
        if (type == typeof(DateTimeOffset)) return new TomlTypeInfo<DateTimeOffset>(options, TomlDateTimeConverters.Offset);
        if (type == typeof(TimeSpan)) return new TomlTypeInfo<TimeSpan>(options, TomlTimeSpanConverter.Instance);
        if (type == typeof(Guid)) return new TomlTypeInfo<Guid>(options, TomlTextConverters.GuidText);
        if (type == typeof(TomlDateTime)) return new TomlTypeInfo<TomlDateTime>(options, TomlDateTimeConverters.Toml);
#if NET6_0_OR_GREATER
        if (type == typeof(DateOnly)) return new TomlTypeInfo<DateOnly>(options, TomlDateTimeConverters.Date);
        if (type == typeof(TimeOnly)) return new TomlTypeInfo<TimeOnly>(options, TomlDateTimeConverters.Time);
#endif
        return null;
    }
}

/// <summary>The constructor and members of an object type, for <see cref="TomlMetadataServices.CreateObjectInfo{T}"/>.</summary>
/// <typeparam name="T">The object type.</typeparam>
[EditorBrowsable(EditorBrowsableState.Never)]
public sealed class TomlObjectInfoValues<T> {
    /// <summary>Gets or sets the function that creates an instance with the parameterless constructor.</summary>
    public Func<T>? ObjectCreator { get; set; }

    /// <summary>
    /// Gets or sets the function that creates an instance from constructor arguments, in parameter order. When set it is
    /// used instead of <see cref="ObjectCreator"/>.
    /// </summary>
    public Func<object?[], T>? ObjectWithParameterizedConstructorCreator { get; set; }

    /// <summary>Gets or sets the function that creates the property metadata, in writing order.</summary>
    public Func<TomlSerializerOptions, TomlPropertyInfo[]>? PropertyMetadataInitializer { get; set; }

    /// <summary>Gets or sets the function that describes the parameters of the constructor <see cref="ObjectWithParameterizedConstructorCreator"/> calls.</summary>
    public Func<TomlParameterInfoValues[]>? ConstructorParameterMetadataInitializer { get; set; }

    /// <summary>Gets or sets the registered derived types, which make the type polymorphic.</summary>
    public OfficeDerivedTypeAttribute[]? DerivedTypes { get; set; }

    /// <summary>Gets or sets the key that holds the discriminator, or null for <c>$type</c>.</summary>
    public string? DiscriminatorPropertyName { get; set; }

    /// <summary>Gets or sets whether values are written as inline tables, from <see cref="TomlInlineTableAttribute"/> on the type.</summary>
    public bool IsInlineTable { get; set; }
}

/// <summary>One property of an object type, for <see cref="TomlMetadataServices.CreatePropertyInfo{T}"/>.</summary>
/// <typeparam name="T">The property type.</typeparam>
[EditorBrowsable(EditorBrowsableState.Never)]
public sealed class TomlPropertyInfoValues<T> {
    /// <summary>Gets or sets the .NET property name.</summary>
    public string Name { get; set; } = string.Empty;

    /// <summary>Gets or sets the name from <see cref="OfficePropertyNameAttribute"/>, or null.</summary>
    public string? TomlName { get; set; }

    /// <summary>Gets or sets the getter, or null for a property that cannot be read.</summary>
    public Func<object, T>? Getter { get; set; }

    /// <summary>Gets or sets the setter, or null for a property that is not set after construction.</summary>
    public Action<object, T>? Setter { get; set; }

    /// <summary>Gets or sets whether a table must contain the property.</summary>
    public bool IsRequired { get; set; }

    /// <summary>Gets or sets the condition from <see cref="OfficeIgnoreAttribute"/>, or null for the options' default.</summary>
    public OfficeIgnoreCondition? IgnoreCondition { get; set; }

    /// <summary>Gets or sets whether the value's tables are written inline, from <see cref="TomlInlineTableAttribute"/> on the property.</summary>
    public bool IsInlineTable { get; set; }
}

/// <summary>One parameter of the constructor used to create an object.</summary>
[EditorBrowsable(EditorBrowsableState.Never)]
public sealed class TomlParameterInfoValues {
    /// <summary>Gets or sets the parameter name, which binds to the property of the same name ignoring case.</summary>
    public string Name { get; set; } = string.Empty;

    /// <summary>Gets or sets the parameter type, which equals the bound property's type.</summary>
    public Type ParameterType { get; set; } = typeof(object);

    /// <summary>Gets or sets the zero-based position.</summary>
    public int Position { get; set; }

    /// <summary>Gets or sets the value passed when the table has no key for the parameter: its default value, or the type's default.</summary>
    public object? DefaultValue { get; set; }
}
