using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace OfficeIMO.Yaml;

/// <summary>
/// Serialization metadata for one .NET type, bound to one <see cref="YamlSerializerOptions"/> instance. Get it from a
/// generated <see cref="YamlSerializerContext"/> (<c>MyContext.Default.Person</c>) or from
/// <see cref="YamlSerializerOptions.GetTypeInfo(Type)"/>, and pass it to the <see cref="YamlSerializer"/> overloads that
/// take one; those overloads use no reflection, so they work with trimming and NativeAOT.
/// </summary>
public abstract class YamlTypeInfo {
    private readonly object _gate = new object();
    private volatile bool _configured;

    /// <summary>
    /// The length of the text last serialized from this type as a root, up to <see cref="MaxOutputLengthHint"/>. The
    /// next serialization starts its buffer at that size instead of growing to it; a stale or racing value only
    /// changes how often the buffer grows.
    /// </summary>
    internal int OutputLengthHint;

#if NET8_0_OR_GREATER
    // The shared array pool keeps arrays up to 2^30 elements here, so a large document's buffer is reused, not grown.
    internal const int MaxOutputLengthHint = 1 << 23;
#else
    // The shared array pool of older frameworks keeps arrays up to 2^20 elements; larger ones would be allocated anew.
    internal const int MaxOutputLengthHint = 1 << 20;
#endif

    private protected YamlTypeInfo(Type type, YamlSerializerOptions options) {
        Type = type;
        Options = options;
    }

    /// <summary>Gets the type the metadata describes.</summary>
    public Type Type { get; }

    /// <summary>Gets the options the metadata is bound to.</summary>
    public YamlSerializerOptions Options { get; }

    /// <summary>The same type without its derived-type registrations, for a polymorphic type; otherwise this metadata.</summary>
    internal virtual YamlTypeInfo NonPolymorphic => this;

    internal abstract bool IsObject { get; }

    /// <summary>Whether values are written as one scalar by a built-in converter, so they never hold a container.</summary>
    internal abstract bool IsScalar { get; }

    internal abstract object? ReadBoxed(YamlEventCursor cursor, YamlReadState state);

    internal abstract void WriteBoxed(YamlWriter writer, object? value, YamlWriteState state);

    internal abstract void CollectBoxed(object? value, YamlReferenceCollector collector);

    /// <summary>Reads with the converter alone, for a polymorphic base that already resolved aliases, nulls and anchors.</summary>
    internal abstract object? ReadCoreBoxed(YamlEventCursor cursor, YamlReadState state);

    /// <summary>Writes with the converter alone, for a polymorphic base that already wrote anchors or entered the value.</summary>
    internal abstract void WriteCoreBoxed(YamlWriter writer, object value, YamlWriteState state);

    internal abstract void CollectCoreBoxed(object value, YamlReferenceCollector collector);

    /// <summary>The object converter, when values are object mappings.</summary>
    internal abstract IYamlObjectConverter? ObjectConverter { get; }

    /// <summary>Whether values are read untyped (<see cref="object"/>), sharing the value of an anchor among its aliases.</summary>
    internal abstract bool IsUntyped { get; }

    /// <summary>Resolves the metadata of the types this one contains, once, on first use.</summary>
    internal void EnsureConfigured() {
        if (_configured) return;
        lock (_gate) {
            if (_configured) return;
            Configure();
            _configured = true;
        }
    }

    private protected abstract void Configure();
}

/// <summary>Serialization metadata for <typeparamref name="T"/>.</summary>
/// <typeparam name="T">The described type.</typeparam>
public sealed class YamlTypeInfo<T> : YamlTypeInfo {
    private static readonly bool CanBeNull = default(T) is null;
    private static readonly bool IsReferenceType = !typeof(T).IsValueType;
    private static readonly bool IsString = typeof(T) == typeof(string);
    private readonly YamlConverter<T> _converter;
    private readonly YamlTypeInfo? _nonPolymorphic;
    // Read once: these are asked for every value written.
    private readonly bool _tracksReferences;
    private readonly bool _handlesNull;
    // Untyped values share the value read for an anchor among its aliases; typed values are read again.
    private readonly bool _sharesAliases;
    // A container of scalars only can never be its own descendant, so it needs no cycle check; set when configured.
    private bool _holdsOnlyScalars;

    internal YamlTypeInfo(YamlSerializerOptions options, YamlConverter<T> converter, YamlTypeInfo? nonPolymorphic = null)
        : base(typeof(T), options) {
        _converter = converter;
        _nonPolymorphic = nonPolymorphic;
        _tracksReferences = IsReferenceType && converter.IsContainer;
        _handlesNull = converter.HandlesNull;
        _sharesAliases = converter is YamlUntypedConverter;
        BuiltInScalar = _handlesNull ? YamlBuiltInScalar.None : converter switch {
            YamlStringConverter => YamlBuiltInScalar.String,
            YamlInt32Converter => YamlBuiltInScalar.Int32,
            YamlInt64Converter => YamlBuiltInScalar.Int64,
            YamlBooleanConverter => YamlBuiltInScalar.Boolean,
            YamlDoubleConverter => YamlBuiltInScalar.Double,
            YamlDecimalConverter => YamlBuiltInScalar.Decimal,
            _ => YamlBuiltInScalar.None
        };
    }

    /// <summary>The built-in converter that reads the values, which a property can do directly, or none.</summary>
    internal YamlBuiltInScalar BuiltInScalar { get; }

    internal YamlConverter<T> Converter => _converter;

    internal override YamlTypeInfo NonPolymorphic => _nonPolymorphic ?? this;

    internal override bool IsObject => _converter.IsObject;

    internal override bool IsUntyped => _sharesAliases;

    internal override bool IsScalar => BuiltInScalar != YamlBuiltInScalar.None || _converter.IsScalar;

    private protected override void Configure() {
        _converter.Configure(Options);
        _holdsOnlyScalars = _tracksReferences && _converter.HoldsOnlyScalars;
    }

    /// <summary>
    /// Reads the node at the cursor and leaves the cursor on the node's last event. Aliases are expanded, or with
    /// <see cref="YamlReferenceHandling.Preserve"/> mapped to the instance read for their anchor. An untyped value
    /// (<see cref="object"/>) read for an anchor is kept, and an untyped alias to it returns that value, charging the
    /// nodes it stands for, unless <see cref="YamlSerializerOptions.CopyAliases"/> asks for copies.
    /// </summary>
    internal T Read(YamlEventCursor cursor, YamlReadState state) {
        EnsureConfigured();
        if (cursor.TokenType == YamlTokenType.Alias) {
            if (state.Preserve && state.TryGetReference(cursor.GetString(), out object? existing) && existing is T reference) return reference;
            if (_sharesAliases && state.SharesAliases && cursor.AliasTarget is YamlRecording target && state.TryGetShared(target, out object? shared)) {
                cursor.ChargeAlias(target);
                return (T)shared!;
            }
            cursor.ExpandAlias();
        }
        if (!_handlesNull && cursor.IsNullScalar) {
            if (CanBeNull) return default!;
            throw YamlSerializationErrors.TypeMismatch(cursor, state, typeof(T));
        }
        if (!state.Preserve) {
            if (_sharesAliases && state.SharesAliases && cursor.DefinedAnchor is YamlRecording defined) {
                if (state.SharedTree && cursor.TokenType is YamlTokenType.MappingStart or YamlTokenType.SequenceStart) cursor.StopLogging(defined);
                T anchored = _converter.Read(cursor, state);
                state.Share(defined, anchored);
                return anchored;
            }
            return _converter.Read(cursor, state);
        }
        string? anchor = cursor.Anchor;
        state.PendingAnchor = anchor;
        T value = _converter.Read(cursor, state);
        if (anchor is not null) state.Register(anchor, value);
        return value;
    }

    /// <summary>Writes a value at the writer's position: an alias for a value already written under Preserve, otherwise the node.</summary>
    internal void Write(YamlWriter writer, T value, YamlWriteState state) {
        // The common scalars go straight to the writer, as their built-in converters would: the JIT folds these checks
        // for value types, and strings use a flag because shared generic code cannot fold them.
        if (typeof(T) == typeof(int)) {
            writer.WriteScalar((long)(int)(object)value!);
            return;
        }
        if (typeof(T) == typeof(long)) {
            writer.WriteScalar((long)(object)value!);
            return;
        }
        if (typeof(T) == typeof(bool)) {
            writer.WriteScalar((bool)(object)value!);
            return;
        }
        if (typeof(T) == typeof(double)) {
            writer.WriteScalar((double)(object)value!);
            return;
        }
        if (IsString) {
            string? text = Unsafe.As<T, string?>(ref value);
            if (text is null) writer.WriteNull();
            else writer.WriteScalar(text);
            return;
        }
        EnsureConfigured();
        if (value is null) {
            writer.WriteNull();
            return;
        }
        if (_tracksReferences) {
            if (_holdsOnlyScalars && !state.Preserve) {
                _converter.Write(writer, value, state);
                return;
            }
            object instance = value;
            if (!state.Begin(writer, instance, typeof(T))) return;
            _converter.Write(writer, value, state);
            state.End();
            return;
        }
        _converter.Write(writer, value, state);
    }

    /// <summary>Visits the value's references before a <see cref="YamlReferenceHandling.Preserve"/> write.</summary>
    internal void Collect(T value, YamlReferenceCollector collector) {
        EnsureConfigured();
        if (value is null) return;
        if (_tracksReferences && !collector.Visit(value)) return;
        _converter.Collect(value, collector);
    }

    internal T ReadCore(YamlEventCursor cursor, YamlReadState state) {
        EnsureConfigured();
        return _converter.Read(cursor, state);
    }

    internal void WriteCore(YamlWriter writer, T value, YamlWriteState state) {
        EnsureConfigured();
        _converter.Write(writer, value, state);
    }

    internal void CollectCore(T value, YamlReferenceCollector collector) {
        EnsureConfigured();
        _converter.Collect(value, collector);
    }

    internal override IYamlObjectConverter? ObjectConverter {
        get {
            EnsureConfigured();
            return _converter as IYamlObjectConverter;
        }
    }

    internal override object? ReadCoreBoxed(YamlEventCursor cursor, YamlReadState state) => ReadCore(cursor, state);

    internal override void WriteCoreBoxed(YamlWriter writer, object value, YamlWriteState state) => WriteCore(writer, (T)value, state);

    internal override void CollectCoreBoxed(object value, YamlReferenceCollector collector) => CollectCore((T)value, collector);

    internal override object? ReadBoxed(YamlEventCursor cursor, YamlReadState state) => Read(cursor, state);

    internal override void WriteBoxed(YamlWriter writer, object? value, YamlWriteState state) => Write(writer, (T)value!, state);

    internal override void CollectBoxed(object? value, YamlReferenceCollector collector) => Collect((T)value!, collector);
}

/// <summary>Reads and writes one type. Converters are internal; metadata reaches them through <see cref="YamlTypeInfo{T}"/>.</summary>
internal abstract class YamlConverter<T> {
    /// <summary>Whether <see cref="Read"/> handles null scalars itself instead of the caller mapping them to null.</summary>
    public virtual bool HandlesNull => false;

    /// <summary>Whether values are mappings or sequences whose references matter for cycles and anchors.</summary>
    public virtual bool IsContainer => false;

    /// <summary>Whether the converter reads and writes an object's members.</summary>
    public virtual bool IsObject => false;

    /// <summary>Whether the converter writes every value as one scalar, never a container.</summary>
    public virtual bool IsScalar => false;

    /// <summary>
    /// Whether a configured container converter writes only values whose metadata <see cref="YamlTypeInfo.IsScalar"/>
    /// reports, so a value can never contain itself.
    /// </summary>
    public virtual bool HoldsOnlyScalars => false;

    public virtual void Configure(YamlSerializerOptions options) {
    }

    /// <summary>Reads the node at the cursor, a non-alias and (unless <see cref="HandlesNull"/>) non-null node.</summary>
    public abstract T Read(YamlEventCursor cursor, YamlReadState state);

    /// <summary>Writes a non-null value.</summary>
    public abstract void Write(YamlWriter writer, T value, YamlWriteState state);

    public virtual void Collect(T value, YamlReferenceCollector collector) {
    }
}

/// <summary>
/// One property of an object type, created by <see cref="YamlMetadataServices.CreatePropertyInfo{T}"/> for generated
/// metadata or by reflection.
/// </summary>
[EditorBrowsable(EditorBrowsableState.Never)]
public abstract class YamlPropertyInfo {
    private protected YamlPropertyInfo(string name, string? yamlName, bool isRequired, YamlIgnoreCondition? ignoreCondition, Type propertyType) {
        ClrName = name;
        ExplicitName = yamlName;
        IsRequired = isRequired;
        IgnoreCondition = ignoreCondition;
        PropertyType = propertyType;
    }

    internal string ClrName { get; }

    internal string? ExplicitName { get; }

    /// <summary>The mapping key, after the naming policy; set when the owning type is configured.</summary>
    internal string Name { get; set; } = string.Empty;

    /// <summary>Whether <see cref="Name"/> can be written plain in every context, which skips the writer's style choice.</summary>
    internal bool NameIsPlain { get; set; }

    /// <summary>The style <see cref="Name"/> takes as a key of a block mapping, chosen once with the name.</summary>
    internal YamlStyle NameStyle { get; set; }

    internal bool IsRequired { get; }

    internal YamlIgnoreCondition? IgnoreCondition { get; }

    internal Type PropertyType { get; }

    /// <summary>The effective condition for writing, combining the property's own and the options' default.</summary>
    internal YamlIgnoreCondition WriteCondition { get; set; }

    /// <summary>The constructor parameter position the property binds to, or -1.</summary>
    internal int ParameterIndex { get; set; } = -1;

    internal bool CanRead { get; private protected set; }

    internal bool CanSet { get; private protected set; }

    internal abstract void Configure(YamlSerializerOptions options);

    /// <summary>Writes the key and value unless the ignore condition leaves the property out.</summary>
    internal abstract void Write(YamlWriter writer, object owner, YamlWriteState state);

    /// <summary>Whether the value's metadata writes one scalar; see <see cref="YamlTypeInfo.IsScalar"/>. Valid once configured.</summary>
    internal abstract bool IsScalar { get; }

    /// <summary>Reads the value at the cursor and sets it on the owner.</summary>
    internal abstract void ReadAndSet(YamlEventCursor cursor, YamlReadState state, object owner);

    /// <summary>Reads the value at the cursor, boxed, for a constructor argument or a value set after construction.</summary>
    internal abstract object? ReadBoxed(YamlEventCursor cursor, YamlReadState state);

    internal abstract void SetBoxed(object owner, object? value);

    internal abstract void Collect(object owner, YamlReferenceCollector collector);
}

/// <summary>A property whose value is <typeparamref name="T"/>.</summary>
internal sealed class YamlPropertyInfo<T> : YamlPropertyInfo {
    private static readonly bool IsValueTypeValue = typeof(T).IsValueType;
    private readonly Func<object, T>? _getter;
    private readonly Action<object, T>? _setter;
    // An instance flag, because shared generic code reads a static through a lookup.
    private readonly bool _isString = typeof(T) == typeof(string);
    private YamlTypeInfo<T>? _typeInfo;

    public YamlPropertyInfo(YamlPropertyInfoValues<T> values)
        : base(values.Name ?? throw new ArgumentException("The property name is required.", nameof(values)), values.YamlName, values.IsRequired, values.IgnoreCondition, typeof(T)) {
        _getter = values.Getter;
        _setter = values.Setter;
        CanRead = _getter is not null;
        CanSet = _setter is not null;
    }

    private YamlTypeInfo<T> TypeInfo => _typeInfo!;

    internal override bool IsScalar => TypeInfo.IsScalar;

    internal override void Configure(YamlSerializerOptions options) {
        _typeInfo = (YamlTypeInfo<T>)options.GetTypeInfo(typeof(T));
    }

    internal override void Write(YamlWriter writer, object owner, YamlWriteState state) {
        T value = _getter!(owner);
        switch (WriteCondition) {
            case YamlIgnoreCondition.WhenWritingNull:
                if (value is null) return;
                break;
            case YamlIgnoreCondition.WhenWritingDefault:
                if (value is null || (IsValueTypeValue && EqualityComparer<T>.Default.Equals(value, default!))) return;
                break;
        }
        if (NameIsPlain) {
            // Strings and the common scalars write the whole entry at once where the writer allows; the text is the
            // same as the key and the value written one by one below.
            if (typeof(T) == typeof(int)) {
                if (writer.TryWritePlainEntry(Name, (long)(int)(object)value!)) return;
            } else if (typeof(T) == typeof(long)) {
                if (writer.TryWritePlainEntry(Name, (long)(object)value!)) return;
            } else if (typeof(T) == typeof(bool)) {
                if (writer.TryWritePlainEntry(Name, (bool)(object)value!)) return;
            } else if (typeof(T) == typeof(double)) {
                if (writer.TryWritePlainEntry(Name, (double)(object)value!)) return;
            } else if (typeof(T) == typeof(decimal)) {
                if (writer.TryWritePlainEntry(Name, (decimal)(object)value!)) return;
            } else if (_isString) {
                string? text = Unsafe.As<T, string?>(ref value);
                if (text is not null && writer.TryWriteStringEntry(Name, text)) return;
            }
            writer.WritePlainKey(Name);
        } else {
            writer.WriteKey(Name, NameStyle);
        }
        TypeInfo.Write(writer, value, state);
    }

    internal override void ReadAndSet(YamlEventCursor cursor, YamlReadState state, object owner) {
        // Strings and the common numbers read a live non-null scalar as their built-in converters would, without the
        // dispatch; anything a converter would reject takes the general path, which reports it.
        if (!state.Preserve && TryReadBuiltIn(cursor, out T value)) {
            _setter!(owner, value);
            return;
        }
        _setter!(owner, TypeInfo.Read(cursor, state));
    }

    private bool TryReadBuiltIn(YamlEventCursor cursor, out T value) {
        value = default!;
        YamlBuiltInScalar kind = TypeInfo.BuiltInScalar;
        if (kind == YamlBuiltInScalar.None) return false;
        if (_isString) {
            if (kind != YamlBuiltInScalar.String || !cursor.TryReadLiveString(out string? text)) return false;
            value = Unsafe.As<string?, T>(ref text);
            return true;
        }
        if (!cursor.TryGetLiveScalar(out ReadOnlySpan<char> scalar, out YamlSchema schema)) return false;
        if (typeof(T) == typeof(int)) {
            if (kind != YamlBuiltInScalar.Int32 || !YamlSchemaResolver.TryParseInt64(scalar, schema, out long number) || number < int.MinValue || number > int.MaxValue) return false;
            int result = (int)number;
            value = Unsafe.As<int, T>(ref result);
            return true;
        }
        if (typeof(T) == typeof(long)) {
            if (kind != YamlBuiltInScalar.Int64 || !YamlSchemaResolver.TryParseInt64(scalar, schema, out long number)) return false;
            value = Unsafe.As<long, T>(ref number);
            return true;
        }
        if (typeof(T) == typeof(bool)) {
            if (kind != YamlBuiltInScalar.Boolean || !YamlSchemaResolver.TryParseBoolean(scalar, schema, out bool flag)) return false;
            value = Unsafe.As<bool, T>(ref flag);
            return true;
        }
        if (typeof(T) == typeof(double)) {
            if (kind != YamlBuiltInScalar.Double || !YamlSchemaResolver.TryParseDouble(scalar, schema, out double number)) return false;
            value = Unsafe.As<double, T>(ref number);
            return true;
        }
        if (typeof(T) == typeof(decimal)) {
            if (kind != YamlBuiltInScalar.Decimal || !YamlSchemaResolver.TryParseDecimal(scalar, schema, out decimal number)) return false;
            value = Unsafe.As<decimal, T>(ref number);
            return true;
        }
        return false;
    }

    internal override object? ReadBoxed(YamlEventCursor cursor, YamlReadState state) => TypeInfo.Read(cursor, state);

    internal override void SetBoxed(object owner, object? value) => _setter!(owner, (T)value!);

    internal override void Collect(object owner, YamlReferenceCollector collector) => TypeInfo.Collect(_getter!(owner), collector);
}

/// <summary>The built-in scalar converters a property reads through directly.</summary>
internal enum YamlBuiltInScalar : byte { None, String, Int32, Int64, Boolean, Double, Decimal }
