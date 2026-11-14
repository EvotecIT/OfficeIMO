using System.ComponentModel;

namespace OfficeIMO.Toml;

/// <summary>
/// Serialization metadata for one .NET type, bound to one <see cref="TomlSerializerOptions"/> instance. Get it from a
/// generated <see cref="TomlSerializerContext"/> (<c>MyContext.Default.Config</c>) or from
/// <see cref="TomlSerializerOptions.GetTypeInfo(Type)"/>, and pass it to the <see cref="TomlSerializer"/> overloads that
/// take one; those overloads use no reflection, so they work with trimming and NativeAOT.
/// </summary>
public abstract class TomlTypeInfo {
    private readonly object _gate = new object();
    private volatile bool _configured;

    private protected TomlTypeInfo(Type type, TomlSerializerOptions options) {
        Type = type;
        Options = options;
    }

    /// <summary>Gets the type the metadata describes.</summary>
    public Type Type { get; }

    /// <summary>Gets the options the metadata is bound to.</summary>
    public TomlSerializerOptions Options { get; }

    /// <summary>The same type without its derived-type registrations, for a polymorphic type; otherwise this metadata.</summary>
    internal virtual TomlTypeInfo NonPolymorphic => this;

    /// <summary>The object converter, when values are objects written as tables of their properties.</summary>
    internal abstract ITomlObjectConverter? ObjectConverter { get; }

    /// <summary>How values of the type are written when the type alone decides it.</summary>
    internal abstract TomlShape StaticShape { get; }

    internal abstract object? ReadBoxed(TomlReadNode node, TomlReadState state);

    internal abstract TomlShape GetShapeBoxed(object value);

    internal abstract void WriteBoxed(TomlEmitter emitter, object value, TomlWriteState state);

    internal abstract void WriteBodyBoxed(TomlEmitter emitter, object value, TomlWriteState state);

    internal abstract void WriteTableArrayBoxed(TomlEmitter emitter, object value, TomlWriteState state);

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
public sealed class TomlTypeInfo<T> : TomlTypeInfo {
    private readonly TomlConverter<T> _converter;
    private readonly TomlTypeInfo? _nonPolymorphic;
    // Set when configured: whether values are reference-type containers, which a cycle check guards.
    private bool _tracksReferences;

    internal TomlTypeInfo(TomlSerializerOptions options, TomlConverter<T> converter, TomlTypeInfo? nonPolymorphic = null)
        : base(typeof(T), options) {
        _converter = converter;
        _nonPolymorphic = nonPolymorphic;
    }

    internal TomlConverter<T> Converter => _converter;

    internal override TomlTypeInfo NonPolymorphic => _nonPolymorphic ?? this;

    internal override TomlShape StaticShape {
        get {
            EnsureConfigured();
            return _converter.Shape;
        }
    }

    internal override ITomlObjectConverter? ObjectConverter {
        get {
            EnsureConfigured();
            return _converter as ITomlObjectConverter;
        }
    }

    private protected override void Configure() {
        _converter.Configure(Options);
        _tracksReferences = !typeof(T).IsValueType && _converter.TracksReferences;
    }

    /// <summary>Reads a node as <typeparamref name="T"/>.</summary>
    internal T Read(TomlReadNode node, TomlReadState state) {
        EnsureConfigured();
        return _converter.Read(node, state);
    }

    /// <summary>Returns how a non-null value is written: by its type, or by the value for untyped and polymorphic types.</summary>
    internal TomlShape GetShape(T value) {
        EnsureConfigured();
        return _converter.HasDynamicShape ? _converter.GetShape(value) : _converter.Shape;
    }

    /// <summary>Writes a non-null value in its inline form: a scalar, an array or an inline table.</summary>
    internal void Write(TomlEmitter emitter, T value, TomlWriteState state) {
        EnsureConfigured();
        if (!_tracksReferences) {
            _converter.Write(emitter, value, state);
            return;
        }
        state.Enter(value!, typeof(T));
        _converter.Write(emitter, value, state);
        state.Exit();
    }

    /// <summary>Writes the key/value pairs and sections of a table-shaped value under the emitter's current header.</summary>
    internal void WriteBody(TomlEmitter emitter, T value, TomlWriteState state) {
        EnsureConfigured();
        if (!_tracksReferences) {
            _converter.WriteBody(emitter, value, state);
            return;
        }
        state.Enter(value!, typeof(T));
        _converter.WriteBody(emitter, value, state);
        state.Exit();
    }

    /// <summary>Writes a collection of tables as one <c>[[header]]</c> section per item.</summary>
    internal void WriteTableArray(TomlEmitter emitter, T value, TomlWriteState state) {
        EnsureConfigured();
        if (!_tracksReferences) {
            _converter.WriteTableArray(emitter, value, state);
            return;
        }
        state.Enter(value!, typeof(T));
        _converter.WriteTableArray(emitter, value, state);
        state.Exit();
    }

    internal override object? ReadBoxed(TomlReadNode node, TomlReadState state) => Read(node, state);

    internal override TomlShape GetShapeBoxed(object value) => GetShape((T)value);

    internal override void WriteBoxed(TomlEmitter emitter, object value, TomlWriteState state) => Write(emitter, (T)value, state);

    internal override void WriteBodyBoxed(TomlEmitter emitter, object value, TomlWriteState state) => WriteBody(emitter, (T)value, state);

    internal override void WriteTableArrayBoxed(TomlEmitter emitter, object value, TomlWriteState state) => WriteTableArray(emitter, (T)value, state);
}

/// <summary>How a value is written.</summary>
internal enum TomlShape : byte {
    /// <summary>A string, number, boolean or date-time on its key's line.</summary>
    Scalar,
    /// <summary>An array on its key's line.</summary>
    Array,
    /// <summary>A table written inline on its key's line (<c>{ a = 1 }</c>).</summary>
    InlineTable,
    /// <summary>A table written as a <c>[header]</c> section, or inline inside arrays and inline tables.</summary>
    Table,
    /// <summary>A non-empty collection of tables written as <c>[[header]]</c> sections, or an inline array inside inline values.</summary>
    TableArray
}

/// <summary>Reads and writes one type. Converters are internal; metadata reaches them through <see cref="TomlTypeInfo{T}"/>.</summary>
internal abstract class TomlConverter<T> {
    /// <summary>How values are written when the type decides it.</summary>
    public virtual TomlShape Shape => TomlShape.Scalar;

    /// <summary>Whether the shape depends on the value: untyped values, polymorphic types and collections of them.</summary>
    public virtual bool HasDynamicShape => false;

    /// <summary>Whether values are dictionaries, whose entries a <see cref="TomlInlineTableAttribute"/> property writes inline.</summary>
    public virtual bool IsDictionary => false;

    /// <summary>
    /// Whether values are containers whose references a cycle check guards. Converters that hand a value to another
    /// type's metadata (untyped and polymorphic values) leave the check to it.
    /// </summary>
    public virtual bool TracksReferences => Shape != TomlShape.Scalar;

    public virtual TomlShape GetShape(T value) => Shape;

    public virtual void Configure(TomlSerializerOptions options) {
    }

    /// <summary>Reads a node.</summary>
    public abstract T Read(TomlReadNode node, TomlReadState state);

    /// <summary>Writes a non-null value in its inline form.</summary>
    public abstract void Write(TomlEmitter emitter, T value, TomlWriteState state);

    /// <summary>Writes a table-shaped value's pairs and sections; only table converters are asked.</summary>
    public virtual void WriteBody(TomlEmitter emitter, T value, TomlWriteState state) => throw new InvalidOperationException("Only tables have a body.");

    /// <summary>Writes a collection of tables as <c>[[header]]</c> sections; only collection converters are asked.</summary>
    public virtual void WriteTableArray(TomlEmitter emitter, T value, TomlWriteState state) => throw new InvalidOperationException("Only collections are arrays of tables.");

    /// <summary>Reads a dictionary key. Scalar converters read the key as they would read a string value's text.</summary>
    public virtual bool TryReadKey(string key, out T value) {
        value = default!;
        return false;
    }

    /// <summary>Returns the key text for a dictionary key, or null when the type cannot be a key.</summary>
    public virtual string? WriteKey(T value) => null;
}

/// <summary>
/// One property of an object type, created by <see cref="TomlMetadataServices.CreatePropertyInfo{T}"/> for generated
/// metadata or by reflection.
/// </summary>
[EditorBrowsable(EditorBrowsableState.Never)]
public abstract class TomlPropertyInfo {
    private protected TomlPropertyInfo(string name, string? tomlName, bool isRequired, OfficeIgnoreCondition? ignoreCondition, bool isInlineTable, Type propertyType) {
        ClrName = name;
        ExplicitName = tomlName;
        IsRequired = isRequired;
        IgnoreCondition = ignoreCondition;
        IsInlineTable = isInlineTable;
        PropertyType = propertyType;
    }

    internal string ClrName { get; }

    internal string? ExplicitName { get; }

    /// <summary>The key, after the naming policy; set when the owning type is configured.</summary>
    internal string Name { get; set; } = string.Empty;

    /// <summary>The key as TOML text, bare or quoted, followed by <c> = </c>; set with <see cref="Name"/>.</summary>
    internal string KeyText { get; set; } = string.Empty;

    internal bool IsRequired { get; }

    internal OfficeIgnoreCondition? IgnoreCondition { get; }

    internal bool IsInlineTable { get; }

    internal Type PropertyType { get; }

    /// <summary>The effective condition for writing, combining the property's own and the options' default.</summary>
    internal OfficeIgnoreCondition WriteCondition { get; set; }

    /// <summary>The constructor parameter position the property binds to, or -1.</summary>
    internal int ParameterIndex { get; set; } = -1;

    internal bool CanRead { get; private protected set; }

    internal bool CanSet { get; private protected set; }

    internal abstract void Configure(TomlSerializerOptions options);

    /// <summary>
    /// Writes the property as a key/value pair, or returns true when its value is a section that must follow the pairs.
    /// With <paramref name="inline"/> every value is written on the key's line.
    /// </summary>
    internal abstract bool WritePair(TomlEmitter emitter, object owner, TomlWriteState state, bool inline);

    /// <summary>Writes a value <see cref="WritePair"/> deferred as a <c>[table]</c> or <c>[[table]]</c> section.</summary>
    internal abstract void WriteSection(TomlEmitter emitter, object owner, TomlWriteState state);

    /// <summary>Reads the node and sets it on the owner.</summary>
    internal abstract void ReadAndSet(TomlReadNode node, TomlReadState state, object owner);

    /// <summary>Reads the node, boxed, for a constructor argument or a value set after construction.</summary>
    internal abstract object? ReadBoxed(TomlReadNode node, TomlReadState state);

    internal abstract void SetBoxed(object owner, object? value);
}

/// <summary>A property whose value is <typeparamref name="T"/>.</summary>
internal sealed class TomlPropertyInfo<T> : TomlPropertyInfo {
    private static readonly bool IsValueTypeValue = typeof(T).IsValueType;
    private readonly Func<object, T>? _getter;
    private readonly Action<object, T>? _setter;
    private TomlTypeInfo<T>? _typeInfo;
    private bool _dictionary;

    public TomlPropertyInfo(TomlPropertyInfoValues<T> values)
        : base(values.Name ?? throw new ArgumentException("The property name is required.", nameof(values)), values.TomlName, values.IsRequired, values.IgnoreCondition,
            values.IsInlineTable, typeof(T)) {
        _getter = values.Getter;
        _setter = values.Setter;
        CanRead = _getter is not null;
        CanSet = _setter is not null;
    }

    private TomlTypeInfo<T> TypeInfo => _typeInfo!;

    internal override void Configure(TomlSerializerOptions options) {
        _typeInfo = (TomlTypeInfo<T>)options.GetTypeInfo(typeof(T));
        _dictionary = _typeInfo.Converter.IsDictionary;
    }

    /// <summary>Whether a value is left out: null, or a default the ignore condition skips.</summary>
    private bool Skip(T value, TomlWriteState state) {
        if (value is null) {
            if (WriteCondition is OfficeIgnoreCondition.WhenWritingNull or OfficeIgnoreCondition.WhenWritingDefault || !state.FailOnNull) return true;
            throw TomlSerializationErrors.NullValue(Name);
        }
        return WriteCondition == OfficeIgnoreCondition.WhenWritingDefault && IsValueTypeValue && EqualityComparer<T>.Default.Equals(value, default!);
    }

    private TomlShape ShapeOf(T value) {
        TomlShape shape = TypeInfo.GetShape(value);
        if (!IsInlineTable) return shape;
        if (shape == TomlShape.Table && !_dictionary) return TomlShape.InlineTable;
        return shape == TomlShape.TableArray ? TomlShape.Array : shape;
    }

    internal override bool WritePair(TomlEmitter emitter, object owner, TomlWriteState state, bool inline) {
        T value = _getter!(owner);
        if (Skip(value, state)) return false;
        if (!inline) {
            TomlShape shape = ShapeOf(value);
            if (shape is TomlShape.Table or TomlShape.TableArray) return true;
        }
        emitter.WriteKeyText(KeyText);
        TypeInfo.Write(emitter, value, state);
        emitter.EndValue();
        return false;
    }

    internal override void WriteSection(TomlEmitter emitter, object owner, TomlWriteState state) {
        T value = _getter!(owner);
        if (ShapeOf(value) == TomlShape.TableArray) {
            emitter.PushKey(Name);
            TypeInfo.WriteTableArray(emitter, value, state);
            emitter.PopKey();
            return;
        }
        emitter.PushTable(Name);
        // A dictionary property marked inline keeps its section and writes its entries' tables inline.
        bool inlineEntries = IsInlineTable && _dictionary;
        state.InlineEntries = inlineEntries;
        TypeInfo.WriteBody(emitter, value, state);
        emitter.PopTable();
    }

    internal override void ReadAndSet(TomlReadNode node, TomlReadState state, object owner) => _setter!(owner, TypeInfo.Read(node, state));

    internal override object? ReadBoxed(TomlReadNode node, TomlReadState state) => TypeInfo.Read(node, state);

    internal override void SetBoxed(object owner, object? value) => _setter!(owner, (T)value!);
}
