namespace OfficeIMO.Toml;

/// <summary>The object half of a converter, for a polymorphic base that chose the derived type.</summary>
internal interface ITomlObjectConverter {
    /// <summary>Reads an object from a table.</summary>
    object ReadTable(TomlReadNode table, TomlReadState state);

    /// <summary>Gets whether the object has a property written under <paramref name="name"/>.</summary>
    bool HasProperty(string name);

    /// <summary>Gets whether the type can be created when reading.</summary>
    bool CanCreate { get; }
}

/// <summary>
/// Objects as tables of their properties. Reading creates the object with its parameterless constructor and sets the
/// properties whose keys the table has, or, when a constructor with parameters is used, collects the values first and
/// passes the bound ones to the constructor. Unknown keys are skipped. Writing puts the key/value pairs first and the
/// properties that are tables or arrays of tables after them as <c>[key]</c> and <c>[[key]]</c> sections, as TOML
/// requires; inside an array or inline table the object is an inline table.
/// </summary>
internal sealed class TomlObjectConverter<T> : TomlConverter<T>, ITomlObjectConverter {
    private readonly Func<T>? _create;
    private readonly Func<object?[], T>? _createWithArguments;
    private readonly Func<TomlSerializerOptions, TomlPropertyInfo[]>? _propertyInitializer;
    private readonly Func<TomlParameterInfoValues[]>? _parameterInitializer;
    private readonly bool _inline;
    private TomlPropertyInfo[] _properties = Array.Empty<TomlPropertyInfo>();
    private TomlPropertyInfo[] _written = Array.Empty<TomlPropertyInfo>();
    private TomlParameterInfoValues[] _parameters = Array.Empty<TomlParameterInfoValues>();
    private Dictionary<string, int> _byName = new Dictionary<string, int>(StringComparer.Ordinal);
    private bool _hasRequired;

    public TomlObjectConverter(TomlObjectInfoValues<T> values) {
        _create = values.ObjectCreator;
        _createWithArguments = values.ObjectWithParameterizedConstructorCreator;
        _propertyInitializer = values.PropertyMetadataInitializer;
        _parameterInitializer = values.ConstructorParameterMetadataInitializer;
        _inline = values.IsInlineTable;
    }

    public override TomlShape Shape => _inline ? TomlShape.InlineTable : TomlShape.Table;

    public bool CanCreate => _create is not null || _createWithArguments is not null;

    public override void Configure(TomlSerializerOptions options) {
        TomlPropertyInfo[] properties = _propertyInitializer?.Invoke(options) ?? Array.Empty<TomlPropertyInfo>();
        OfficeNamingPolicy? policy = options.PropertyNamingPolicy;
        var byName = new Dictionary<string, int>(properties.Length, StringComparer.Ordinal);
        for (int i = 0; i < properties.Length; i++) {
            TomlPropertyInfo property = properties[i];
            property.Configure(options);
            string name = property.ExplicitName ?? (policy is null ? property.ClrName : policy.ConvertName(property.ClrName));
            if (byName.ContainsKey(name)) {
                throw new InvalidOperationException("The type " + TomlSerializationErrors.TypeName(typeof(T)) + " has more than one property named '" + name + "' in TOML.");
            }
            property.Name = name;
            property.KeyText = TomlEmitter.FormatKey(name);
            property.WriteCondition = property.IgnoreCondition ?? options.DefaultIgnoreCondition;
            byName.Add(name, i);
            if (property.IsRequired) _hasRequired = true;
        }

        TomlParameterInfoValues[] parameters = _parameterInitializer?.Invoke() ?? Array.Empty<TomlParameterInfoValues>();
        foreach (TomlParameterInfoValues parameter in parameters) {
            int match = -1;
            for (int i = 0; i < properties.Length; i++) {
                if (string.Equals(properties[i].ClrName, parameter.Name, StringComparison.OrdinalIgnoreCase) && properties[i].PropertyType == parameter.ParameterType) {
                    match = i;
                    break;
                }
            }
            if (match < 0) {
                throw new InvalidOperationException("The constructor parameter '" + parameter.Name + "' of " + TomlSerializationErrors.TypeName(typeof(T))
                    + " has no property with the same name and type.");
            }
            properties[match].ParameterIndex = parameter.Position;
        }

        _properties = properties;
        _written = Array.FindAll(properties, static property => property.CanRead);
        _parameters = parameters;
        _byName = byName;
    }

    public bool HasProperty(string name) => _byName.ContainsKey(name);

    /// <summary>Writes the discriminator a polymorphic base left for this object, as its first pair.</summary>
    private static void WriteDiscriminator(TomlEmitter emitter, TomlWriteState state) {
        if (state.DiscriminatorKeyText is not string key) return;
        string value = state.DiscriminatorValue!;
        state.DiscriminatorKeyText = null;
        state.DiscriminatorValue = null;
        emitter.WriteKeyText(key);
        emitter.WriteString(value);
        emitter.EndValue();
    }

    public override void Write(TomlEmitter emitter, T value, TomlWriteState state) {
        object owner = value!;
        emitter.BeginInlineTable();
        WriteDiscriminator(emitter, state);
        TomlPropertyInfo[] properties = _written;
        for (int i = 0; i < properties.Length; i++) properties[i].WritePair(emitter, owner, state, inline: true);
        emitter.EndInlineTable();
    }

    public override void WriteBody(TomlEmitter emitter, T value, TomlWriteState state) {
        object owner = value!;
        WriteDiscriminator(emitter, state);
        TomlPropertyInfo[] properties = _written;
        // Sections follow every pair; their positions are kept on the stack for the usual small type.
        int count = properties.Length;
        int[]? rented = null;
        Span<int> sections = count <= 64 ? stackalloc int[count] : (rented = ArrayPool<int>.Shared.Rent(count)).AsSpan(0, count);
        int deferred = 0;
        try {
            for (int i = 0; i < count; i++) {
                if (properties[i].WritePair(emitter, owner, state, inline: false)) sections[deferred++] = i;
            }
            for (int i = 0; i < deferred; i++) properties[sections[i]].WriteSection(emitter, owner, state);
        } finally {
            if (rented is not null) ArrayPool<int>.Shared.Return(rented);
        }
    }

    public override T Read(TomlReadNode node, TomlReadState state) {
        if (!node.IsTable) throw TomlSerializationErrors.TypeMismatch(node, typeof(T));
        return (T)ReadTable(node, state);
    }

    public object ReadTable(TomlReadNode table, TomlReadState state) {
        if (!CanCreate) {
            throw new NotSupportedException("The type " + TomlSerializationErrors.TypeName(typeof(T))
                + " cannot be created: it needs a public parameterless constructor, a single public constructor or one marked [OfficeConstructor].");
        }
        TomlPropertyInfo[] properties = _properties;
        int count = properties.Length;
        bool[]? rented = null;
        Span<bool> found = count <= 128 ? stackalloc bool[count] : (rented = ArrayPool<bool>.Shared.Rent(count)).AsSpan(0, count);
        if (rented is not null) found.Clear();
        try {
            object owner;
            int expected = 0;
            if (_createWithArguments is null) {
                owner = _create!()!;
                foreach (TomlReadNode entry in table.GetChildren()) {
                    int index = Find(entry.Key!, ref expected);
                    if (index < 0 || !properties[index].CanSet) continue;
                    properties[index].ReadAndSet(entry, state, owner);
                    found[index] = true;
                }
            } else {
                var values = new object?[count];
                foreach (TomlReadNode entry in table.GetChildren()) {
                    int index = Find(entry.Key!, ref expected);
                    if (index < 0 || (!properties[index].CanSet && properties[index].ParameterIndex < 0)) continue;
                    values[index] = properties[index].ReadBoxed(entry, state);
                    found[index] = true;
                }
                owner = Construct(values, found);
            }
            if (_hasRequired) {
                for (int i = 0; i < count; i++) {
                    if (properties[i].IsRequired && !found[i]) throw TomlSerializationErrors.MissingRequired(table, typeof(T), properties[i].Name);
                }
            }
            return owner;
        } finally {
            if (rented is not null) ArrayPool<bool>.Shared.Return(rented);
        }
    }

    private object Construct(object?[] values, Span<bool> found) {
        TomlParameterInfoValues[] parameters = _parameters;
        var arguments = new object?[parameters.Length];
        for (int i = 0; i < parameters.Length; i++) arguments[i] = parameters[i].DefaultValue;
        TomlPropertyInfo[] properties = _properties;
        for (int i = 0; i < properties.Length; i++) {
            if (found[i] && properties[i].ParameterIndex >= 0) arguments[properties[i].ParameterIndex] = values[i];
        }
        object owner = _createWithArguments!(arguments)!;
        for (int i = 0; i < properties.Length; i++) {
            if (found[i] && properties[i].ParameterIndex < 0) properties[i].SetBoxed(owner, values[i]);
        }
        return owner;
    }

    /// <summary>Finds the property for a key. Keys usually arrive in declaration order, so the next property is tried first.</summary>
    private int Find(string key, ref int expected) {
        TomlPropertyInfo[] properties = _properties;
        if ((uint)expected < (uint)properties.Length && string.Equals(properties[expected].Name, key, StringComparison.Ordinal)) return expected++;
        if (!_byName.TryGetValue(key, out int index)) return -1;
        expected = index + 1;
        return index;
    }
}

/// <summary>
/// A type with <see cref="OfficeDerivedTypeAttribute"/> registrations. Values are written with their runtime type's
/// members and its discriminator as the first key, <c>$type</c> unless
/// <see cref="OfficePolymorphicAttribute.DiscriminatorPropertyName"/> names another. Reading finds the discriminator
/// anywhere in the table and creates only a registered derived type; without one, the base type itself when it can
/// be created.
/// </summary>
internal sealed class TomlPolymorphicConverter<T> : TomlConverter<T> {
    internal const string DefaultDiscriminator = "$type";

    private readonly TomlTypeInfo<T> _base;
    private readonly OfficeDerivedTypeAttribute[] _registrations;
    private readonly string _propertyName;
    private readonly string _keyText;
    private Dictionary<Type, (TomlTypeInfo Info, string Discriminator)>? _byType;
    private Dictionary<string, TomlTypeInfo>? _byDiscriminator;

    public TomlPolymorphicConverter(TomlTypeInfo<T> baseInfo, OfficeDerivedTypeAttribute[] registrations, string? propertyName) {
        _base = baseInfo;
        _registrations = registrations;
        _propertyName = propertyName ?? DefaultDiscriminator;
        _keyText = TomlEmitter.FormatKey(_propertyName);
    }

    public override TomlShape Shape => TomlShape.Table;

    public override bool HasDynamicShape => true;

    public override bool TracksReferences => false;

    public override void Configure(TomlSerializerOptions options) {
        var byType = new Dictionary<Type, (TomlTypeInfo, string)>();
        var byDiscriminator = new Dictionary<string, TomlTypeInfo>(StringComparer.Ordinal);
        foreach (OfficeDerivedTypeAttribute registration in _registrations) {
            if (!typeof(T).IsAssignableFrom(registration.DerivedType) || registration.DerivedType == typeof(T)) {
                throw new InvalidOperationException("The type " + TomlSerializationErrors.TypeName(registration.DerivedType) + " registered on "
                    + TomlSerializationErrors.TypeName(typeof(T)) + " does not derive from it.");
            }
            TomlTypeInfo derived = options.GetTypeInfo(registration.DerivedType).NonPolymorphic;
            if (derived.ObjectConverter is null) {
                throw new InvalidOperationException("The derived type " + TomlSerializationErrors.TypeName(registration.DerivedType) + " must be written as a table.");
            }
            if (byDiscriminator.ContainsKey(registration.Discriminator) || byType.ContainsKey(registration.DerivedType)) {
                throw new InvalidOperationException("The type " + TomlSerializationErrors.TypeName(typeof(T)) + " registers the discriminator '"
                    + registration.Discriminator + "' or the type " + TomlSerializationErrors.TypeName(registration.DerivedType) + " more than once.");
            }
            byType.Add(registration.DerivedType, (derived, registration.Discriminator));
            byDiscriminator.Add(registration.Discriminator, derived);
        }
        _byType = byType;
        _byDiscriminator = byDiscriminator;
    }

    public override TomlShape GetShape(T value) {
        Type runtime = value!.GetType();
        if (runtime == typeof(T)) return _base.StaticShape;
        return _byType!.TryGetValue(runtime, out (TomlTypeInfo Info, string Discriminator) entry) ? entry.Info.StaticShape : TomlShape.Table;
    }

    /// <summary>Finds the derived metadata for a value and leaves its discriminator for the object to write first.</summary>
    private TomlTypeInfo? Derived(T value, TomlWriteState state) {
        Type runtime = value!.GetType();
        if (runtime == typeof(T)) return null;
        if (!_byType!.TryGetValue(runtime, out (TomlTypeInfo Info, string Discriminator) entry)) {
            throw new NotSupportedException("The runtime type " + TomlSerializationErrors.TypeName(runtime) + " is not registered with [OfficeDerivedType] on "
                + TomlSerializationErrors.TypeName(typeof(T)) + ".");
        }
        if (entry.Info.ObjectConverter!.HasProperty(_propertyName)) {
            throw new InvalidOperationException("The type " + TomlSerializationErrors.TypeName(runtime) + " has a property named '" + _propertyName
                + "', which " + TomlSerializationErrors.TypeName(typeof(T)) + " uses for its discriminator.");
        }
        state.DiscriminatorKeyText = _keyText;
        state.DiscriminatorValue = entry.Discriminator;
        return entry.Info;
    }

    public override void Write(TomlEmitter emitter, T value, TomlWriteState state) {
        TomlTypeInfo? derived = Derived(value, state);
        if (derived is null) _base.Write(emitter, value, state);
        else derived.WriteBoxed(emitter, value!, state);
    }

    public override void WriteBody(TomlEmitter emitter, T value, TomlWriteState state) {
        TomlTypeInfo? derived = Derived(value, state);
        if (derived is null) _base.WriteBody(emitter, value, state);
        else derived.WriteBodyBoxed(emitter, value!, state);
    }

    public override T Read(TomlReadNode node, TomlReadState state) {
        if (!node.IsTable) throw TomlSerializationErrors.TypeMismatch(node, typeof(T));
        if (!node.TryGetChild(_propertyName, out TomlReadNode discriminatorNode)) {
            if (!_base.ObjectConverter!.CanCreate) throw TomlSerializationErrors.UnknownDiscriminator(node, typeof(T), null);
            return _base.Read(node, state);
        }
        if (discriminatorNode.Kind != TomlNodeKind.String) throw TomlSerializationErrors.TypeMismatch(discriminatorNode, typeof(string), "a discriminator must be a string");
        string discriminator = TomlValueText.GetString(discriminatorNode);
        if (!_byDiscriminator!.TryGetValue(discriminator, out TomlTypeInfo? info)) throw TomlSerializationErrors.UnknownDiscriminator(discriminatorNode, typeof(T), discriminator);
        return (T)info.ObjectConverter!.ReadTable(node, state);
    }
}
