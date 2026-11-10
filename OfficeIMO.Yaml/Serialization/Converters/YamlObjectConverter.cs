namespace OfficeIMO.Yaml;

/// <summary>Reads the entries of an object mapping after its start event, for a polymorphic base that consumed the start.</summary>
internal interface IYamlObjectConverter {
    /// <summary>Reads the remaining entries of a mapping whose start (and possibly some entries) the caller has read.</summary>
    object ReadBody(YamlEventCursor cursor, YamlReadState state, YamlPosition start);

    /// <summary>Gets whether the object has a property written under <paramref name="name"/>.</summary>
    bool HasProperty(string name);

    /// <summary>Gets whether the type can be created when reading.</summary>
    bool CanCreate { get; }
}

/// <summary>
/// Objects as mappings of their properties. Reading creates the object with its parameterless constructor and sets
/// properties as their keys arrive, or, when a constructor with parameters is used, collects the values first and
/// passes the bound ones to the constructor. Unknown keys are skipped. Merge keys (<c>&lt;&lt;</c>) set the properties
/// the mapping does not set itself.
/// </summary>
internal sealed class YamlObjectConverter<T> : YamlConverter<T>, IYamlObjectConverter {
    private const byte Unset = 0;
    private const byte Merged = 1;
    private const byte Explicit = 2;

    private readonly Func<T>? _create;
    private readonly Func<object?[], T>? _createWithArguments;
    private readonly Func<YamlSerializerOptions, YamlPropertyInfo[]>? _propertyInitializer;
    private readonly Func<YamlParameterInfoValues[]>? _parameterInitializer;
    private YamlPropertyInfo[] _properties = Array.Empty<YamlPropertyInfo>();
    private YamlParameterInfoValues[] _parameters = Array.Empty<YamlParameterInfoValues>();
    private Dictionary<string, int> _byName = new Dictionary<string, int>(StringComparer.Ordinal);
#if NET9_0_OR_GREATER
    private Dictionary<string, int>.AlternateLookup<ReadOnlySpan<char>> _bySpan;
#endif
    private bool _hasRequired;

    public YamlObjectConverter(YamlObjectInfoValues<T> values) {
        _create = values.ObjectCreator;
        _createWithArguments = values.ObjectWithParameterizedConstructorCreator;
        _propertyInitializer = values.PropertyMetadataInitializer;
        _parameterInitializer = values.ConstructorParameterMetadataInitializer;
    }

    public override bool IsContainer => true;

    public override bool IsObject => true;

    public bool CanCreate => _create is not null || _createWithArguments is not null;

    private bool UsesArguments => _createWithArguments is not null;

    public override void Configure(YamlSerializerOptions options) {
        YamlPropertyInfo[] properties = _propertyInitializer?.Invoke(options) ?? Array.Empty<YamlPropertyInfo>();
        YamlNamingPolicy? policy = options.PropertyNamingPolicy;
        var byName = new Dictionary<string, int>(properties.Length, StringComparer.Ordinal);
        for (int i = 0; i < properties.Length; i++) {
            YamlPropertyInfo property = properties[i];
            property.Configure(options);
            string name = property.ExplicitName ?? (policy is null ? property.ClrName : policy.ConvertName(property.ClrName));
            if (byName.ContainsKey(name)) {
                throw new InvalidOperationException("The type " + YamlSerializationErrors.TypeName(typeof(T)) + " has more than one property named '" + name + "' in YAML.");
            }
            property.Name = name;
            property.NameIsPlain = YamlScalarFormat.CanBePlain(name.AsSpan(), flow: true, key: true) && !YamlScalarFormat.IsSchemaAmbiguous(name.AsSpan());
            property.WriteCondition = property.IgnoreCondition ?? options.DefaultIgnoreCondition;
            byName.Add(name, i);
            if (property.IsRequired) _hasRequired = true;
        }

        YamlParameterInfoValues[] parameters = _parameterInitializer?.Invoke() ?? Array.Empty<YamlParameterInfoValues>();
        foreach (YamlParameterInfoValues parameter in parameters) {
            int match = -1;
            for (int i = 0; i < properties.Length; i++) {
                if (string.Equals(properties[i].ClrName, parameter.Name, StringComparison.OrdinalIgnoreCase) && properties[i].PropertyType == parameter.ParameterType) {
                    match = i;
                    break;
                }
            }
            if (match < 0) {
                throw new InvalidOperationException("The constructor parameter '" + parameter.Name + "' of " + YamlSerializationErrors.TypeName(typeof(T))
                    + " has no property with the same name and type.");
            }
            properties[match].ParameterIndex = parameter.Position;
        }

        _properties = properties;
        _parameters = parameters;
        _byName = byName;
#if NET9_0_OR_GREATER
        _bySpan = byName.GetAlternateLookup<ReadOnlySpan<char>>();
#endif
    }

    public bool HasProperty(string name) => _byName.ContainsKey(name);

    public override void Write(YamlWriter writer, T value, YamlWriteState state) {
        object owner = value!;
        writer.WriteStartMapping();
        if (state.DiscriminatorKey is string key) {
            string discriminator = state.DiscriminatorValue!;
            state.DiscriminatorKey = null;
            state.DiscriminatorValue = null;
            writer.WriteKey(key);
            writer.WriteScalar(discriminator);
        }
        YamlPropertyInfo[] properties = _properties;
        for (int i = 0; i < properties.Length; i++) {
            YamlPropertyInfo property = properties[i];
            if (property.CanRead) property.Write(writer, owner, state);
        }
        writer.WriteEndMapping();
    }

    public override void Collect(T value, YamlReferenceCollector collector) {
        object owner = value!;
        foreach (YamlPropertyInfo property in _properties) {
            if (property.CanRead) property.Collect(owner, collector);
        }
    }

    public override T Read(YamlEventCursor cursor, YamlReadState state) {
        if (cursor.TokenType != YamlTokenType.MappingStart) throw YamlSerializationErrors.TypeMismatch(cursor, state, typeof(T));
        return (T)ReadBody(cursor, state, cursor.Position);
    }

    public object ReadBody(YamlEventCursor cursor, YamlReadState state, YamlPosition start) {
        if (!CanCreate) {
            throw new NotSupportedException("The type " + YamlSerializationErrors.TypeName(typeof(T))
                + " cannot be created: it needs a public parameterless constructor, a single public constructor or one marked [YamlConstructor].");
        }
        int count = _properties.Length;
        byte[]? rented = null;
        Span<byte> flags = count <= 128 ? stackalloc byte[count] : (rented = ArrayPool<byte>.Shared.Rent(count)).AsSpan(0, count);
        flags.Clear();
        try {
            object owner;
            if (!UsesArguments) {
                owner = _create!()!;
                if (state.Preserve && !typeof(T).IsValueType) state.TakePendingAnchor(owner);
                int expected = 0;
                ReadMembers(cursor, state, owner, null, flags, ref expected, merging: false);
            } else {
                var values = new object?[count];
                int expected = 0;
                ReadMembers(cursor, state, null, values, flags, ref expected, merging: false);
                owner = Construct(values, flags);
            }
            if (_hasRequired) CheckRequired(state, flags, start);
            return owner;
        } finally {
            if (rented is not null) ArrayPool<byte>.Shared.Return(rented);
        }
    }

    private object Construct(object?[] values, Span<byte> flags) {
        YamlParameterInfoValues[] parameters = _parameters;
        var arguments = new object?[parameters.Length];
        for (int i = 0; i < parameters.Length; i++) arguments[i] = parameters[i].DefaultValue;
        YamlPropertyInfo[] properties = _properties;
        for (int i = 0; i < properties.Length; i++) {
            if (flags[i] != Unset && properties[i].ParameterIndex >= 0) arguments[properties[i].ParameterIndex] = values[i];
        }
        object owner = _createWithArguments!(arguments)!;
        for (int i = 0; i < properties.Length; i++) {
            if (flags[i] != Unset && properties[i].ParameterIndex < 0) properties[i].SetBoxed(owner, values[i]);
        }
        return owner;
    }

    private void CheckRequired(YamlReadState state, Span<byte> flags, YamlPosition start) {
        for (int i = 0; i < _properties.Length; i++) {
            if (_properties[i].IsRequired && flags[i] == Unset) throw YamlSerializationErrors.MissingRequired(start, state, typeof(T), _properties[i].Name);
        }
    }

    private void ReadMembers(YamlEventCursor cursor, YamlReadState state, object? owner, object?[]? values, Span<byte> flags, ref int expected, bool merging) {
        while (true) {
            if (!cursor.Read()) throw YamlSerializationErrors.UnexpectedEnd(cursor);
            YamlTokenType type = cursor.TokenType;
            if (type == YamlTokenType.MappingEnd) return;
            if (type == YamlTokenType.Alias) {
                cursor.ExpandAlias();
                type = cursor.TokenType;
            }
            int index = -1;
            if (type == YamlTokenType.Scalar) {
                if (cursor.IsMergeKey) {
                    if (!cursor.Read()) throw YamlSerializationErrors.UnexpectedEnd(cursor);
                    ReadMerge(cursor, state, owner, values, flags);
                    continue;
                }
                index = Find(cursor, ref expected);
            } else {
                cursor.Skip();
            }
            if (!cursor.Read()) throw YamlSerializationErrors.UnexpectedEnd(cursor);
            if (index < 0) {
                cursor.Skip();
                continue;
            }
            YamlPropertyInfo property = _properties[index];
            bool settable = property.CanSet || (values is not null && property.ParameterIndex >= 0);
            if (!settable || (merging && flags[index] != Unset)) {
                cursor.Skip();
                continue;
            }
            if (flags[index] == Explicit) throw YamlSerializationErrors.DuplicateKey(cursor, state, property.Name);
            state.PushKey(property.Name);
            if (values is not null) values[index] = property.ReadBoxed(cursor, state);
            else property.ReadAndSet(cursor, state, owner!);
            state.Pop();
            flags[index] = merging ? Merged : Explicit;
        }
    }

    private void ReadMerge(YamlEventCursor cursor, YamlReadState state, object? owner, object?[]? values, Span<byte> flags) {
        int expected = 0;
        if (cursor.TokenType == YamlTokenType.Alias) cursor.ExpandAlias();
        if (cursor.TokenType == YamlTokenType.MappingStart) {
            ReadMembers(cursor, state, owner, values, flags, ref expected, merging: true);
            return;
        }
        if (cursor.TokenType != YamlTokenType.SequenceStart) throw YamlSerializationErrors.TypeMismatch(cursor, state, typeof(T), "a merge key needs a mapping or a sequence of mappings");
        while (true) {
            if (!cursor.Read()) throw YamlSerializationErrors.UnexpectedEnd(cursor);
            if (cursor.TokenType == YamlTokenType.SequenceEnd) return;
            if (cursor.TokenType == YamlTokenType.Alias) cursor.ExpandAlias();
            if (cursor.TokenType != YamlTokenType.MappingStart) throw YamlSerializationErrors.TypeMismatch(cursor, state, typeof(T), "a merge key needs a mapping or a sequence of mappings");
            expected = 0;
            ReadMembers(cursor, state, owner, values, flags, ref expected, merging: true);
        }
    }

    /// <summary>
    /// Finds the property for the current key. Keys usually arrive in declaration order, so the property after the last
    /// match is tried first; the lookup reads the key without allocating where the runtime allows.
    /// </summary>
    private int Find(YamlEventCursor cursor, ref int expected) {
        ReadOnlySpan<char> key = cursor.ScalarSpan;
        YamlPropertyInfo[] properties = _properties;
        if ((uint)expected < (uint)properties.Length && key.SequenceEqual(properties[expected].Name.AsSpan())) return expected++;
        int index;
#if NET9_0_OR_GREATER
        if (!_bySpan.TryGetValue(key, out index)) return -1;
#else
        if (properties.Length <= 8) {
            index = -1;
            for (int i = 0; i < properties.Length; i++) {
                if (key.SequenceEqual(properties[i].Name.AsSpan())) {
                    index = i;
                    break;
                }
            }
            if (index < 0) return -1;
        } else if (!_byName.TryGetValue(cursor.GetString(), out index)) {
            return -1;
        }
#endif
        expected = index + 1;
        return index;
    }
}

/// <summary>
/// A type with <see cref="YamlDerivedTypeAttribute"/> registrations. Values are written with their runtime type's
/// members and its discriminator, as a local tag (<c>!cat</c>) or a first mapping entry. Reading creates only a
/// registered derived type, chosen by its discriminator; without one, the base type itself when it can be created.
/// </summary>
internal sealed class YamlPolymorphicConverter<T> : YamlConverter<T> {
    private readonly YamlTypeInfo<T> _base;
    private readonly YamlDerivedTypeAttribute[] _registrations;
    private readonly string? _propertyName;
    private Dictionary<Type, (YamlTypeInfo Info, string Discriminator)>? _byType;
    private Dictionary<string, YamlTypeInfo>? _byDiscriminator;

    public YamlPolymorphicConverter(YamlTypeInfo<T> baseInfo, YamlDerivedTypeAttribute[] registrations, string? propertyName) {
        _base = baseInfo;
        _registrations = registrations;
        _propertyName = propertyName;
    }

    public override bool IsContainer => true;

    public override bool IsObject => true;

    public override void Configure(YamlSerializerOptions options) {
        var byType = new Dictionary<Type, (YamlTypeInfo, string)>();
        var byDiscriminator = new Dictionary<string, YamlTypeInfo>(StringComparer.Ordinal);
        foreach (YamlDerivedTypeAttribute registration in _registrations) {
            if (!typeof(T).IsAssignableFrom(registration.DerivedType) || registration.DerivedType == typeof(T)) {
                throw new InvalidOperationException("The type " + YamlSerializationErrors.TypeName(registration.DerivedType) + " registered on "
                    + YamlSerializationErrors.TypeName(typeof(T)) + " does not derive from it.");
            }
            YamlTypeInfo derived = options.GetTypeInfo(registration.DerivedType).NonPolymorphic;
            if (!derived.IsObject) {
                throw new InvalidOperationException("The derived type " + YamlSerializationErrors.TypeName(registration.DerivedType) + " must be written as a mapping.");
            }
            if (byDiscriminator.ContainsKey(registration.Discriminator) || byType.ContainsKey(registration.DerivedType)) {
                throw new InvalidOperationException("The type " + YamlSerializationErrors.TypeName(typeof(T)) + " registers the discriminator '"
                    + registration.Discriminator + "' or the type " + YamlSerializationErrors.TypeName(registration.DerivedType) + " more than once.");
            }
            byType.Add(registration.DerivedType, (derived, registration.Discriminator));
            byDiscriminator.Add(registration.Discriminator, derived);
        }
        _byType = byType;
        _byDiscriminator = byDiscriminator;
    }

    private void CheckDiscriminatorProperty(YamlTypeInfo info) {
        if (_propertyName is null) return;
        info.EnsureConfigured();
        if (info.ObjectConverter!.HasProperty(_propertyName)) {
            throw new InvalidOperationException("The type " + YamlSerializationErrors.TypeName(info.Type) + " has a property named '" + _propertyName
                + "', which " + YamlSerializationErrors.TypeName(typeof(T)) + " uses for its discriminator.");
        }
    }

    public override void Write(YamlWriter writer, T value, YamlWriteState state) {
        Type runtime = value!.GetType();
        if (runtime == typeof(T)) {
            _base.WriteCore(writer, value, state);
            return;
        }
        if (!_byType!.TryGetValue(runtime, out (YamlTypeInfo Info, string Discriminator) entry)) {
            throw new NotSupportedException("The runtime type " + YamlSerializationErrors.TypeName(runtime) + " is not registered with [YamlDerivedType] on "
                + YamlSerializationErrors.TypeName(typeof(T)) + ".");
        }
        if (_propertyName is null) {
            writer.WriteTag("!" + entry.Discriminator);
        } else {
            CheckDiscriminatorProperty(entry.Info);
            state.DiscriminatorKey = _propertyName;
            state.DiscriminatorValue = entry.Discriminator;
        }
        entry.Info.WriteCoreBoxed(writer, value, state);
    }

    public override void Collect(T value, YamlReferenceCollector collector) {
        Type runtime = value!.GetType();
        if (runtime == typeof(T)) _base.CollectCore(value, collector);
        else if (_byType!.TryGetValue(runtime, out (YamlTypeInfo Info, string Discriminator) entry)) entry.Info.CollectCoreBoxed(value, collector);
    }

    public override T Read(YamlEventCursor cursor, YamlReadState state) {
        if (_propertyName is null) return ReadTagged(cursor, state);
        if (cursor.TokenType != YamlTokenType.MappingStart) throw YamlSerializationErrors.TypeMismatch(cursor, state, typeof(T));
        return ReadWithProperty(cursor, state);
    }

    private T ReadTagged(YamlEventCursor cursor, YamlReadState state) {
        string? tag = cursor.Tag;
        if (tag is not null && tag.Length > 1 && tag[0] == '!') {
            if (_byDiscriminator!.TryGetValue(tag.Substring(1), out YamlTypeInfo? derived)) return (T)derived.ReadCoreBoxed(cursor, state)!;
            throw YamlSerializationErrors.UnknownDiscriminator(cursor, state, typeof(T), tag);
        }
        if (tag is not null && tag != YamlSchemaResolver.TagPrefix + "map") throw YamlSerializationErrors.UnknownDiscriminator(cursor, state, typeof(T), tag);
        if (!_base.ObjectConverter!.CanCreate) throw YamlSerializationErrors.UnknownDiscriminator(cursor, state, typeof(T), null);
        return _base.ReadCore(cursor, state);
    }

    private T ReadWithProperty(YamlEventCursor cursor, YamlReadState state) {
        YamlPosition start = cursor.Position;
        string propertyName = _propertyName!;
        YamlRecording capture = cursor.BeginCapture();
        string? discriminator = null;
        bool first = true;
        while (true) {
            if (!cursor.Read()) throw YamlSerializationErrors.UnexpectedEnd(cursor);
            if (cursor.TokenType == YamlTokenType.MappingEnd) break;
            bool match = cursor.TokenType == YamlTokenType.Scalar && cursor.ScalarSpan.SequenceEqual(propertyName.AsSpan());
            cursor.Skip();
            if (!cursor.Read()) throw YamlSerializationErrors.UnexpectedEnd(cursor);
            if (match) {
                if (cursor.TokenType == YamlTokenType.Alias) cursor.ExpandAlias();
                if (cursor.TokenType != YamlTokenType.Scalar) throw YamlSerializationErrors.TypeMismatch(cursor, state, typeof(string), "a discriminator must be a scalar");
                discriminator = cursor.GetString();
                if (first) {
                    // The usual case: the discriminator comes first, so the rest is read straight from the input.
                    cursor.EndCapture(capture);
                    return (T)Derived(cursor, state, discriminator).ReadBody(cursor, state, start);
                }
                break;
            }
            cursor.Skip();
            first = false;
        }
        if (discriminator is null) {
            cursor.EndCapture(capture);
            cursor.Replay(capture);
            if (!_base.ObjectConverter!.CanCreate) throw YamlSerializationErrors.UnknownDiscriminator(cursor, state, typeof(T), null);
            return _base.ReadCore(cursor, state);
        }
        // Read past the rest of the mapping into the capture, then read the captured mapping as the derived type.
        int depth = 1;
        while (depth > 0) {
            if (!cursor.Read()) throw YamlSerializationErrors.UnexpectedEnd(cursor);
            if (cursor.TokenType is YamlTokenType.MappingStart or YamlTokenType.SequenceStart) depth++;
            else if (cursor.TokenType is YamlTokenType.MappingEnd or YamlTokenType.SequenceEnd) depth--;
        }
        cursor.EndCapture(capture);
        IYamlObjectConverter derived = Derived(cursor, state, discriminator);
        cursor.Replay(capture);
        return (T)derived.ReadBody(cursor, state, start);
    }

    private IYamlObjectConverter Derived(YamlEventCursor cursor, YamlReadState state, string discriminator) {
        if (!_byDiscriminator!.TryGetValue(discriminator, out YamlTypeInfo? info)) throw YamlSerializationErrors.UnknownDiscriminator(cursor, state, typeof(T), discriminator);
        CheckDiscriminatorProperty(info);
        return info.ObjectConverter!;
    }
}
